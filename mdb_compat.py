from __future__ import annotations

import os
import re
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional


_PATTERNFILL_RE = re.compile(
    rb'(<patternFill\b[^>]*>)(.*?)(</patternFill>)',
    re.IGNORECASE | re.DOTALL,
)
_EXTLST_RE = re.compile(
    rb'<extLst\b[^>]*>.*?</extLst>',
    re.IGNORECASE | re.DOTALL,
)
_DATA_VALIDATION_TAG_RE = re.compile(rb'<dataValidation\b[^>]*>', re.IGNORECASE)
_VIEW_ATTR_RE = re.compile(rb'\s+view="[^"]*"', re.IGNORECASE)
_VENDOR_URI_TOKEN = b'uri="smNativeData"'


@dataclass(frozen=True)
class MDBCompatibilityReport:
    path: Path
    patternfill_extlst_count: int = 0
    data_validation_view_count: int = 0
    vendor_extlst_count: int = 0
    styles_xml_present: bool = False
    zip_ok: bool = False
    error: str = ""

    @property
    def known_blocker_count(self) -> int:
        return self.patternfill_extlst_count + self.data_validation_view_count

    @property
    def needs_compatible_copy(self) -> bool:
        return self.known_blocker_count > 0


@dataclass(frozen=True)
class MDBCompatibilityCopyResult:
    source_path: Path
    compatible_path: Path
    removed_vendor_extlst: int
    removed_data_validation_view: int


def _count_patternfill_extlst(styles_xml: bytes) -> int:
    count = 0
    for match in _PATTERNFILL_RE.finditer(styles_xml):
        count += len(_EXTLST_RE.findall(match.group(2)))
    return count


def _count_data_validation_view(xml_bytes: bytes) -> int:
    count = 0
    for match in _DATA_VALIDATION_TAG_RE.finditer(xml_bytes):
        count += len(_VIEW_ATTR_RE.findall(match.group(0)))
    return count


def _count_vendor_extlst(xml_bytes: bytes) -> int:
    return sum(1 for match in _EXTLST_RE.finditer(xml_bytes) if _VENDOR_URI_TOKEN in match.group(0))


def inspect_mdb_compatibility(path: Path) -> MDBCompatibilityReport:
    path = Path(path).resolve()
    if not path.is_file():
        return MDBCompatibilityReport(path=path, error="Workbook file does not exist.")

    patternfill_count = 0
    validation_view_count = 0
    vendor_extlst_count = 0
    styles_present = False
    try:
        with zipfile.ZipFile(path, "r") as archive:
            bad_member = archive.testzip()
            if bad_member:
                return MDBCompatibilityReport(
                    path=path,
                    zip_ok=False,
                    error=f"Workbook ZIP integrity check failed at {bad_member}.",
                )
            for info in archive.infolist():
                name = info.filename
                if not name.endswith(".xml"):
                    continue
                if not name.startswith("xl/"):
                    continue
                data = archive.read(name)
                vendor_extlst_count += _count_vendor_extlst(data)
                if name == "xl/styles.xml":
                    styles_present = True
                    patternfill_count = _count_patternfill_extlst(data)
                elif name.startswith("xl/worksheets/"):
                    validation_view_count += _count_data_validation_view(data)
    except zipfile.BadZipFile as exc:
        return MDBCompatibilityReport(path=path, error=f"Workbook is not a readable XLSX ZIP container: {exc}")
    except Exception as exc:
        return MDBCompatibilityReport(path=path, error=f"Could not inspect workbook: {exc}")

    return MDBCompatibilityReport(
        path=path,
        patternfill_extlst_count=patternfill_count,
        data_validation_view_count=validation_view_count,
        vendor_extlst_count=vendor_extlst_count,
        styles_xml_present=styles_present,
        zip_ok=True,
    )


def _remove_vendor_extlst(xml_bytes: bytes) -> tuple[bytes, int]:
    removed = 0

    def replace_extlst(match: re.Match[bytes]) -> bytes:
        nonlocal removed
        block = match.group(0)
        if _VENDOR_URI_TOKEN in block:
            removed += 1
            return b""
        return block

    return _EXTLST_RE.sub(replace_extlst, xml_bytes), removed


def _clean_worksheet_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    removed = 0

    def replace_validation_tag(match: re.Match[bytes]) -> bytes:
        nonlocal removed
        new_tag, count = _VIEW_ATTR_RE.subn(b"", match.group(0))
        removed += count
        return new_tag

    return _DATA_VALIDATION_TAG_RE.sub(replace_validation_tag, xml_bytes), removed


def recommended_compatible_copy_path(source: Path) -> Path:
    source = Path(source).resolve()
    base = source.with_name(f"{source.stem}_CAS_Curator_Compatible{source.suffix}")
    if not base.exists():
        return base
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return source.with_name(f"{source.stem}_CAS_Curator_Compatible_{stamp}{source.suffix}")


def create_compatible_copy(source: Path, destination: Optional[Path] = None) -> MDBCompatibilityCopyResult:
    """Create a separate scanner-compatible workbook, leaving source untouched.

    The compatible copy removes spreadsheet-app-specific ``smNativeData`` extLst
    blocks from OOXML parts because openpyxl cannot preserve them when it saves a
    workbook anyway. It also removes the non-standard ``view`` attribute from
    dataValidation rules. Standard workbook data is left in place and the copy
    must pass ZIP integrity checks before it is returned. The caller should then
    verify it with the exact openpyxl runtime used by the frozen scanner.
    """
    source = Path(source).resolve()
    report = inspect_mdb_compatibility(source)
    if not report.zip_ok:
        raise RuntimeError(report.error or "Workbook ZIP integrity check failed.")
    if not report.needs_compatible_copy:
        raise RuntimeError("No known scanner-blocking compatibility entries were detected.")

    destination = Path(destination).resolve() if destination else recommended_compatible_copy_path(source)
    if destination == source:
        raise RuntimeError("The compatible copy must use a different filename from the original MDB.")
    if destination.exists():
        raise RuntimeError(f"Compatible-copy destination already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)

    removed_vendor_extlst = 0
    removed_validation_view = 0
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{destination.stem}_tmp_",
        suffix=destination.suffix,
        dir=str(destination.parent),
    )
    os.close(fd)
    Path(temp_name).unlink(missing_ok=True)
    temp_path = Path(temp_name)

    try:
        with zipfile.ZipFile(source, "r") as src, zipfile.ZipFile(temp_path, "w") as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename.startswith("xl/") and info.filename.endswith(".xml"):
                    data, count = _remove_vendor_extlst(data)
                    removed_vendor_extlst += count
                    if info.filename.startswith("xl/worksheets/"):
                        data, count = _clean_worksheet_xml(data)
                        removed_validation_view += count
                dst.writestr(info, data)

        if removed_vendor_extlst != report.vendor_extlst_count:
            raise RuntimeError(
                f"Vendor-extension cleanup count mismatch: expected {report.vendor_extlst_count}, removed {removed_vendor_extlst}."
            )
        if removed_validation_view != report.data_validation_view_count:
            raise RuntimeError(
                f"Data-validation cleanup count mismatch: expected {report.data_validation_view_count}, removed {removed_validation_view}."
            )

        with zipfile.ZipFile(temp_path, "r") as archive:
            bad_member = archive.testzip()
            if bad_member:
                raise RuntimeError(f"Compatible-copy ZIP verification failed at {bad_member}.")

        temp_path.replace(destination)
        return MDBCompatibilityCopyResult(
            source_path=source,
            compatible_path=destination,
            removed_vendor_extlst=removed_vendor_extlst,
            removed_data_validation_view=removed_validation_view,
        )
    except Exception:
        try:
            temp_path.unlink(missing_ok=True)
        except Exception:
            pass
        try:
            destination.unlink(missing_ok=True)
        except Exception:
            pass
        raise
