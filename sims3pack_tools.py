from __future__ import annotations

import hashlib
import os
import re
import struct
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path, PurePosixPath
from typing import Dict, Iterable, List

_WINDOWS_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_INVALID_WIN_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# CAS Curator is deliberately conservative.  For the first release of
# Sims3Pack extraction support, only CASpart containers are eligible for a
# package-copy extraction.  Everything else stays packed and continues to use
# the existing SIMS3PACK routing path.
_ELIGIBLE_TYPES = {"caspart"}
_PROTECTED_TYPES = {
    "lot",
    "world",
    "sim",
    "household",
    "roomtemplate",
    "pattern",
    "object",
}


def _local_name(tag: str) -> str:
    return str(tag or "").split("}")[-1].casefold()


def _find_child_text(node: ET.Element, name: str) -> str:
    wanted = name.casefold()
    for child in list(node):
        if _local_name(child.tag) == wanted:
            return (child.text or "").strip()
    return ""


def _parse_int(text: str, label: str) -> int:
    raw = str(text or "").strip()
    if not raw:
        raise ValueError(f"Missing {label} in Sims3Pack manifest.")
    try:
        return int(raw, 0)
    except ValueError:
        # Real manifests normally use decimal for offsets/lengths, but accept a
        # bare hexadecimal value only when it contains A-F.
        if re.fullmatch(r"[0-9A-Fa-f]+", raw) and re.search(r"[A-Fa-f]", raw):
            return int(raw, 16)
        raise ValueError(f"Invalid {label} in Sims3Pack manifest: {raw!r}")


def _safe_component(text: str, fallback: str = "Sims3Pack") -> str:
    value = str(text or "").strip().rstrip(". ")
    value = _INVALID_WIN_CHARS.sub("_", value)
    value = re.sub(r"\s+", " ", value).strip().rstrip(". ")
    if not value:
        value = fallback
    stem = value.split(".", 1)[0].upper()
    if stem in _WINDOWS_RESERVED:
        value = f"_{value}"
    return value[:180]


def _safe_manifest_filename(name: str, index: int) -> str:
    raw = str(name or "").replace("\\", "/")
    base = PurePosixPath(raw).name
    base = _safe_component(base, f"embedded_{index}.package")
    if not base.casefold().endswith(".package"):
        base = f"{base}.package"
    return base


def _basic_dbpf_check(blob: bytes) -> None:
    if len(blob) < 96:
        raise ValueError("Embedded package is too small to contain a Sims 3 DBPF header.")
    if blob[:4] != b"DBPF":
        raise ValueError("Embedded package does not begin with the DBPF signature.")
    major, minor = struct.unpack_from("<II", blob, 4)
    if (major, minor) != (2, 0):
        raise ValueError(f"Embedded DBPF is not Sims 3 DBPF 2.0 (found {major}.{minor}).")
    entry_count = struct.unpack_from("<I", blob, 36)[0]
    index_size = struct.unpack_from("<I", blob, 44)[0]
    index_version = struct.unpack_from("<I", blob, 60)[0]
    index_pos = struct.unpack_from("<I", blob, 64)[0]
    if entry_count:
        if not index_size or not index_pos:
            raise ValueError("Embedded DBPF has an incomplete index header.")
        if index_pos + index_size > len(blob):
            raise ValueError("Embedded DBPF index points outside the extracted package bytes.")
    if index_version not in (0, 3):
        # Keep this conservative but allow zero for unusual/empty packages.
        raise ValueError(f"Embedded DBPF has an unexpected index version ({index_version}).")


def inspect_sims3pack(path: Path) -> Dict[str, object]:
    """Read a Sims3Pack manifest and embedded package metadata without writing files."""
    source = Path(path).resolve()
    if source.suffix.casefold() != ".sims3pack":
        raise ValueError("Not a .Sims3Pack file.")
    data = source.read_bytes()
    if len(data) < 17:
        raise ValueError("Sims3Pack is truncated before its header is complete.")

    sig_len = struct.unpack_from("<I", data, 0)[0]
    if sig_len <= 0 or sig_len > 64:
        raise ValueError("Invalid Sims3Pack signature length.")
    pos = 4
    if pos + sig_len + 6 > len(data):
        raise ValueError("Sims3Pack header is truncated.")
    signature = data[pos:pos + sig_len]
    pos += sig_len
    if signature != b"TS3Pack":
        raise ValueError("File does not contain the TS3Pack signature.")
    version = struct.unpack_from("<H", data, pos)[0]
    pos += 2
    xml_len = struct.unpack_from("<I", data, pos)[0]
    pos += 4
    if xml_len <= 0 or pos + xml_len > len(data):
        raise ValueError("Sims3Pack XML manifest length is invalid.")
    xml_bytes = data[pos:pos + xml_len]
    archive_start = pos + xml_len

    try:
        root = ET.fromstring(xml_bytes.decode("utf-8-sig"))
    except (UnicodeDecodeError, ET.ParseError) as exc:
        raise ValueError(f"Sims3Pack manifest could not be read: {exc}") from exc

    if _local_name(root.tag) != "sims3package":
        raise ValueError("Sims3Pack manifest root is not Sims3Package.")

    pack_type = str(root.attrib.get("Type", root.attrib.get("type", "")) or "").strip()
    subtype = str(root.attrib.get("SubType", root.attrib.get("subtype", "")) or "").strip()
    display_name = _find_child_text(root, "DisplayName")
    description = _find_child_text(root, "Description")

    packaged_files: List[Dict[str, object]] = []
    package_entries: List[Dict[str, object]] = []
    for idx, node in enumerate(root.iter(), start=1):
        if _local_name(node.tag) != "packagedfile":
            continue
        name = _find_child_text(node, "Name")
        length = _parse_int(_find_child_text(node, "Length"), "Length")
        offset = _parse_int(_find_child_text(node, "Offset"), "Offset")
        content_type = _find_child_text(node, "ContentType")
        crc = _find_child_text(node, "Crc")
        guid = _find_child_text(node, "Guid")
        if length < 0 or offset < 0:
            raise ValueError("Sims3Pack manifest contains a negative Offset/Length.")
        start = archive_start + offset
        end = start + length
        if start < archive_start or end > len(data):
            raise ValueError(f"Embedded file points outside the Sims3Pack: {name or f'entry {idx}'}. ")
        blob = data[start:end]
        magic = "DBPF" if blob[:4] == b"DBPF" else ("PNG" if blob[:8] == b"\x89PNG\r\n\x1a\n" else "OTHER")
        entry: Dict[str, object] = {
            "index": idx,
            "name": name,
            "length": length,
            "offset": offset,
            "content_type": content_type,
            "crc": crc,
            "guid": guid,
            "magic": magic,
        }
        if magic == "DBPF":
            _basic_dbpf_check(blob)
            entry["sha256"] = hashlib.sha256(blob).hexdigest()
            entry["output_name"] = _safe_manifest_filename(name, idx)
            package_entries.append(entry)
        packaged_files.append(entry)

    if not packaged_files:
        raise ValueError("Sims3Pack manifest contains no PackagedFile entries.")

    type_key = pack_type.casefold()
    if type_key in _ELIGIBLE_TYPES:
        status = "eligible"
        reason = "CASpart Sims3Pack"
    elif type_key in _PROTECTED_TYPES:
        status = "protected"
        reason = f"protected type: {pack_type or 'unknown'}"
    else:
        status = "protected"
        reason = f"unknown type: {pack_type or '(blank)'}"

    return {
        "path": source,
        "version": version,
        "type": pack_type,
        "subtype": subtype,
        "display_name": display_name,
        "description": description,
        "status": status,
        "reason": reason,
        "packaged_files": packaged_files,
        "package_entries": package_entries,
        "package_count": len(package_entries),
        "file_size": len(data),
        "archive_start": archive_start,
    }


def read_embedded_package_bytes(plan: Dict[str, object], entry: Dict[str, object]) -> bytes:
    source = Path(plan["path"])
    data = source.read_bytes()
    start = int(plan["archive_start"]) + int(entry["offset"])
    end = start + int(entry["length"])
    blob = data[start:end]
    _basic_dbpf_check(blob)
    digest = hashlib.sha256(blob).hexdigest()
    if digest != str(entry.get("sha256", "")):
        raise ValueError("Sims3Pack changed after analysis; analyse it again before extracting.")
    return blob


def package_hashes(paths: Iterable[Path]) -> Dict[str, List[Path]]:
    result: Dict[str, List[Path]] = {}
    for raw in paths:
        path = Path(raw)
        h = hashlib.sha256()
        with path.open("rb") as fh:
            for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                h.update(chunk)
        result.setdefault(h.hexdigest(), []).append(path.resolve())
    return result


def annotate_package_alternatives(plan: Dict[str, object], existing_hashes: Dict[str, List[Path]]) -> Dict[str, object]:
    matched = 0
    missing = 0
    for entry in plan.get("package_entries", []):
        matches = existing_hashes.get(str(entry.get("sha256", "")), [])
        entry["standalone_matches"] = [str(p) for p in matches]
        if matches:
            matched += 1
        else:
            missing += 1
    plan["matched_package_count"] = matched
    plan["missing_package_count"] = missing
    return plan


def _friendly_output_paths(plan: Dict[str, object], folder: Path, entries: List[Dict[str, object]]) -> List[Path]:
    """Build readable extracted-package names from the original Sims3Pack name.

    A Sims3Pack with one embedded package becomes ``<Sims3Pack name>.package``.
    When several embedded packages exist, their stable package-order number is
    appended as ``__1``, ``__2`` and so on.  Package identity/duplicate
    detection continues to use SHA-256; the friendly filename is presentation
    only.
    """
    source = Path(plan["path"])
    base_stem = _safe_component(source.stem, "Sims3Pack")
    all_entries = list(plan.get("package_entries", []))
    order_by_manifest_index = {
        int(entry.get("index", pos)): pos
        for pos, entry in enumerate(all_entries, start=1)
    }

    claimed = set()
    outputs: List[Path] = []
    for fallback_pos, entry in enumerate(entries, start=1):
        if len(all_entries) <= 1:
            base = f"{base_stem}.package"
        else:
            package_pos = order_by_manifest_index.get(int(entry.get("index", fallback_pos)), fallback_pos)
            base = f"{base_stem}__{package_pos}.package"

        stem = Path(base).stem
        suffix = Path(base).suffix or ".package"
        candidate = folder / base
        n = 2
        while candidate.name.casefold() in claimed:
            candidate = folder / f"{stem}_{n}{suffix}"
            n += 1
        claimed.add(candidate.name.casefold())
        outputs.append(candidate)
    return outputs


def extract_missing_cas_packages(plan: Dict[str, object]) -> List[Path]:
    if str(plan.get("status")) != "eligible":
        raise ValueError("This Sims3Pack is protected and cannot be extracted by CAS Curator.")

    missing_entries = [
        entry for entry in plan.get("package_entries", [])
        if not entry.get("standalone_matches")
    ]
    if not missing_entries:
        return []

    source = Path(plan["path"])
    outer = source.parent / _safe_component(source.stem, "Sims3Pack")
    outputs = _friendly_output_paths(plan, outer, missing_entries)

    # Preflight before creating anything.
    for out in outputs:
        if out.exists():
            raise FileExistsError(f"Extraction destination already exists:\n{out}")

    created: List[Path] = []
    try:
        outer.mkdir(parents=True, exist_ok=True)
        for entry, out in zip(missing_entries, outputs):
            blob = read_embedded_package_bytes(plan, entry)
            fd, temp_name = tempfile.mkstemp(prefix=out.name + ".", suffix=".part", dir=str(outer))
            try:
                with os.fdopen(fd, "wb") as fh:
                    fh.write(blob)
                    fh.flush()
                    os.fsync(fh.fileno())
                os.replace(temp_name, out)
            except Exception:
                try:
                    os.unlink(temp_name)
                except OSError:
                    pass
                raise
            created.append(out.resolve())
        return created
    except Exception:
        for path in reversed(created):
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass
        try:
            if outer.exists() and not any(outer.iterdir()):
                outer.rmdir()
        except OSError:
            pass
        raise


def planned_missing_outputs(plan: Dict[str, object]) -> List[Path]:
    """Return the exact output paths for unmatched eligible embedded packages."""
    if str(plan.get("status")) != "eligible":
        return []
    missing_entries = [
        entry for entry in plan.get("package_entries", [])
        if not entry.get("standalone_matches")
    ]
    if not missing_entries:
        return []
    source = Path(plan["path"])
    outer = source.parent / _safe_component(source.stem, "Sims3Pack")
    return _friendly_output_paths(plan, outer, missing_entries)


def preflight_missing_cas_batch(plans: Iterable[Dict[str, object]]) -> List[Dict[str, object]]:
    """Validate all CAS extraction destinations before any Sims3Pack writes occur."""
    selected: List[Dict[str, object]] = []
    claimed: Dict[str, Path] = {}
    for plan in plans:
        outputs = planned_missing_outputs(plan)
        if not outputs:
            continue
        for out in outputs:
            key = str(out).casefold()
            if key in claimed:
                raise FileExistsError(
                    "Two Sims3Packs would extract to the same destination:\n"
                    f"{out}\n\nOther source:\n{claimed[key]}"
                )
            claimed[key] = Path(plan["path"])
            if out.exists():
                raise FileExistsError(f"Extraction destination already exists:\n{out}")
        selected.append(plan)
    return selected


def extract_missing_cas_batch(plans: Iterable[Dict[str, object]]) -> List[Path]:
    """Extract a preflighted batch and roll back files created by this batch on failure."""
    selected = preflight_missing_cas_batch(plans)
    created: List[Path] = []
    created_dirs = set()
    try:
        for plan in selected:
            outputs = planned_missing_outputs(plan)
            for out in outputs:
                created_dirs.add(out.parent)
            new_files = extract_missing_cas_packages(plan)
            created.extend(new_files)
        return created
    except Exception:
        for path in reversed(created):
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass
        for folder in sorted(created_dirs, key=lambda p: len(str(p)), reverse=True):
            try:
                if folder.exists() and not any(folder.iterdir()):
                    folder.rmdir()
            except OSError:
                pass
        raise
