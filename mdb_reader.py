from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple, Union
from urllib.parse import unquote, urlparse
from xml.sax.saxutils import escape
import hashlib
import os
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
import zipfile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN_NS, "r": REL_NS}

EDITABLE_FIELDS = ("Image", "Status", "Creator", "Converter", "File Location", "CC Magic Profiles", "Merge Profiles", "Notes")


def _normalise_cc_magic_profiles(value: str) -> str:
    text = str(value or "").replace("\r", "\n")
    raw_parts = []
    for line in text.split("\n"):
        raw_parts.extend(line.split(";"))
    result = []
    seen = set()
    for part in raw_parts:
        clean = part.strip()
        if not clean:
            continue
        key = clean.casefold()
        if key in seen:
            continue
        seen.add(key)
        result.append(clean)
    return "; ".join(result)


class WorkbookChangedError(RuntimeError):
    """Raised when the MDB changed on disk after it was loaded."""


@dataclass
class SaveResult:
    backup_path: Path
    changed_rows: int
    changed_fields: int
    source_sha256: str
    source_size: int
    source_mtime_ns: int
    deleted_rows: int = 0


def _col_num(cell_ref: str) -> int:
    m = re.match(r"([A-Z]+)", cell_ref or "")
    if not m:
        return 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n


def _col_letters(n: int) -> str:
    if n < 1:
        raise ValueError("Column number must be >= 1")
    out = []
    while n:
        n, rem = divmod(n - 1, 26)
        out.append(chr(65 + rem))
    return "".join(reversed(out))


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _cell_text(cell: ET.Element, shared_strings: List[str]) -> str:
    kind = cell.attrib.get("t")
    if kind == "inlineStr":
        is_node = cell.find(f"{{{MAIN_NS}}}is")
        if is_node is None:
            return ""
        return "".join((n.text or "") for n in is_node.iter(f"{{{MAIN_NS}}}t"))

    value = cell.find(f"{{{MAIN_NS}}}v")
    if value is None or value.text is None:
        return ""
    if kind == "s":
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return value.text
    return value.text


def _load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    name = "xl/sharedStrings.xml"
    if name not in zf.namelist():
        return []
    root = ET.fromstring(zf.read(name))
    result: List[str] = []
    for si in root.findall(f"{{{MAIN_NS}}}si"):
        result.append("".join((n.text or "") for n in si.iter(f"{{{MAIN_NS}}}t")))
    return result


def _sheet_paths(zf: zipfile.ZipFile) -> Dict[str, str]:
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rels = {r.attrib["Id"]: r.attrib["Target"] for r in rel_root}
    out: Dict[str, str] = {}
    for sheet in wb_root.find(f"{{{MAIN_NS}}}sheets") or []:
        name = sheet.attrib["name"]
        rid = sheet.attrib[f"{{{REL_NS}}}id"]
        target = rels[rid].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        out[name] = target
    return out


def _read_sheet_rows(
    zf: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: List[str],
) -> Iterable[Tuple[int, Dict[int, str]]]:
    # Streaming read keeps memory use low even when Packages has 10k+ rows.
    with zf.open(sheet_path) as stream:
        for event, elem in ET.iterparse(stream, events=("end",)):
            if elem.tag != f"{{{MAIN_NS}}}row":
                continue
            row_num = int(elem.attrib.get("r", "0") or 0)
            values: Dict[int, str] = {}
            for cell in elem.findall(f"{{{MAIN_NS}}}c"):
                n = _col_num(cell.attrib.get("r", ""))
                if n:
                    values[n] = _cell_text(cell, shared_strings)
            yield row_num, values
            elem.clear()


def _load_casp_observations(
    zf: zipfile.ZipFile,
    sheet_path: Optional[str],
    shared_strings: List[str],
) -> Tuple[Dict[str, List[Dict[str, str]]], Dict[str, List[Dict[str, str]]]]:
    """Read direct CASP Master rows, indexed by package *and* CASP instance.

    Package filename remains the normal link.  The CASP-instance index is a
    direct-identity fallback for older/routed MDBs where the package label may
    no longer match even though CC Master still holds the exact CASP IDs.
    These observations are read-only and never alter scanner/workbook output.
    """
    if not sheet_path:
        return {}, {}
    headers_by_col: Dict[int, str] = {}
    by_package: Dict[str, List[Dict[str, str]]] = {}
    by_instance: Dict[str, List[Dict[str, str]]] = {}
    for row_num, values in _read_sheet_rows(zf, sheet_path, shared_strings):
        if row_num == 1:
            headers_by_col = {n: v for n, v in values.items() if v}
            continue
        package_col = next((n for n, h in headers_by_col.items() if h == "Package"), None)
        if not package_col:
            continue
        package = values.get(package_col, "").strip()
        if not package:
            continue
        row: Dict[str, str] = {}
        for col_num, header in headers_by_col.items():
            value = values.get(col_num, "")
            if value:
                row[header] = value
        by_package.setdefault(package.casefold(), []).append(row)
        instance = str(row.get("CASP Instance", "") or "").strip().casefold()
        if instance:
            by_instance.setdefault(instance, []).append(row)
    return by_package, by_instance


def _friendly_flag_values(text: str, label: str) -> List[str]:
    """Extract friendly CASP flag names such as Teen/Adult or Male/Female."""
    match = re.search(
        rf"(?mi)^\s*{re.escape(label)}\s*:\s*0x[0-9a-f]+\s*\(([^)]*)\)",
        str(text or ""),
    )
    if not match:
        return []
    values: List[str] = []
    seen = set()
    for raw in match.group(1).split(","):
        value = raw.strip()
        if not value or value.casefold() in {"none", "0"}:
            continue
        key = value.casefold()
        if key in seen:
            continue
        seen.add(key)
        values.append(value)
    return values




def _friendly_ea_resource_status(value: str) -> str:
    """Return a plain-language display for legacy/experimental EA overlap results.

    This is display-only metadata. It does not change the stored Content Type,
    scanner logic, or Routing decisions.
    """
    raw = str(value or "").strip()
    key = raw.casefold()
    if key in {"custom", "custom content"}:
        return "No EA overlap detected"
    if key == "ea override":
        return "Likely EA override / default replacement"
    if key == "ea candidate":
        return "EA resource overlap — review"
    if key == "ea match":
        return "Matches EA resources"
    if key in {"unsupported", "not assessed"}:
        return "Not assessed"
    return raw

_AGE_FLAG_ORDER = [
    (0x01, "Baby"),
    (0x02, "Toddler"),
    (0x04, "Child"),
    (0x08, "Teen"),
    (0x10, "YoungAdult"),
    (0x20, "Adult"),
    (0x40, "Elder"),
]
_GENDER_FLAG_ORDER = [
    (0x01, "Male"),
    (0x02, "Female"),
]


def _friendly_age_gender_values(text: str) -> Tuple[List[str], List[str]]:
    """Read friendly Age/Gender names from s3pi's direct AgeGender text.

    Most CASPs render labelled lines such as ``Age: 0x30 (YoungAdult, Adult)``
    and ``Gender: 0x02 (Female)``.  Keep that as the primary path, then fall
    back to the labelled hexadecimal flags if a wrapper/version omits names.
    """
    raw = str(text or "")
    ages = _friendly_flag_values(raw, "Age")
    genders = _friendly_flag_values(raw, "Gender")

    if not ages:
        match = re.search(r"(?mi)^\s*Age\s*:\s*0x([0-9a-f]+)", raw)
        if match:
            value = int(match.group(1), 16)
            ages = [name for bit, name in _AGE_FLAG_ORDER if value & bit]
    if not genders:
        match = re.search(r"(?mi)^\s*Gender\s*:\s*0x([0-9a-f]+)", raw)
        if match:
            value = int(match.group(1), 16)
            genders = [name for bit, name in _GENDER_FLAG_ORDER if value & bit]

    return ages, genders


def _aggregate_casp_age_gender(observations: List[Dict[str, str]]) -> Tuple[str, str]:
    """Combine direct Age/Gender flags across every CASP row for a package."""
    ages: List[str] = []
    genders: List[str] = []
    seen_ages = set()
    seen_genders = set()
    for observation in observations:
        direct_ages, direct_genders = _friendly_age_gender_values(
            observation.get("Age Gender", "")
        )
        for value in direct_ages:
            key = value.casefold()
            if key not in seen_ages:
                seen_ages.add(key)
                ages.append(value)
        for value in direct_genders:
            key = value.casefold()
            if key not in seen_genders:
                seen_genders.add(key)
                genders.append(value)
    return ", ".join(ages), ", ".join(genders)


def _casp_instance_ids(text: str) -> List[str]:
    """Extract normalised CASP IDs from a CC Master multi-value cell."""
    seen = set()
    result: List[str] = []
    for match in re.findall(r"0x[0-9a-fA-F]{1,16}", str(text or "")):
        key = match.casefold()
        if key in seen:
            continue
        seen.add(key)
        result.append(key)
    return result


def _linked_casp_observations(
    record: Dict[str, str],
    by_package: Dict[str, List[Dict[str, str]]],
    by_instance: Dict[str, List[Dict[str, str]]],
) -> List[Dict[str, str]]:
    """Return direct CASP rows for a package, with CASP-ID fallback linking."""
    package_key = str(record.get("Package", "") or "").strip().casefold()
    observations = list(by_package.get(package_key, []))
    if observations:
        return observations

    # Filename links can become stale after older organisation/routing work.
    # CASP IDs are direct resource identities and are safe as a fallback.
    linked: List[Dict[str, str]] = []
    seen = set()
    for instance in _casp_instance_ids(record.get("CASP Instance IDs", "")):
        for observation in by_instance.get(instance, []):
            identity = (
                str(observation.get("Package", "") or "").casefold(),
                str(observation.get("CASP Instance", "") or "").casefold(),
            )
            if identity in seen:
                continue
            seen.add(identity)
            linked.append(observation)
    return linked


_OUTFIT_CATEGORY_ORDER = [
    "Everyday",
    "Formal",
    "Sleepwear",
    "Swimwear",
    "Athletic",
    "Career",
    "Outerwear",
]
_OUTFIT_CATEGORY_ALIASES = {
    "everyday": "Everyday",
    "formal": "Formal",
    "formalwear": "Formal",
    "sleepwear": "Sleepwear",
    "swimwear": "Swimwear",
    "athletic": "Athletic",
    "career": "Career",
    "outerwear": "Outerwear",
    # Older s3pi builds exposed the Seasons Outerwear bit by its historical
    # placeholder name.  Treat that exact flag as Outerwear for display only.
    "unknown12": "Outerwear",
}


def _friendly_outfit_categories(text: str) -> List[str]:
    """Return the seven user-facing outfit categories from CASP ClothingCategory.

    CASP Master stores the direct ``str(ClothingCategory)`` observation.  Common
    s3pi output looks like ``0x0020013F (Naked, Everyday, Formalwear, ...)``.
    Ignore technical/special-purpose flags such as Naked, ValidForRandom,
    ValidForMaternity and IsHiddenInCAS; this summary is deliberately limited
    to the seven normal outfit filters users recognise in CAS.
    """
    raw = str(text or "").strip()
    if not raw:
        return []

    groups = re.findall(r"\(([^()]*)\)", raw)
    candidate_text = ",".join(groups) if groups else raw
    found = set()
    for token in re.split(r"[,;|]", candidate_text):
        clean = token.strip()
        if not clean:
            continue
        # Defensive cleanup for alternate representations such as
        # ``Everyday = True`` without depending on them being present.
        clean = re.split(r"\s*=\s*", clean, maxsplit=1)[0].strip()
        friendly = _OUTFIT_CATEGORY_ALIASES.get(clean.casefold())
        if friendly:
            found.add(friendly)
    return [name for name in _OUTFIT_CATEGORY_ORDER if name in found]


def _aggregate_casp_outfit_categories(observations: List[Dict[str, str]]) -> str:
    """Combine direct ClothingCategory flags across every CASP in a package."""
    found = set()
    for observation in observations:
        found.update(_friendly_outfit_categories(observation.get("Category", "")))
    return ", ".join(name for name in _OUTFIT_CATEGORY_ORDER if name in found)


def _hyperlinks_for_sheet(zf: zipfile.ZipFile, sheet_path: str) -> Dict[str, str]:
    rel_path = str(Path(sheet_path).parent / "_rels" / (Path(sheet_path).name + ".rels")).replace("\\", "/")
    if rel_path not in zf.namelist():
        return {}
    rel_root = ET.fromstring(zf.read(rel_path))
    rels = {r.attrib["Id"]: r.attrib.get("Target", "") for r in rel_root}
    root = ET.fromstring(zf.read(sheet_path))
    links = root.find(f"{{{MAIN_NS}}}hyperlinks")
    if links is None:
        return {}
    out: Dict[str, str] = {}
    for h in links:
        rid = h.attrib.get(f"{{{REL_NS}}}id")
        ref = h.attrib.get("ref")
        if rid and ref and rid in rels:
            out[ref] = rels[rid]
    return out


def target_to_local_path(target: str) -> Optional[Path]:
    if not target:
        return None
    if target.lower().startswith("file:"):
        parsed = urlparse(target)
        raw = unquote(parsed.path or "")
        if re.match(r"^/[A-Za-z]:/", raw):
            raw = raw[1:]
        return Path(raw.replace("/", "\\"))
    return Path(unquote(target))


_THUMBNAIL_EXTENSIONS = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp")


def _portable_thumbnail_target(mdb_path: Path, target: str, package_hash: str) -> str:
    """Resolve a broken scanner thumbnail link beside a moved MDB.

    Scanner-created thumbnails are named from the package SHA-256 and live in a
    ``Thumbnails`` folder beside the MDB. Manual preview images are deliberately
    left alone: portability fallback is attempted only when the linked filename
    itself is the package hash.
    """
    if not target:
        return target

    original = target_to_local_path(target)
    if original is not None and original.exists():
        return target

    hash_text = str(package_hash or "").strip().lower()
    if not re.fullmatch(r"[0-9a-f]{64}", hash_text):
        return target

    original_name = original.name if original is not None else ""
    original_suffix = Path(original_name).suffix.lower()
    original_stem = Path(original_name).stem.lower()
    if original_stem != hash_text or original_suffix not in _THUMBNAIL_EXTENSIONS:
        return target

    thumbnail_dir = Path(mdb_path).parent / "Thumbnails"
    if not thumbnail_dir.is_dir():
        return target

    # Prefer the exact old filename first, then accept another supported image
    # extension with the same scanner hash name.
    candidates = [thumbnail_dir / original_name]
    candidates.extend(thumbnail_dir / f"{hash_text}{ext}" for ext in _THUMBNAIL_EXTENSIONS)
    seen = set()
    for candidate in candidates:
        key = str(candidate).casefold()
        if key in seen:
            continue
        seen.add(key)
        if candidate.is_file():
            return str(candidate.resolve())

    return target


@dataclass
class MDBData:
    path: Path
    headers: List[str]
    records: List[Dict[str, str]]
    casp_observations_by_package: Dict[str, List[Dict[str, str]]]
    status_options: List[str]
    library_options: List[str]
    profile_options: List[str]
    file_locations: List[Tuple[str, str]]
    merge_profile_options: List[str]
    profile_layouts: List[Tuple[str, str, str, str]]
    thumbnail_count: int
    casp_count: int
    package_history_count: int
    source_sha256: str
    source_size: int
    source_mtime_ns: int


class MDBWorkbook:
    """Reader + guarded writer for Sims3_CC_Master_Database.xlsx.

    Reading is performed directly from OOXML for speed. Writing is deliberately
    restricted to the manual-review fields: Image, Status, File Location, Profiles and Notes.
    Legacy Library Location values are read-only and preserved. Saves create a timestamped full-workbook backup, verify the source has
    not changed since load, patch only the CC Master worksheet cells, validate the
    temporary workbook, then atomically replace the original file.
    """

    CC_SHEET = "CC Master Database"
    SETTINGS_SHEET = "Settings"
    PACKAGES_SHEET = "Packages"
    CASP_SHEET = "CASP Master"

    def load(self, path: Union[str, Path]) -> MDBData:
        path = Path(path)
        if path.suffix.lower() != ".xlsx":
            raise ValueError("Please select the .xlsx Master Database file.")
        if not path.exists():
            raise FileNotFoundError(path)

        signature = _sha256_file(path)
        stat = path.stat()

        with zipfile.ZipFile(path) as zf:
            shared = _load_shared_strings(zf)
            sheets = _sheet_paths(zf)
            if self.CC_SHEET not in sheets:
                raise ValueError("This workbook does not contain 'CC Master Database'.")

            cc_path = sheets[self.CC_SHEET]
            hyperlinks = _hyperlinks_for_sheet(zf, cc_path)

            headers_by_col: Dict[int, str] = {}
            records: List[Dict[str, str]] = []
            thumbnail_count = 0

            for row_num, values in _read_sheet_rows(zf, cc_path, shared):
                if row_num == 1:
                    headers_by_col = {n: v for n, v in values.items() if v}
                    continue
                package = values.get(1, "").strip()
                # Crucial: Excel contains formatted/empty trailing rows.
                if not package:
                    continue

                record: Dict[str, str] = {"Package": package, "_ExcelRow": str(row_num)}
                for col_num, header in headers_by_col.items():
                    if header == "Package":
                        continue
                    value = values.get(col_num, "")
                    if value:
                        record[header] = value

                image_col = next((col for col, header in headers_by_col.items() if header == "Image"), None)
                image_ref = f"{_col_letters(image_col)}{row_num}" if image_col else ""
                image_value = str(record.get("Image", "") or "").strip()
                looks_like_image_path = (
                    image_value.lower().startswith("file:")
                    or Path(unquote(image_value)).suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}
                )
                target = image_value if looks_like_image_path else (hyperlinks.get(image_ref, "") if image_ref else "")
                if target:
                    target = _portable_thumbnail_target(path, target, record.get("Package Hash", ""))
                    record["_ThumbnailTarget"] = target
                    record["Thumbnail"] = "Yes"
                    thumbnail_count += 1
                else:
                    record["Thumbnail"] = "No"
                records.append(record)

            headers = [headers_by_col[n] for n in sorted(headers_by_col)]
            # CC Magic Profiles is a library-owned manual field. Older MDBs may not
            # have the column yet; expose it virtually so the GUI can edit it, and
            # create the physical worksheet column safely on the first save.
            # Creator is a library-owned manual field. Older MDBs expose it
            # virtually and get a physical worksheet column only on first save.
            if "Creator" not in headers:
                headers.append("Creator")
            if "Converter" not in headers:
                headers.append("Converter")
            if "CC Magic Profiles" not in headers:
                headers.append("CC Magic Profiles")
            if "File Location" not in headers:
                headers.append("File Location")
            if "Merge Profiles" not in headers:
                headers.append("Merge Profiles")

            # CASP Master already stores direct Age/Gender and ClothingCategory
            # flags per CASP. Expose Age/Gender as normal read-only package metadata
            # and Outfit Category as a provisional virtual field for Experimental
            # validation. None of these values are written back to CC Master.
            casp_observations_by_package, casp_observations_by_instance = _load_casp_observations(
                zf, sheets.get(self.CASP_SHEET), shared
            )
            if self.CASP_SHEET in sheets:
                if "Age" not in headers:
                    headers.append("Age")
                if "Gender" not in headers:
                    headers.append("Gender")
                if "Outfit Category" not in headers:
                    headers.append("Outfit Category")
            for record in records:
                observations = _linked_casp_observations(
                    record, casp_observations_by_package, casp_observations_by_instance
                )
                age, gender = _aggregate_casp_age_gender(observations)
                outfit_category = _aggregate_casp_outfit_categories(observations)
                if age:
                    record["Age"] = age
                if gender:
                    record["Gender"] = gender
                if outfit_category:
                    record["Outfit Category"] = outfit_category

            # Content Type is an older technical scanner field. Keep the raw
            # workbook value untouched, but expose a friendlier virtual summary
            # for the opt-in Experimental UI.
            if "Content Type" in headers and "EA Resource Status" not in headers:
                headers.append("EA Resource Status")
            for record in records:
                ea_status = _friendly_ea_resource_status(record.get("Content Type", ""))
                if ea_status:
                    record["EA Resource Status"] = ea_status

            status_options, library_options, profile_options, file_locations, merge_profile_options, profile_layouts = self._load_settings(zf, sheets, shared)
            # Older workbooks may not yet have a persisted profile registry.
            # Always include profile names already assigned to packages so the UI
            # can manage them, then persist the registry on the next profile save.
            existing_profiles = []
            seen_profiles = set()
            for rec in records:
                for profile in _normalise_cc_magic_profiles(rec.get("CC Magic Profiles", "")).split("; "):
                    profile = profile.strip()
                    if not profile:
                        continue
                    key = profile.casefold()
                    if key in seen_profiles:
                        continue
                    seen_profiles.add(key)
                    existing_profiles.append(profile)
            for profile in profile_options:
                seen_profiles.add(profile.casefold())
            profile_options.extend(p for p in existing_profiles if p.casefold() not in {x.casefold() for x in profile_options})
            profile_options = sorted(profile_options, key=str.casefold)
            # Merge Profiles are a separate managed registry. Preserve names already
            # assigned in existing workbooks even if the registry is missing.
            existing_merge_profiles = []
            seen_merge_profiles = {p.casefold() for p in merge_profile_options}
            for rec in records:
                for profile in _normalise_cc_magic_profiles(rec.get("Merge Profiles", "")).split("; "):
                    profile = profile.strip()
                    if not profile or profile.casefold() in seen_merge_profiles:
                        continue
                    seen_merge_profiles.add(profile.casefold())
                    existing_merge_profiles.append(profile)
            merge_profile_options.extend(existing_merge_profiles)
            merge_profile_options = sorted(merge_profile_options, key=str.casefold)
            package_history_count = self._count_real_rows(zf, sheets.get(self.PACKAGES_SHEET), shared)
            casp_count = self._count_real_rows(zf, sheets.get(self.CASP_SHEET), shared)

        return MDBData(
            path=path,
            headers=headers,
            records=records,
            casp_observations_by_package=casp_observations_by_package,
            status_options=status_options,
            library_options=library_options,
            profile_options=profile_options,
            file_locations=file_locations,
            merge_profile_options=merge_profile_options,
            profile_layouts=profile_layouts,
            thumbnail_count=thumbnail_count,
            casp_count=casp_count,
            package_history_count=package_history_count,
            source_sha256=signature,
            source_size=stat.st_size,
            source_mtime_ns=stat.st_mtime_ns,
        )

    def save_manual_changes(
        self,
        mdb: MDBData,
        changes: Dict[int, Dict[str, str]],
        profile_options: Optional[List[str]] = None,
        status_options: Optional[List[str]] = None,
        delete_rows: Optional[Iterable[int]] = None,
        file_locations: Optional[List[Tuple[str, str]]] = None,
        merge_profile_options: Optional[List[str]] = None,
        profile_layouts: Optional[List[Tuple[str, str, str, str]]] = None,
    ) -> SaveResult:
        """Safely write manual Library fields, registry changes, and staged record deletions."""
        clean: Dict[int, Dict[str, str]] = {}
        delete_row_nums = {int(value) for value in (delete_rows or [])}
        if any(value <= 1 for value in delete_row_nums):
            raise ValueError("Only package data rows can be deleted; the header row is protected.")
        known_record_rows = {int(rec.get("_ExcelRow", "0") or 0) for rec in mdb.records}
        unknown_delete_rows = delete_row_nums - known_record_rows
        if unknown_delete_rows:
            raise ValueError(
                "Deletion was refused because row(s) are not package records in the loaded workbook: "
                + ", ".join(str(value) for value in sorted(unknown_delete_rows)[:20])
            )
        changed_fields = 0
        for raw_row, row_changes in changes.items():
            row_num = int(raw_row)
            allowed: Dict[str, str] = {}
            for field, value in row_changes.items():
                if field not in EDITABLE_FIELDS:
                    raise ValueError(f"'{field}' is not an editable MDB field.")
                clean_value = "" if value is None else str(value)
                if field in ("CC Magic Profiles", "Merge Profiles"):
                    clean_value = _normalise_cc_magic_profiles(clean_value)
                allowed[field] = clean_value
            if allowed:
                clean[row_num] = allowed
                changed_fields += len(allowed)

        clean_profiles: Optional[List[str]] = None
        if profile_options is not None:
            seen = set()
            clean_profiles = []
            for raw in profile_options:
                name = str(raw or "").strip()
                if not name:
                    continue
                key = name.casefold()
                if key in seen:
                    continue
                seen.add(key)
                clean_profiles.append(name)
            clean_profiles.sort(key=str.casefold)

        clean_status_options: Optional[List[str]] = None
        if status_options is not None:
            seen_statuses = set()
            clean_status_options = []
            for raw in status_options:
                value = str(raw or "").strip()
                if not value or value.casefold() in seen_statuses:
                    continue
                seen_statuses.add(value.casefold())
                clean_status_options.append(value)

        clean_merge_profiles: Optional[List[str]] = None
        if merge_profile_options is not None:
            seen_merge = set()
            clean_merge_profiles = []
            for raw in merge_profile_options:
                name = str(raw or "").strip()
                if not name:
                    continue
                key = name.casefold()
                if key in seen_merge:
                    continue
                seen_merge.add(key)
                clean_merge_profiles.append(name)
            clean_merge_profiles.sort(key=str.casefold)

        clean_file_locations: Optional[List[Tuple[str, str]]] = None
        if file_locations is not None:
            clean_file_locations = []
            seen_location_names = set()
            for raw_location in file_locations:
                if not isinstance(raw_location, (tuple, list)) or len(raw_location) < 2:
                    raise ValueError("Each File Location entry must contain a friendly name and folder path.")
                name = str(raw_location[0] or "").strip()
                folder_path = str(raw_location[1] or "").strip()
                if not name or not folder_path:
                    raise ValueError("File Location names and folder paths cannot be blank.")
                key = name.casefold()
                if key in seen_location_names:
                    raise ValueError(f"Duplicate File Location name: {name}")
                seen_location_names.add(key)
                clean_file_locations.append((name, folder_path))
            clean_file_locations.sort(key=lambda item: item[0].casefold())

        clean_profile_layouts: Optional[List[Tuple[str, str, str, str]]] = None
        if profile_layouts is not None:
            clean_profile_layouts = []
            seen_layout_keys = set()
            for raw_layout in profile_layouts:
                if not isinstance(raw_layout, (tuple, list)) or len(raw_layout) < 4:
                    raise ValueError(
                        "Each Profile Layout entry must contain Profile, Package Hash, Package and File Location."
                    )
                profile = str(raw_layout[0] or "").strip()
                package_hash = str(raw_layout[1] or "").strip().lower()
                package_name = Path(str(raw_layout[2] or "").strip()).name
                location_name = str(raw_layout[3] or "").strip()
                if not profile or not location_name or (not package_hash and not package_name):
                    raise ValueError(
                        "Profile Layout entries require a Profile, File Location and package identity."
                    )
                identity = ("h:" + package_hash) if package_hash else ("p:" + package_name.casefold())
                key = (profile.casefold(), identity)
                if key in seen_layout_keys:
                    raise ValueError(
                        f"Duplicate Profile Layout entry for '{profile}' / '{package_name or package_hash}'."
                    )
                seen_layout_keys.add(key)
                clean_profile_layouts.append((profile, package_hash, package_name, location_name))
            clean_profile_layouts.sort(
                key=lambda item: (item[0].casefold(), item[2].casefold(), item[1], item[3].casefold())
            )

        if set(clean).intersection(delete_row_nums):
            raise ValueError("A record cannot have pending field edits and deletion in the same save.")
        if (
            not clean and not delete_row_nums and clean_profiles is None
            and clean_status_options is None and clean_file_locations is None
            and clean_merge_profiles is None and clean_profile_layouts is None
        ):
            raise ValueError("There are no changes to save.")

        path = Path(mdb.path)
        if not path.exists():
            raise FileNotFoundError(path)

        current_hash = _sha256_file(path)
        if current_hash != mdb.source_sha256:
            raise WorkbookChangedError(
                "The Master Database changed on disk after it was opened in the library. "
                "Reload it before saving so external changes are not overwritten."
            )

        backup_dir = path.parent / "MDB_Backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
        backup_path = backup_dir / f"{path.stem}_backup_{stamp}{path.suffix}"
        shutil.copy2(path, backup_path)

        fd, tmp_name = tempfile.mkstemp(
            prefix=f".{path.stem}_editing_",
            suffix=".xlsx",
            dir=str(path.parent),
        )
        os.close(fd)
        tmp_path = Path(tmp_name)

        try:
            with zipfile.ZipFile(path, "r") as zin:
                shared = _load_shared_strings(zin)
                sheets = _sheet_paths(zin)
                cc_path = sheets.get(self.CC_SHEET)
                if not cc_path:
                    raise ValueError("This workbook does not contain 'CC Master Database'.")

                patched_xml = None
                patched_sheet_rels_xml = None
                sheet_rels_path = str(Path(cc_path).parent / "_rels" / (Path(cc_path).name + ".rels")).replace("\\", "/")
                if clean or delete_row_nums:
                    headers_by_col: Dict[int, str] = {}
                    for row_num, values in _read_sheet_rows(zin, cc_path, shared):
                        if row_num == 1:
                            headers_by_col = {n: v for n, v in values.items() if v}
                            break
                    header_to_col = {header: col for col, header in headers_by_col.items()}
                    sheet_xml = zin.read(cc_path).decode("utf-8")

                    if clean:
                        needed_fields = {field for row_changes in clean.values() for field in row_changes}
                        virtual_fields = {"Creator", "Converter", "CC Magic Profiles", "File Location", "Merge Profiles"}
                        missing_required = [
                            field for field in needed_fields
                            if field not in header_to_col and field not in virtual_fields
                        ]
                        if missing_required:
                            raise ValueError(
                                "Editable column(s) missing from CC Master Database: "
                                + ", ".join(sorted(missing_required))
                            )

                        # Older workbooks may lack library-owned fields. Add only the fields
                        # actually being saved, preserving every existing column and value.
                        for virtual_field in ("Creator", "Converter", "CC Magic Profiles", "File Location", "Merge Profiles"):
                            if virtual_field in needed_fields and virtual_field not in header_to_col:
                                next_col = max(header_to_col.values(), default=0) + 1
                                header_to_col[virtual_field] = next_col
                                sheet_xml, header_rows = self._patch_sheet_xml(
                                    sheet_xml,
                                    {1: {virtual_field: virtual_field}},
                                    header_to_col,
                                    {next_col: None},
                                )
                                if 1 not in header_rows:
                                    raise ValueError(f"Could not add the {virtual_field} storage column to the MDB header.")
                                sheet_xml = self._expand_dimension_for_column(sheet_xml, next_col)

                        column_styles = self._find_column_styles(sheet_xml, header_to_col, needed_fields)
                        patched_xml, patched_rows = self._patch_sheet_xml(
                            sheet_xml, clean, header_to_col, column_styles,
                        )
                        if patched_rows != set(clean):
                            missing_rows = sorted(set(clean) - patched_rows)
                            raise ValueError(
                                "Could not locate CC Master Database row(s): "
                                + ", ".join(str(n) for n in missing_rows[:20])
                            )
                    else:
                        patched_xml = sheet_xml

                    if delete_row_nums:
                        patched_xml, cleared_rows = self._clear_record_rows(patched_xml, delete_row_nums)
                        if cleared_rows != delete_row_nums:
                            missing_rows = sorted(delete_row_nums - cleared_rows)
                            raise ValueError(
                                "Could not locate record row(s) to delete: "
                                + ", ".join(str(n) for n in missing_rows[:20])
                            )
                        rels_xml = zin.read(sheet_rels_path).decode("utf-8") if sheet_rels_path in zin.namelist() else ""
                        patched_xml, cleaned_rels_xml = self._remove_row_hyperlinks(
                            patched_xml, rels_xml, delete_row_nums
                        )
                        if sheet_rels_path in zin.namelist():
                            patched_sheet_rels_xml = cleaned_rels_xml

                settings_path = sheets.get(self.SETTINGS_SHEET)
                patched_settings_xml = None
                if (
                    clean_profiles is not None or clean_status_options is not None
                    or clean_file_locations is not None or clean_merge_profiles is not None
                    or clean_profile_layouts is not None
                ):
                    if not settings_path:
                        raise ValueError("This workbook does not contain a Settings sheet for its registries.")
                    patched_settings_xml = zin.read(settings_path).decode("utf-8")
                    if clean_status_options is not None:
                        patched_settings_xml = self._patch_settings_status_xml(
                            patched_settings_xml, clean_status_options
                        )
                    if clean_profiles is not None:
                        patched_settings_xml = self._patch_settings_profiles_xml(
                            patched_settings_xml, clean_profiles
                        )
                    if clean_merge_profiles is not None:
                        # Column F is reserved for the Merge Profiles registry only when
                        # unused or already carrying our header.
                        settings_rows_for_merge = _read_sheet_rows(zin, settings_path, shared)
                        merge_header = {}
                        occupied_merge_col = False
                        for row_num, values in settings_rows_for_merge:
                            if row_num == 1:
                                merge_header = {col: value.strip() for col, value in values.items()}
                            elif str(values.get(6, "") or "").strip():
                                occupied_merge_col = True
                        existing_merge_header = merge_header.get(6, "")
                        if existing_merge_header and existing_merge_header != "Merge Profiles":
                            raise ValueError(
                                f"Settings column 6 is already used by '{existing_merge_header}'. "
                                "The Merge Profiles registry will not overwrite it."
                            )
                        if occupied_merge_col and existing_merge_header != "Merge Profiles":
                            raise ValueError(
                                "Settings column 6 contains existing data without a Merge Profiles registry header. "
                                "The Merge Profiles registry will not overwrite it."
                            )
                        patched_settings_xml = self._patch_settings_merge_profiles_xml(
                            patched_settings_xml, clean_merge_profiles
                        )
                    if clean_file_locations is not None:
                        # D/E are reserved for the File Location registry only when they
                        # are unused or already carry our headers. Never overwrite unknown
                        # Settings-sheet data from a future/third-party workbook format.
                        settings_rows = _read_sheet_rows(zin, settings_path, shared)
                        settings_header = {}
                        occupied_location_cols = set()
                        for row_num, values in settings_rows:
                            if row_num == 1:
                                settings_header = {col: value.strip() for col, value in values.items()}
                            else:
                                occupied_location_cols.update(
                                    col for col in (4, 5) if str(values.get(col, "") or "").strip()
                                )
                        expected_location_headers = {4: "File Location Name", 5: "File Location Path"}
                        for col in (4, 5):
                            existing_header = settings_header.get(col, "")
                            if existing_header and existing_header != expected_location_headers[col]:
                                raise ValueError(
                                    f"Settings column {col} is already used by '{existing_header}'. "
                                    "The File Location registry will not overwrite it."
                                )
                            if col in occupied_location_cols and existing_header != expected_location_headers[col]:
                                raise ValueError(
                                    f"Settings column {col} contains existing data without a File Location registry header. "
                                    "The File Location registry will not overwrite it."
                                )
                        patched_settings_xml = self._patch_settings_file_locations_xml(
                            patched_settings_xml, clean_file_locations
                        )
                    if clean_profile_layouts is not None:
                        # G:J are reserved for Experimental Profile Layouts only when
                        # unused or already carrying our exact headers. This keeps the
                        # writer conservative around future/third-party Settings data.
                        settings_rows = _read_sheet_rows(zin, settings_path, shared)
                        settings_header = {}
                        occupied_layout_cols = set()
                        for row_num, values in settings_rows:
                            if row_num == 1:
                                settings_header = {col: value.strip() for col, value in values.items()}
                            else:
                                occupied_layout_cols.update(
                                    col for col in (7, 8, 9, 10)
                                    if str(values.get(col, "") or "").strip()
                                )
                        expected_layout_headers = {
                            7: "Profile Layout Profile",
                            8: "Profile Layout Package Hash",
                            9: "Profile Layout Package",
                            10: "Profile Layout File Location",
                        }
                        for col in (7, 8, 9, 10):
                            existing_header = settings_header.get(col, "")
                            if existing_header and existing_header != expected_layout_headers[col]:
                                raise ValueError(
                                    f"Settings column {col} is already used by '{existing_header}'. "
                                    "Experimental Profile Layouts will not overwrite it."
                                )
                            if col in occupied_layout_cols and existing_header != expected_layout_headers[col]:
                                raise ValueError(
                                    f"Settings column {col} contains existing data without a Profile Layout header. "
                                    "Experimental Profile Layouts will not overwrite it."
                                )
                        patched_settings_xml = self._patch_settings_profile_layouts_xml(
                            patched_settings_xml, clean_profile_layouts
                        )

                with zipfile.ZipFile(tmp_path, "w") as zout:
                    for info in zin.infolist():
                        if patched_xml is not None and info.filename == cc_path:
                            data = patched_xml.encode("utf-8")
                        elif patched_sheet_rels_xml is not None and info.filename == sheet_rels_path:
                            data = patched_sheet_rels_xml.encode("utf-8")
                        elif patched_settings_xml is not None and info.filename == settings_path:
                            data = patched_settings_xml.encode("utf-8")
                        else:
                            data = zin.read(info.filename)
                        zout.writestr(info, data)

            # Verify the new XLSX is a healthy ZIP and the requested values can be read back.
            with zipfile.ZipFile(tmp_path, "r") as ztest:
                bad_member = ztest.testzip()
                if bad_member:
                    raise ValueError(f"Workbook validation failed at ZIP member: {bad_member}")
                if clean:
                    self._verify_saved_values(ztest, clean)
                if delete_row_nums:
                    self._verify_deleted_rows(ztest, delete_row_nums)
                if clean_profiles is not None:
                    self._verify_profile_registry(ztest, clean_profiles)
                if clean_status_options is not None:
                    self._verify_status_registry(ztest, clean_status_options)
                if clean_file_locations is not None:
                    self._verify_file_locations(ztest, clean_file_locations)
                if clean_merge_profiles is not None:
                    self._verify_merge_profile_registry(ztest, clean_merge_profiles)
                if clean_profile_layouts is not None:
                    self._verify_profile_layouts(ztest, clean_profile_layouts)

            try:
                os.replace(tmp_path, path)
            except PermissionError as exc:
                raise PermissionError(
                    "Windows would not let the library replace the Master Database. "
                    "If the workbook is open in Excel, close Excel and try Save Changes again."
                ) from exc

            new_stat = path.stat()
            new_hash = _sha256_file(path)
            mdb.source_sha256 = new_hash
            mdb.source_size = new_stat.st_size
            mdb.source_mtime_ns = new_stat.st_mtime_ns
            if clean_profiles is not None:
                mdb.profile_options = list(clean_profiles)
            if clean_status_options is not None:
                mdb.status_options = list(clean_status_options)
            if clean_file_locations is not None:
                mdb.file_locations = list(clean_file_locations)
            if clean_merge_profiles is not None:
                mdb.merge_profile_options = list(clean_merge_profiles)
            if clean_profile_layouts is not None:
                mdb.profile_layouts = list(clean_profile_layouts)

            return SaveResult(
                backup_path=backup_path,
                changed_rows=len(clean),
                changed_fields=changed_fields,
                source_sha256=new_hash,
                source_size=new_stat.st_size,
                source_mtime_ns=new_stat.st_mtime_ns,
                deleted_rows=len(delete_row_nums),
            )
        except Exception:
            try:
                tmp_path.unlink(missing_ok=True)
            except Exception:
                pass
            raise
        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except Exception:
                    pass

    def _clear_record_rows(self, xml_text: str, rows_to_clear: set[int]) -> Tuple[str, set[int]]:
        """Clear every cell value in selected package rows while preserving worksheet row/style structure."""
        row_pattern = re.compile(r'(<row\b[^>]*\br="(\d+)"[^>]*>)(.*?)(</row>)', re.DOTALL)
        cleared: set[int] = set()

        def clear_row(match: re.Match) -> str:
            row_num = int(match.group(2))
            if row_num not in rows_to_clear:
                return match.group(0)
            body = match.group(3)
            # Keep cell references and formatting, but remove values, formulas and inline strings.
            # Match self-closing cells first. Otherwise a pattern for <c>...</c> can
            # start at <c .../> and consume through the next populated cell's </c>.
            cell_pattern = re.compile(r'<c\b([^>]*)/>|<c\b([^>]*)>.*?</c>', re.DOTALL)

            def empty_cell(cell_match: re.Match) -> str:
                attrs = cell_match.group(1) if cell_match.group(1) is not None else cell_match.group(2)
                # For self-closing cells the regex capture can include the closing "/".
                # Strip it before rebuilding the tag or we would emit invalid XML like <c ...//>.
                attrs = re.sub(r"\s*/\s*$", "", attrs)
                attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
                attrs = re.sub(r"\s+t='[^']*'", "", attrs)
                return f'<c{attrs}/>'

            body = cell_pattern.sub(empty_cell, body)
            cleared.add(row_num)
            return match.group(1) + body + match.group(4)

        return row_pattern.sub(clear_row, xml_text), cleared

    def _remove_row_hyperlinks(
        self, sheet_xml: str, rels_xml: str, rows_to_clear: set[int]
    ) -> Tuple[str, str]:
        """Remove hyperlink elements/relationships attached to cleared record rows only."""
        hyperlink_pattern = re.compile(r'<hyperlink\b[^>]*/>|<hyperlink\b[^>]*>.*?</hyperlink>', re.DOTALL)
        removed_ids: set[str] = set()

        def keep_or_remove(match: re.Match) -> str:
            tag = match.group(0)
            ref_match = re.search(r'\bref="([^"]+)"', tag)
            if not ref_match:
                return tag
            ref = ref_match.group(1)
            cell_refs = ref.split(":")
            row_numbers = []
            for cell_ref in cell_refs:
                cell_match = re.fullmatch(r'[A-Z]+(\d+)', cell_ref)
                if cell_match:
                    row_numbers.append(int(cell_match.group(1)))
            if not any(row_num in rows_to_clear for row_num in row_numbers):
                return tag
            id_match = re.search(r'\br:id="([^"]+)"', tag)
            if id_match:
                removed_ids.add(id_match.group(1))
            return ""

        sheet_xml = hyperlink_pattern.sub(keep_or_remove, sheet_xml)
        if not removed_ids:
            return sheet_xml, rels_xml

        rel_pattern = re.compile(r'<Relationship\b[^>]*/>|<Relationship\b[^>]*>.*?</Relationship>', re.DOTALL)
        def keep_relationship(match: re.Match) -> str:
            id_match = re.search(r'\bId="([^"]+)"', match.group(0))
            return "" if id_match and id_match.group(1) in removed_ids else match.group(0)
        rels_xml = rel_pattern.sub(keep_relationship, rels_xml)
        return sheet_xml, rels_xml

    def _verify_deleted_rows(self, zf: zipfile.ZipFile, rows_to_clear: set[int]) -> None:
        """Verify deleted package rows no longer contain a Package value or any cell data."""
        shared = _load_shared_strings(zf)
        sheets = _sheet_paths(zf)
        cc_path = sheets[self.CC_SHEET]
        headers_by_col: Dict[int, str] = {}
        seen: Dict[int, Dict[int, str]] = {}
        for row_num, values in _read_sheet_rows(zf, cc_path, shared):
            if row_num == 1:
                headers_by_col = {n: v for n, v in values.items() if v}
                continue
            if row_num in rows_to_clear:
                seen[row_num] = values
        missing = rows_to_clear - set(seen)
        if missing:
            raise ValueError("Deletion verification could not find row(s): " + ", ".join(map(str, sorted(missing))))
        package_col = next((col for col, header in headers_by_col.items() if header == "Package"), 1)
        for row_num in rows_to_clear:
            values = seen[row_num]
            if any(str(value or "").strip() for value in values.values()):
                raise ValueError(f"Deletion verification found remaining cell data in row {row_num}.")
            if values.get(package_col, "").strip():
                raise ValueError(f"Deletion verification found a remaining Package value in row {row_num}.")

    def _patch_sheet_xml(
        self,
        xml_text: str,
        changes: Dict[int, Dict[str, str]],
        header_to_col: Dict[str, int],
        column_styles: Dict[int, Optional[str]],
    ) -> Tuple[str, set]:
        # Accept both default-namespace worksheet XML (<row>/<c>) and files
        # serialized with a namespace prefix such as <x:row>/<x:c>.
        row_pattern = re.compile(
            r'(<(?:[A-Za-z_][\w.-]*:)?row\b[^>]*\br="(\d+)"[^>]*>)(.*?)(</(?:[A-Za-z_][\w.-]*:)?row>)',
            re.DOTALL,
        )
        patched_rows: set = set()

        def patch_row(match: re.Match) -> str:
            row_num = int(match.group(2))
            updates = changes.get(row_num)
            if not updates:
                return match.group(0)
            body = match.group(3)
            for field, value in updates.items():
                col_num = header_to_col[field]
                body = self._patch_cell_in_row(
                    body,
                    row_num,
                    col_num,
                    value,
                    column_styles.get(col_num),
                )
            patched_rows.add(row_num)
            return match.group(1) + body + match.group(4)

        return row_pattern.sub(patch_row, xml_text), patched_rows

    def _patch_cell_in_row(
        self,
        body: str,
        row_num: int,
        col_num: int,
        value: str,
        fallback_style: Optional[str],
    ) -> str:
        ref = f"{_col_letters(col_num)}{row_num}"
        ref_escaped = re.escape(ref)
        # Match both <c> and prefixed forms such as <x:c>. Preserve the source
        # prefix when rewriting an existing cell or inserting a new one.
        cell_pattern = re.compile(
            rf'<(?P<prefix>[A-Za-z_][\w.-]*:)?c\b([^>]*\br="{ref_escaped}"[^>]*)/>'
            rf'|<(?P<prefix2>[A-Za-z_][\w.-]*:)?c\b([^>]*\br="{ref_escaped}"[^>]*)>(.*?)</(?:[A-Za-z_][\w.-]*:)?c>',
            re.DOTALL,
        )
        m = cell_pattern.search(body)

        escaped = escape(value)
        if m:
            prefix = m.group("prefix") if m.group("prefix") is not None else (m.group("prefix2") or "")
            attrs = m.group(2) if m.group(2) is not None else m.group(4)
            attrs = re.sub(r"\s*/\s*$", "", attrs)
            attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
            attrs = re.sub(r"\s+t='[^']*'", "", attrs)
            text_node = f'<{prefix}is><{prefix}t xml:space="preserve">{escaped}</{prefix}t></{prefix}is>'
            replacement = f'<{prefix}c{attrs} t="inlineStr">{text_node}</{prefix}c>'
            return body[: m.start()] + replacement + body[m.end() :]

        prefix_match = re.search(r'<([A-Za-z_][\w.-]*:)?c\b', body)
        prefix = (prefix_match.group(1) or "") if prefix_match else ""
        text_node = f'<{prefix}is><{prefix}t xml:space="preserve">{escaped}</{prefix}t></{prefix}is>'
        style_attr = f' s="{fallback_style}"' if fallback_style else ""
        new_cell = f'<{prefix}c r="{ref}"{style_attr} t="inlineStr">{text_node}</{prefix}c>'

        starts = list(re.finditer(r'<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*\br="([A-Z]+)\d+"[^>]*>', body))
        for cell_start in starts:
            if _col_num(cell_start.group(1)) > col_num:
                pos = cell_start.start()
                return body[:pos] + new_cell + body[pos:]
        return body + new_cell

    def _expand_dimension_for_column(self, sheet_xml: str, col_num: int) -> str:
        """Expand worksheet dimension metadata if a newly-added manual column lies beyond it."""
        pattern = re.compile(r'<dimension\s+ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*/>')
        m = pattern.search(sheet_xml)
        if not m:
            return sheet_xml
        start_col, start_row, end_col, end_row = m.groups()
        if _col_num(end_col) >= col_num:
            return sheet_xml
        new_ref = f'{start_col}{start_row}:{_col_letters(col_num)}{end_row}'
        return sheet_xml[:m.start()] + f'<dimension ref="{new_ref}"/>' + sheet_xml[m.end():]

    def _find_column_styles(
        self,
        sheet_xml: str,
        header_to_col: Dict[str, int],
        fields=None,
    ) -> Dict[int, Optional[str]]:
        result: Dict[int, Optional[str]] = {}
        for field in (fields or EDITABLE_FIELDS):
            if field not in header_to_col:
                continue
            col_num = header_to_col[field]
            letters = _col_letters(col_num)
            pattern = re.compile(
                rf'<c\b(?=[^>]*\br="{letters}\d+")(?=[^>]*\bs="(\d+)")[^>]*>',
                re.DOTALL,
            )
            m = pattern.search(sheet_xml)
            result[col_num] = m.group(1) if m else None
        return result

    def _verify_saved_values(self, zf: zipfile.ZipFile, changes: Dict[int, Dict[str, str]]) -> None:
        shared = _load_shared_strings(zf)
        sheets = _sheet_paths(zf)
        cc_path = sheets[self.CC_SHEET]
        headers_by_col: Dict[int, str] = {}
        target_rows = set(changes)
        seen: Dict[int, Dict[str, str]] = {}

        for row_num, values in _read_sheet_rows(zf, cc_path, shared):
            if row_num == 1:
                headers_by_col = {n: v for n, v in values.items() if v}
                continue
            if row_num not in target_rows:
                continue
            row_values: Dict[str, str] = {}
            for col, header in headers_by_col.items():
                if header in EDITABLE_FIELDS:
                    row_values[header] = values.get(col, "")
            seen[row_num] = row_values
            if len(seen) == len(target_rows):
                break

        for row_num, expected_fields in changes.items():
            actual = seen.get(row_num)
            if actual is None:
                raise ValueError(f"Save verification could not read row {row_num}.")
            for field, expected in expected_fields.items():
                if actual.get(field, "") != expected:
                    raise ValueError(
                        f"Save verification failed for row {row_num}, field '{field}'."
                    )

    def _load_settings(
        self, zf, sheets, shared
    ) -> Tuple[
        List[str], List[str], List[str], List[Tuple[str, str]], List[str],
        List[Tuple[str, str, str, str]],
    ]:
        path = sheets.get(self.SETTINGS_SHEET)
        if not path:
            return [], [], [], [], [], []
        status: List[str] = []
        library: List[str] = []
        profiles: List[str] = []
        file_locations: List[Tuple[str, str]] = []
        merge_profiles: List[str] = []
        profile_layouts: List[Tuple[str, str, str, str]] = []
        seen_profiles = set()
        seen_merge_profiles = set()
        seen_location_names = set()
        has_file_location_headers = False
        has_merge_profile_header = False
        has_profile_layout_headers = False
        seen_profile_layouts = set()
        for row_num, values in _read_sheet_rows(zf, path, shared):
            if row_num == 1:
                has_file_location_headers = (
                    values.get(4, "").strip() == "File Location Name"
                    and values.get(5, "").strip() == "File Location Path"
                )
                has_merge_profile_header = values.get(6, "").strip() == "Merge Profiles"
                has_profile_layout_headers = (
                    values.get(7, "").strip() == "Profile Layout Profile"
                    and values.get(8, "").strip() == "Profile Layout Package Hash"
                    and values.get(9, "").strip() == "Profile Layout Package"
                    and values.get(10, "").strip() == "Profile Layout File Location"
                )
                continue
            a = values.get(1, "").strip()
            b = values.get(2, "").strip()
            c = values.get(3, "").strip()
            location_name = values.get(4, "").strip() if has_file_location_headers else ""
            location_path = values.get(5, "").strip() if has_file_location_headers else ""
            merge_profile = values.get(6, "").strip() if has_merge_profile_header else ""
            layout_profile = values.get(7, "").strip() if has_profile_layout_headers else ""
            layout_hash = values.get(8, "").strip().lower() if has_profile_layout_headers else ""
            layout_package = Path(values.get(9, "").strip()).name if has_profile_layout_headers else ""
            layout_location = values.get(10, "").strip() if has_profile_layout_headers else ""
            if a:
                status.append(a)
            if b:
                library.append(b)
            if c and c.casefold() not in seen_profiles:
                seen_profiles.add(c.casefold())
                profiles.append(c)
            if location_name and location_path and location_name.casefold() not in seen_location_names:
                seen_location_names.add(location_name.casefold())
                file_locations.append((location_name, location_path))
            if row_num > 1 and merge_profile and merge_profile.casefold() not in seen_merge_profiles:
                seen_merge_profiles.add(merge_profile.casefold())
                merge_profiles.append(merge_profile)
            if layout_profile and layout_location and (layout_hash or layout_package):
                identity = ("h:" + layout_hash) if layout_hash else ("p:" + layout_package.casefold())
                layout_key = (layout_profile.casefold(), identity)
                if layout_key not in seen_profile_layouts:
                    seen_profile_layouts.add(layout_key)
                    profile_layouts.append(
                        (layout_profile, layout_hash, layout_package, layout_location)
                    )
        file_locations.sort(key=lambda item: item[0].casefold())
        merge_profiles.sort(key=str.casefold)
        profile_layouts.sort(
            key=lambda item: (item[0].casefold(), item[2].casefold(), item[1], item[3].casefold())
        )
        return status, library, profiles, file_locations, merge_profiles, profile_layouts

    def _patch_settings_status_xml(self, xml_text: str, statuses: List[str]) -> str:
        """Replace only the Status registry in Settings column A, preserving other columns."""
        ET.register_namespace("", MAIN_NS)
        root = ET.fromstring(xml_text)
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise ValueError("Settings sheet has no sheetData element.")

        rows = {}
        for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
            try:
                rows[int(row.attrib.get("r", "0") or 0)] = row
            except ValueError:
                continue

        def ensure_row(row_num: int):
            row = rows.get(row_num)
            if row is not None:
                return row
            row = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(row_num)})
            inserted = False
            for idx, existing in enumerate(list(sheet_data)):
                try:
                    existing_num = int(existing.attrib.get("r", "0") or 0)
                except ValueError:
                    existing_num = 0
                if existing_num > row_num:
                    sheet_data.insert(idx, row)
                    inserted = True
                    break
            if not inserted:
                sheet_data.append(row)
            rows[row_num] = row
            return row

        def set_col_a(row_num: int, value: str):
            row = ensure_row(row_num)
            target_ref = f"A{row_num}"
            old = None
            style = None
            for cell in list(row):
                ref = cell.attrib.get("r", "")
                if ref == target_ref:
                    old = cell
                    style = cell.attrib.get("s")
                    break
            if old is not None:
                row.remove(old)
            if not value:
                return
            attrs = {"r": target_ref, "t": "inlineStr"}
            if style:
                attrs["s"] = style
            cell = ET.Element(f"{{{MAIN_NS}}}c", attrs)
            is_node = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
            t = ET.SubElement(is_node, f"{{{MAIN_NS}}}t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
            t.text = value
            row.append(cell)
            cells = list(row)
            cells.sort(key=lambda c: _col_num(c.attrib.get("r", "")))
            row[:] = cells

        set_col_a(1, "Status")
        max_existing = max(rows.keys(), default=1)
        max_needed = max(max_existing, len(statuses) + 1)
        for row_num in range(2, max_needed + 1):
            value = statuses[row_num - 2] if row_num - 2 < len(statuses) else ""
            set_col_a(row_num, value)

        dim = root.find(f"{{{MAIN_NS}}}dimension")
        if dim is not None:
            current_ref = dim.attrib.get("ref", "A1:C1")
            end = current_ref.split(":")[-1]
            match = re.match(r"([A-Z]+)", end)
            last_col = match.group(1) if match else "C"
            last_row = max(max_existing, len(statuses) + 1)
            dim.set("ref", f"A1:{last_col}{last_row}")
        return ET.tostring(root, encoding="unicode")

    def _verify_status_registry(self, zf: zipfile.ZipFile, expected: List[str]) -> None:
        shared = _load_shared_strings(zf)
        sheets = _sheet_paths(zf)
        path = sheets.get(self.SETTINGS_SHEET)
        if not path:
            raise ValueError("Save verification could not find the Settings sheet.")
        actual = []
        for row_num, values in _read_sheet_rows(zf, path, shared):
            if row_num <= 1:
                continue
            value = values.get(1, "").strip()
            if value:
                actual.append(value)
        if actual != expected:
            raise ValueError("Save verification failed for the Status registry.")

    def _patch_settings_profiles_xml(self, xml_text: str, profiles: List[str]) -> str:
        """Persist the generic Profiles registry in Settings column C."""
        ET.register_namespace("", MAIN_NS)
        root = ET.fromstring(xml_text)
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise ValueError("Settings sheet has no sheetData element.")

        rows = {}
        for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
            try:
                rows[int(row.attrib.get("r", "0") or 0)] = row
            except ValueError:
                continue

        def ensure_row(row_num: int):
            row = rows.get(row_num)
            if row is not None:
                return row
            row = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(row_num)})
            # Insert rows in numeric order.
            inserted = False
            for idx, existing in enumerate(list(sheet_data)):
                try:
                    existing_num = int(existing.attrib.get("r", "0") or 0)
                except ValueError:
                    existing_num = 0
                if existing_num > row_num:
                    sheet_data.insert(idx, row)
                    inserted = True
                    break
            if not inserted:
                sheet_data.append(row)
            rows[row_num] = row
            return row

        def set_col_c(row_num: int, value: str):
            row = ensure_row(row_num)
            target_ref = f"C{row_num}"
            old = None
            style = None
            for cell in list(row):
                ref = cell.attrib.get("r", "")
                if ref == target_ref:
                    old = cell
                    style = cell.attrib.get("s")
                    break
                if ref.startswith("B") and cell.attrib.get("s"):
                    style = cell.attrib.get("s")
            if old is not None:
                row.remove(old)
            if not value:
                return
            attrs = {"r": target_ref, "t": "inlineStr"}
            if style:
                attrs["s"] = style
            cell = ET.Element(f"{{{MAIN_NS}}}c", attrs)
            is_node = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
            t = ET.SubElement(is_node, f"{{{MAIN_NS}}}t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
            t.text = value
            row.append(cell)
            cells = list(row)
            cells.sort(key=lambda c: _col_num(c.attrib.get("r", "")))
            row[:] = cells

        set_col_c(1, "Profiles")
        max_existing = max(rows.keys(), default=1)
        max_needed = max(max_existing, len(profiles) + 1)
        for row_num in range(2, max_needed + 1):
            value = profiles[row_num - 2] if row_num - 2 < len(profiles) else ""
            set_col_c(row_num, value)

        dim = root.find(f"{{{MAIN_NS}}}dimension")
        if dim is not None:
            current_ref = dim.attrib.get("ref", "A1:C1")
            end = current_ref.split(":")[-1]
            match = re.match(r"([A-Z]+)", end)
            last_col = match.group(1) if match else "C"
            if _col_num(last_col + "1") < _col_num("C1"):
                last_col = "C"
            last_row = max(max_existing, len(profiles) + 1)
            dim.set("ref", f"A1:{last_col}{last_row}")
        return ET.tostring(root, encoding="unicode")

    def _verify_profile_registry(self, zf: zipfile.ZipFile, expected: List[str]) -> None:
        shared = _load_shared_strings(zf)
        sheets = _sheet_paths(zf)
        path = sheets.get(self.SETTINGS_SHEET)
        if not path:
            raise ValueError("Save verification could not find the Settings sheet.")
        actual = []
        for row_num, values in _read_sheet_rows(zf, path, shared):
            if row_num <= 1:
                continue
            value = values.get(3, "").strip()
            if value:
                actual.append(value)
        if actual != expected:
            raise ValueError("Save verification failed for the Profiles registry.")

    def _patch_settings_merge_profiles_xml(self, xml_text: str, profiles: List[str]) -> str:
        """Persist the managed Merge Profiles registry in Settings column F."""
        ET.register_namespace("", MAIN_NS)
        root = ET.fromstring(xml_text)
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise ValueError("Settings sheet has no sheetData element.")

        rows = {}
        for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
            try:
                rows[int(row.attrib.get("r", "0") or 0)] = row
            except ValueError:
                continue

        def ensure_row(row_num: int):
            row = rows.get(row_num)
            if row is not None:
                return row
            row = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(row_num)})
            inserted = False
            for idx, existing in enumerate(list(sheet_data)):
                try:
                    existing_num = int(existing.attrib.get("r", "0") or 0)
                except ValueError:
                    existing_num = 0
                if existing_num > row_num:
                    sheet_data.insert(idx, row)
                    inserted = True
                    break
            if not inserted:
                sheet_data.append(row)
            rows[row_num] = row
            return row

        def set_col_f(row_num: int, value: str):
            row = ensure_row(row_num)
            target_ref = f"F{row_num}"
            old = None
            style = None
            for cell in list(row):
                ref = cell.attrib.get("r", "")
                if ref == target_ref:
                    old = cell
                    style = cell.attrib.get("s")
                    break
                if ref.startswith("E") and cell.attrib.get("s"):
                    style = cell.attrib.get("s")
            if old is not None:
                row.remove(old)
            if not value:
                return
            attrs = {"r": target_ref, "t": "inlineStr"}
            if style:
                attrs["s"] = style
            cell = ET.Element(f"{{{MAIN_NS}}}c", attrs)
            is_node = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
            t = ET.SubElement(is_node, f"{{{MAIN_NS}}}t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
            t.text = value
            row.append(cell)
            row[:] = sorted(list(row), key=lambda item: _col_num(item.attrib.get("r", "")))

        set_col_f(1, "Merge Profiles")
        max_existing = max(rows.keys(), default=1)
        max_needed = max(max_existing, len(profiles) + 1)
        for row_num in range(2, max_needed + 1):
            value = profiles[row_num - 2] if row_num - 2 < len(profiles) else ""
            set_col_f(row_num, value)

        dim = root.find(f"{{{MAIN_NS}}}dimension")
        if dim is not None:
            current_ref = dim.attrib.get("ref", "A1:F1")
            end = current_ref.split(":")[-1]
            match = re.match(r"([A-Z]+)", end)
            last_col = match.group(1) if match else "F"
            if _col_num(last_col + "1") < _col_num("F1"):
                last_col = "F"
            last_row = max(max_existing, len(profiles) + 1)
            dim.set("ref", f"A1:{last_col}{last_row}")
        return ET.tostring(root, encoding="unicode")

    def _verify_merge_profile_registry(self, zf: zipfile.ZipFile, expected: List[str]) -> None:
        shared = _load_shared_strings(zf)
        sheets = _sheet_paths(zf)
        path = sheets.get(self.SETTINGS_SHEET)
        if not path:
            raise ValueError("Save verification could not find the Settings sheet.")
        actual = []
        for row_num, values in _read_sheet_rows(zf, path, shared):
            if row_num <= 1:
                continue
            value = values.get(6, "").strip()
            if value:
                actual.append(value)
        if actual != expected:
            raise ValueError("Save verification failed for the Merge Profiles registry.")

    def _patch_settings_file_locations_xml(
        self, xml_text: str, file_locations: List[Tuple[str, str]]
    ) -> str:
        """Persist named physical folder locations in Settings columns D/E."""
        ET.register_namespace("", MAIN_NS)
        root = ET.fromstring(xml_text)
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise ValueError("Settings sheet has no sheetData element.")
        rows = {}
        for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
            try:
                rows[int(row.attrib.get("r", "0") or 0)] = row
            except ValueError:
                continue

        def ensure_row(row_num: int):
            row = rows.get(row_num)
            if row is not None:
                return row
            row = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(row_num)})
            inserted = False
            for idx, existing in enumerate(list(sheet_data)):
                try:
                    existing_num = int(existing.attrib.get("r", "0") or 0)
                except ValueError:
                    existing_num = 0
                if existing_num > row_num:
                    sheet_data.insert(idx, row)
                    inserted = True
                    break
            if not inserted:
                sheet_data.append(row)
            rows[row_num] = row
            return row

        def set_cell(row_num: int, col: str, value: str):
            row = ensure_row(row_num)
            target_ref = f"{col}{row_num}"
            old = None
            style = None
            for cell in list(row):
                if cell.attrib.get("r", "") == target_ref:
                    old = cell
                    style = cell.attrib.get("s")
                    break
            if style is None:
                previous_col = {"D": "C", "E": "D"}.get(col)
                if previous_col:
                    for cell in list(row):
                        if cell.attrib.get("r", "") == f"{previous_col}{row_num}":
                            style = cell.attrib.get("s")
                            break
            if old is not None:
                row.remove(old)
            if not value:
                return
            attrs = {"r": target_ref, "t": "inlineStr"}
            if style:
                attrs["s"] = style
            cell = ET.Element(f"{{{MAIN_NS}}}c", attrs)
            is_node = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
            t = ET.SubElement(is_node, f"{{{MAIN_NS}}}t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
            t.text = value
            row.append(cell)
            row[:] = sorted(list(row), key=lambda item: _col_num(item.attrib.get("r", "")))

        set_cell(1, "D", "File Location Name")
        set_cell(1, "E", "File Location Path")
        max_existing = max(rows.keys(), default=1)
        max_needed = max(max_existing, len(file_locations) + 1)
        for row_num in range(2, max_needed + 1):
            idx = row_num - 2
            name, folder_path = file_locations[idx] if idx < len(file_locations) else ("", "")
            set_cell(row_num, "D", name)
            set_cell(row_num, "E", folder_path)

        dim = root.find(f"{{{MAIN_NS}}}dimension")
        if dim is not None:
            current_ref = dim.attrib.get("ref", "A1:E1")
            end = current_ref.split(":")[-1]
            match = re.match(r"([A-Z]+)", end)
            last_col = match.group(1) if match else "E"
            if _col_num(last_col + "1") < _col_num("E1"):
                last_col = "E"
            last_row = max(max_existing, len(file_locations) + 1)
            dim.set("ref", f"A1:{last_col}{last_row}")
        return ET.tostring(root, encoding="unicode")

    def _verify_file_locations(self, zf: zipfile.ZipFile, expected: List[Tuple[str, str]]) -> None:
        shared = _load_shared_strings(zf)
        sheets = _sheet_paths(zf)
        path = sheets.get(self.SETTINGS_SHEET)
        if not path:
            raise ValueError("Save verification could not find the Settings sheet.")
        actual = []
        for row_num, values in _read_sheet_rows(zf, path, shared):
            if row_num <= 1:
                continue
            name = values.get(4, "").strip()
            folder_path = values.get(5, "").strip()
            if name and folder_path:
                actual.append((name, folder_path))
        actual.sort(key=lambda item: item[0].casefold())
        if actual != expected:
            raise ValueError("Save verification failed for the File Location registry.")

    def _patch_settings_profile_layouts_xml(
        self, xml_text: str, layouts: List[Tuple[str, str, str, str]]
    ) -> str:
        """Persist Experimental Profile Layouts in Settings columns G:J."""
        ET.register_namespace("", MAIN_NS)
        root = ET.fromstring(xml_text)
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise ValueError("Settings sheet has no sheetData element.")

        rows = {}
        for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
            try:
                rows[int(row.attrib.get("r", "0") or 0)] = row
            except ValueError:
                continue

        def ensure_row(row_num: int):
            row = rows.get(row_num)
            if row is not None:
                return row
            row = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(row_num)})
            inserted = False
            for idx, existing in enumerate(list(sheet_data)):
                try:
                    existing_num = int(existing.attrib.get("r", "0") or 0)
                except ValueError:
                    existing_num = 0
                if existing_num > row_num:
                    sheet_data.insert(idx, row)
                    inserted = True
                    break
            if not inserted:
                sheet_data.append(row)
            rows[row_num] = row
            return row

        def set_cell(row_num: int, col: str, value: str):
            row = ensure_row(row_num)
            target_ref = f"{col}{row_num}"
            old = None
            style = None
            for cell in list(row):
                if cell.attrib.get("r", "") == target_ref:
                    old = cell
                    style = cell.attrib.get("s")
                    break
            if style is None:
                previous_col = {"G": "F", "H": "G", "I": "H", "J": "I"}.get(col)
                if previous_col:
                    for cell in list(row):
                        if cell.attrib.get("r", "") == f"{previous_col}{row_num}":
                            style = cell.attrib.get("s")
                            break
            if old is not None:
                row.remove(old)
            if not value:
                return
            attrs = {"r": target_ref, "t": "inlineStr"}
            if style:
                attrs["s"] = style
            cell = ET.Element(f"{{{MAIN_NS}}}c", attrs)
            is_node = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
            t = ET.SubElement(
                is_node,
                f"{{{MAIN_NS}}}t",
                {"{http://www.w3.org/XML/1998/namespace}space": "preserve"},
            )
            t.text = value
            row.append(cell)
            row[:] = sorted(list(row), key=lambda item: _col_num(item.attrib.get("r", "")))

        headers = {
            "G": "Profile Layout Profile",
            "H": "Profile Layout Package Hash",
            "I": "Profile Layout Package",
            "J": "Profile Layout File Location",
        }
        for col, value in headers.items():
            set_cell(1, col, value)

        max_existing = max(rows.keys(), default=1)
        max_needed = max(max_existing, len(layouts) + 1)
        for row_num in range(2, max_needed + 1):
            idx = row_num - 2
            profile, package_hash, package_name, location_name = (
                layouts[idx] if idx < len(layouts) else ("", "", "", "")
            )
            set_cell(row_num, "G", profile)
            set_cell(row_num, "H", package_hash)
            set_cell(row_num, "I", package_name)
            set_cell(row_num, "J", location_name)

        dim = root.find(f"{{{MAIN_NS}}}dimension")
        if dim is not None:
            current_ref = dim.attrib.get("ref", "A1:J1")
            end = current_ref.split(":")[-1]
            match = re.match(r"([A-Z]+)", end)
            last_col = match.group(1) if match else "J"
            if _col_num(last_col + "1") < _col_num("J1"):
                last_col = "J"
            last_row = max(max_existing, len(layouts) + 1)
            dim.set("ref", f"A1:{last_col}{last_row}")
        return ET.tostring(root, encoding="unicode")

    def _verify_profile_layouts(
        self, zf: zipfile.ZipFile, expected: List[Tuple[str, str, str, str]]
    ) -> None:
        shared = _load_shared_strings(zf)
        sheets = _sheet_paths(zf)
        path = sheets.get(self.SETTINGS_SHEET)
        if not path:
            raise ValueError("Save verification could not find the Settings sheet.")
        actual: List[Tuple[str, str, str, str]] = []
        rows = _read_sheet_rows(zf, path, shared)
        headers_ok = False
        for row_num, values in rows:
            if row_num == 1:
                headers_ok = (
                    values.get(7, "").strip() == "Profile Layout Profile"
                    and values.get(8, "").strip() == "Profile Layout Package Hash"
                    and values.get(9, "").strip() == "Profile Layout Package"
                    and values.get(10, "").strip() == "Profile Layout File Location"
                )
                continue
            if not headers_ok:
                continue
            profile = values.get(7, "").strip()
            package_hash = values.get(8, "").strip().lower()
            package_name = Path(values.get(9, "").strip()).name
            location_name = values.get(10, "").strip()
            if profile and location_name and (package_hash or package_name):
                actual.append((profile, package_hash, package_name, location_name))
        actual.sort(
            key=lambda item: (item[0].casefold(), item[2].casefold(), item[1], item[3].casefold())
        )
        if actual != expected:
            raise ValueError("Save verification failed for Experimental Profile Layouts.")

    def _count_real_rows(self, zf, path: Optional[str], shared: List[str]) -> int:
        if not path:
            return 0
        count = 0
        for row_num, values in _read_sheet_rows(zf, path, shared):
            if row_num > 1 and values.get(1, "").strip():
                count += 1
        return count
