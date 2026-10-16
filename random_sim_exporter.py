from __future__ import annotations

"""DEV-only Random Sim Studio Collection Exporter helpers.

This module deliberately has no Qt or scanner dependency. It consumes the records
already loaded from the CAS Curator Master Database, builds a conservative copy
plan, then exports package copies plus human- and machine-readable manifests.

Dependency behaviour is intentionally conservative: dependencies are only followed
when a record contains an explicit dependency field and the referenced package can
be matched exactly and unambiguously in the open MDB. Nothing is guessed from CAS
metadata or filenames beyond exact package-name matching.
"""

from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, MutableMapping, Sequence, Tuple
from urllib.parse import unquote, urlparse
import hashlib
import json
import os
import re
import shutil

from random_sim_core import category_label, record_key

DEPENDENCY_FIELDS: Tuple[str, ...] = (
    "Dependencies",
    "Dependency Files",
    "Required Packages",
    "Required Package",
    "Requires",
)

BAD_HEALTH = {"broken", "malformed"}
DELETED_STATUSES = {"delete", "deleted"}


class ExportPreflightError(RuntimeError):
    """Raised before any final export folder is committed."""

    def __init__(self, problems: Sequence[str]):
        self.problems = [str(problem) for problem in problems if str(problem).strip()]
        super().__init__("\n".join(self.problems) or "Collection export preflight failed.")


def _text(value: object) -> str:
    return str(value or "").strip()


def source_path_from_record(record: Mapping[str, object]) -> Path | None:
    """Return the scanner-recorded source path, including file:// URI support."""
    raw = _text(record.get("Original Source Path", ""))
    if not raw:
        return None
    if raw.lower().startswith("file:"):
        parsed = urlparse(raw)
        path_text = unquote(parsed.path or "")
        if re.match(r"^/[A-Za-z]:/", path_text):
            path_text = path_text[1:]
        if parsed.netloc and not path_text.startswith("//"):
            path_text = f"//{parsed.netloc}{path_text}"
        return Path(path_text.replace("/", os.sep))
    return Path(unquote(raw))


def _basename(value: object) -> str:
    text = _text(value).replace("\\", "/")
    return text.rsplit("/", 1)[-1] if text else ""


def package_identity(record: Mapping[str, object]) -> str:
    """Identity used to avoid copying the same package twice."""
    package_hash = _text(record.get("Package Hash", "")).casefold()
    if package_hash:
        return f"hash:{package_hash}"
    source = source_path_from_record(record)
    if source is not None:
        return f"path:{str(source).casefold()}"
    key = record_key(record)
    if key:
        return key
    return f"package:{_text(record.get('Package', '')).casefold()}"


def dependency_references(record: Mapping[str, object]) -> List[str]:
    """Read only explicit dependency fields; never infer dependencies."""
    out: List[str] = []
    seen = set()
    for field in DEPENDENCY_FIELDS:
        raw = _text(record.get(field, ""))
        if not raw:
            continue
        parts = re.split(r"[;\n\r|]+", raw)
        # A few hand-maintained MDBs use comma-separated .package names. Split
        # only when the whole value clearly contains multiple package references.
        if len(parts) == 1 and raw.lower().count(".package") > 1 and "," in raw:
            parts = raw.split(",")
        for part in parts:
            clean = part.strip().strip('"\'')
            if not clean:
                continue
            key = clean.casefold()
            if key in seen:
                continue
            seen.add(key)
            out.append(clean)
    return out


def _is_unsafe_dependency(record: Mapping[str, object]) -> bool:
    if _text(record.get("Status", "")).casefold() in DELETED_STATUSES:
        return True
    if _text(record.get("Health Status", "")).casefold() in BAD_HEALTH:
        return True
    return False


def _dependency_lookup(records: Sequence[Mapping[str, object]]):
    by_name: MutableMapping[str, List[Mapping[str, object]]] = {}
    by_hash: MutableMapping[str, List[Mapping[str, object]]] = {}
    for record in records:
        package_name = _basename(record.get("Package", ""))
        if package_name:
            by_name.setdefault(package_name.casefold(), []).append(record)
        package_hash = _text(record.get("Package Hash", ""))
        if package_hash:
            by_hash.setdefault(package_hash.casefold(), []).append(record)
    return by_name, by_hash


def _resolve_dependency_reference(
    reference: str,
    by_name: Mapping[str, Sequence[Mapping[str, object]]],
    by_hash: Mapping[str, Sequence[Mapping[str, object]]],
) -> Tuple[Mapping[str, object] | None, str]:
    clean = _text(reference)
    if not clean:
        return None, "empty dependency reference"

    hash_candidate = clean
    if clean.casefold().startswith("sha256:"):
        hash_candidate = clean.split(":", 1)[1].strip()
    if re.fullmatch(r"[0-9a-fA-F]{64}", hash_candidate):
        matches = list(by_hash.get(hash_candidate.casefold(), ()))
    else:
        matches = list(by_name.get(_basename(clean).casefold(), ()))

    if len(matches) == 1:
        return matches[0], ""
    if not matches:
        return None, "not found in open MDB"
    return None, f"ambiguous: {len(matches)} exact MDB matches"


def _collision_token(item: Mapping[str, object]) -> str:
    package_hash = _text(item.get("package_hash", ""))
    if package_hash:
        return re.sub(r"[^0-9a-zA-Z]", "", package_hash)[:8] or "duplicate"
    identity = _text(item.get("identity", ""))
    return hashlib.sha256(identity.encode("utf-8", errors="replace")).hexdigest()[:8]


def _safe_export_filename(record: Mapping[str, object]) -> str:
    source = source_path_from_record(record)
    name = source.name if source is not None else _basename(record.get("Package", ""))
    name = name or "package.package"
    # The source filename should already be valid on its own platform, but strip
    # path separators/control characters so a malformed MDB cannot escape the folder.
    name = re.sub(r"[\\/:*?\"<>|\x00-\x1f]+", "_", name).strip(" .")
    if not name:
        name = "package.package"
    if not name.casefold().endswith(".package"):
        name += ".package"
    return name


def build_export_plan(
    selected_rows: Sequence[Mapping[str, object]],
    all_records: Sequence[Mapping[str, object]],
) -> Dict[str, object]:
    """Build a de-duplicated selected+declared-dependency copy plan.

    selected_rows are the Random Sim UI row objects: {"category": ..., "record": ...}.
    The returned plan is serialisable except for no Path objects being retained.
    """
    items_by_identity: Dict[str, Dict[str, object]] = {}
    queue: List[Dict[str, object]] = []
    warnings: List[Dict[str, str]] = []
    by_name, by_hash = _dependency_lookup(all_records)

    def add_record(record: Mapping[str, object], kind: str, category: str = "", required_by: str = ""):
        identity = package_identity(record)
        item = items_by_identity.get(identity)
        if item is None:
            item = {
                "identity": identity,
                "kind": kind,
                "categories": [],
                "required_by": [],
                "record": dict(record),
                "package": _text(record.get("Package", "")),
                "package_hash": _text(record.get("Package Hash", "")),
                "cc_type": _text(record.get("CC Type", "")),
                "health": _text(record.get("Health Status", "")),
                "risk": _text(record.get("Risk", "")),
                "source": str(source_path_from_record(record) or ""),
            }
            items_by_identity[identity] = item
            queue.append(item)
        elif kind == "selected":
            # A directly selected piece takes precedence over dependency-only status.
            item["kind"] = "selected"
        if category and category not in item["categories"]:
            item["categories"].append(category)
        if required_by and required_by not in item["required_by"]:
            item["required_by"].append(required_by)
        return item

    for row in selected_rows:
        record = row.get("record")
        if not isinstance(record, Mapping):
            continue
        category = _text(row.get("category", ""))
        add_record(record, "selected", category=category)

    # Follow explicit dependency metadata recursively. Exact, unique MDB matches only.
    processed = set()
    while queue:
        item = queue.pop(0)
        identity = _text(item.get("identity", ""))
        if identity in processed:
            continue
        processed.add(identity)
        record = item["record"]
        for reference in dependency_references(record):
            dependency, reason = _resolve_dependency_reference(reference, by_name, by_hash)
            if dependency is None:
                warnings.append({
                    "package": _text(record.get("Package", "")),
                    "reference": reference,
                    "reason": reason,
                })
                continue
            if _is_unsafe_dependency(dependency):
                warnings.append({
                    "package": _text(record.get("Package", "")),
                    "reference": reference,
                    "reason": "matched MDB package is Deleted or Broken/Malformed; not auto-copied",
                })
                continue
            add_record(
                dependency,
                "dependency",
                required_by=_text(record.get("Package", "")) or identity,
            )

    items = list(items_by_identity.values())
    used_names: Dict[str, str] = {}
    for item in items:
        record = item["record"]
        wanted = _safe_export_filename(record)
        key = wanted.casefold()
        if key in used_names and used_names[key] != item["identity"]:
            stem, suffix = os.path.splitext(wanted)
            suffix = suffix or ".package"
            wanted = f"{stem}__{_collision_token(item)}{suffix}"
            serial = 2
            while wanted.casefold() in used_names:
                wanted = f"{stem}__{_collision_token(item)}_{serial}{suffix}"
                serial += 1
        used_names[wanted.casefold()] = item["identity"]
        item["exported_as"] = wanted
        item.pop("record", None)
        item["category_labels"] = [category_label(cat) for cat in item.get("categories", [])]

    return {
        "items": items,
        "warnings": warnings,
        "selected_piece_count": len(selected_rows),
        "unique_selected_package_count": sum(1 for item in items if item.get("kind") == "selected"),
        "dependency_package_count": sum(1 for item in items if item.get("kind") == "dependency"),
    }


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def preflight_export_plan(plan: Mapping[str, object]) -> Dict[str, str]:
    """Validate all sources and verify scanner SHA256 values when available."""
    problems: List[str] = []
    verified_hashes: Dict[str, str] = {}
    for item in plan.get("items", []):
        source_text = _text(item.get("source", ""))
        package = _text(item.get("package", "")) or _text(item.get("exported_as", ""))
        if not source_text:
            problems.append(f"{package}: MDB has no Original Source Path.")
            continue
        source = Path(source_text)
        if not source.is_file():
            problems.append(f"{package}: source file is missing: {source}")
            continue
        expected = _text(item.get("package_hash", "")).casefold()
        if re.fullmatch(r"[0-9a-f]{64}", expected):
            actual = _sha256_file(source)
            if actual.casefold() != expected:
                problems.append(
                    f"{package}: source SHA256 no longer matches the MDB. Refresh/rescan before exporting."
                )
            else:
                verified_hashes[_text(item.get("identity", ""))] = actual
    if problems:
        raise ExportPreflightError(problems)
    return verified_hashes


def sanitise_set_name(name: object) -> str:
    clean = _text(name) or "Random Sim"
    clean = re.sub(r"[\\/:*?\"<>|\x00-\x1f]+", "_", clean)
    clean = re.sub(r"\s+", " ", clean).strip(" .")
    return clean[:80] or "Random Sim"


def _unique_export_folder(root: Path, base_name: str, stamp: str) -> Path:
    candidate = root / f"{base_name} - {stamp}"
    serial = 2
    while candidate.exists() or (root / f".{candidate.name}.partial").exists():
        candidate = root / f"{base_name} - {stamp} ({serial})"
        serial += 1
    return candidate


def _manifest_text(manifest: Mapping[str, object]) -> str:
    lines = [
        "CAS Curator — Random Sim Studio Collection Export",
        "=" * 53,
        f"Set: {manifest.get('set_name', '')}",
        f"Created: {manifest.get('created_at', '')}",
        f"CAS Curator: {manifest.get('app_version', '')}",
        f"Master Database: {manifest.get('master_database', '')}",
        f"Generated pieces: {manifest.get('selected_piece_count', 0)}",
        f"Unique selected packages: {manifest.get('unique_selected_package_count', 0)}",
        f"Auto-resolved dependency packages: {manifest.get('dependency_package_count', 0)}",
        "",
        "Copied packages",
        "---------------",
    ]
    for item in manifest.get("items", []):
        categories = ", ".join(item.get("category_labels", []) or []) or "Dependency"
        lines.append(
            f"- [{item.get('kind', '')}] {categories}: {item.get('package', '')} -> {item.get('exported_as', '')}"
        )
        lines.append(f"  Source: {item.get('source', '')}")
        if item.get("required_by"):
            lines.append(f"  Required by: {', '.join(item.get('required_by', []))}")
        if item.get("package_hash"):
            lines.append(f"  SHA256 (MDB): {item.get('package_hash')}")
    warnings = manifest.get("warnings", []) or []
    lines.extend(["", "Dependency notes", "----------------"])
    if warnings:
        for warning in warnings:
            lines.append(
                f"- {warning.get('package', '')}: {warning.get('reference', '')} — {warning.get('reason', '')}"
            )
    else:
        lines.append("- No unresolved explicit dependency references.")
    lines.extend([
        "",
        "Safety",
        "------",
        "This folder contains COPIES only. CAS Curator did not move, delete, or edit the source packages or MDB.",
        "Dependencies are auto-included only when explicitly declared in the MDB and matched exactly/unambiguously.",
        "",
    ])
    return "\n".join(lines)


def export_collection(
    plan: Mapping[str, object],
    destination_root: Path | str,
    set_name: object,
    app_version: str,
    master_database: Path | str | None = None,
    now: datetime | None = None,
    generation_settings: Mapping[str, object] | None = None,
) -> Tuple[Path, Dict[str, object]]:
    """Copy the planned files transactionally and write JSON/TXT manifests."""
    items = list(plan.get("items", []))
    if not items:
        raise ExportPreflightError(["There is no generated package set to export."])

    verified_hashes = preflight_export_plan(plan)
    root = Path(destination_root).expanduser()
    root.mkdir(parents=True, exist_ok=True)
    if not root.is_dir():
        raise ExportPreflightError([f"Export destination is not a folder: {root}"])

    now = now or datetime.now()
    stamp = now.strftime("%Y-%m-%d %H%M%S")
    clean_name = sanitise_set_name(set_name)
    final_folder = _unique_export_folder(root, clean_name, stamp)
    partial = root / f".{final_folder.name}.partial"
    if partial.exists():
        shutil.rmtree(partial, ignore_errors=True)
    partial.mkdir(parents=False, exist_ok=False)

    manifest_items: List[Dict[str, object]] = []
    try:
        for item in items:
            source = Path(_text(item.get("source", "")))
            target = partial / _text(item.get("exported_as", ""))
            shutil.copy2(source, target)
            copied = dict(item)
            copied["copied_size"] = target.stat().st_size
            identity = _text(item.get("identity", ""))
            if identity in verified_hashes:
                copied["sha256_verified"] = True
            manifest_items.append(copied)

        manifest: Dict[str, object] = {
            "format": "CAS Curator Random Sim Collection Manifest v1",
            "set_name": clean_name,
            "created_at": now.isoformat(timespec="seconds"),
            "app_version": app_version,
            "master_database": str(master_database or ""),
            "selected_piece_count": int(plan.get("selected_piece_count", 0) or 0),
            "unique_selected_package_count": int(plan.get("unique_selected_package_count", 0) or 0),
            "dependency_package_count": int(plan.get("dependency_package_count", 0) or 0),
            "copied_file_count": len(manifest_items),
            "warnings": list(plan.get("warnings", []) or []),
            "generation_settings": dict(generation_settings or {}),
            "items": manifest_items,
        }
        (partial / "CAS_Curator_Collection_Manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        (partial / "CAS_Curator_Collection_Manifest.txt").write_text(
            _manifest_text(manifest), encoding="utf-8"
        )
        partial.rename(final_folder)
    except Exception:
        shutil.rmtree(partial, ignore_errors=True)
        raise

    return final_folder, manifest
