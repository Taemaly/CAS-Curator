from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path


def _safe_zip_stem(path: Path) -> str:
    """Return a Windows-friendly outer folder name derived from a ZIP filename."""
    stem = str(path.stem or "Archive").strip().rstrip(". ") or "Archive"
    stem = re.sub(r'[<>:"/\\|?*]', "_", stem).strip().rstrip(". ")
    return stem or "Archive"


def _zip_preview_plan(zip_path: Path) -> dict:
    """Build a safe, read-only extraction plan for a ZIP archive.

    If every file is already beneath one shared top-level folder, that folder is
    preserved as-is. Otherwise the ZIP filename becomes the outer folder so loose
    files and multi-root archives never spill directly into the chosen destination.
    """
    archive = Path(zip_path)
    if archive.suffix.lower() != ".zip":
        raise ValueError("Only ZIP archives are supported.")
    if not archive.is_file():
        raise FileNotFoundError(str(archive))

    entries = []
    file_parts = []
    with zipfile.ZipFile(archive, "r") as zf:
        bad = zf.testzip()
        if bad:
            raise ValueError(f"ZIP integrity check failed at: {bad}")
        for index, info in enumerate(zf.infolist()):
            raw = str(info.filename or "").replace("\\", "/")
            while raw.startswith("./"):
                raw = raw[2:]
            if not raw:
                continue
            if "\x00" in raw or raw.startswith("/") or re.match(r"^[A-Za-z]:", raw):
                raise ValueError(f"Unsafe absolute path in ZIP: {info.filename}")
            raw = raw.strip("/")
            parts = [part for part in raw.split("/") if part not in ("", ".")]
            if not parts or any(part == ".." for part in parts):
                raise ValueError(f"Unsafe parent path in ZIP: {info.filename}")
            reserved = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
            for part in parts:
                if part.rstrip(". ") != part or not part.rstrip(". "):
                    raise ValueError(f"Unsafe Windows path component in ZIP: {info.filename}")
                if re.search(r'[<>:"|?*]', part):
                    raise ValueError(f"Unsafe Windows path character in ZIP: {info.filename}")
                if part.split(".", 1)[0].upper() in reserved:
                    raise ValueError(f"Reserved Windows path component in ZIP: {info.filename}")
            entry = {
                "index": index,
                "member": info.filename,
                "parts": parts,
                "is_dir": bool(info.is_dir()),
                "size": int(info.file_size or 0),
            }
            entries.append(entry)
            if not info.is_dir():
                file_parts.append(parts)

    if not file_parts:
        raise ValueError("This ZIP contains no files to extract.")

    shared_top = (
        len({parts[0] for parts in file_parts}) == 1
        and all(len(parts) >= 2 for parts in file_parts)
    )
    outer = "" if shared_top else _safe_zip_stem(archive)
    for entry in entries:
        rel_parts = list(entry["parts"])
        if outer:
            rel_parts.insert(0, outer)
        entry["relative"] = str(Path(*rel_parts))

    archive_stat = archive.stat()
    return {
        "archive": archive,
        "archive_size": int(archive_stat.st_size),
        "archive_mtime_ns": int(archive_stat.st_mtime_ns),
        "entries": entries,
        "file_count": sum(1 for entry in entries if not entry["is_dir"]),
        "total_size": sum(int(entry["size"]) for entry in entries if not entry["is_dir"]),
        "uses_archive_outer_folder": bool(outer),
        "outer_folder": outer or file_parts[0][0],
    }


def _zip_conflicts(plan: dict, destination_root: Path) -> list[str]:
    root = Path(destination_root).resolve()
    conflicts = []
    for entry in plan.get("entries", []):
        if entry.get("is_dir"):
            continue
        target = (root / entry["relative"]).resolve()
        try:
            target.relative_to(root)
        except ValueError:
            conflicts.append(str(target))
            continue
        if target.exists():
            conflicts.append(str(target))
    return conflicts


def _extract_zip_plan(plan: dict, destination_root: Path) -> list[Path]:
    """Extract a previewed ZIP plan without overwriting existing files."""
    archive = Path(plan["archive"])
    archive_stat = archive.stat()
    if (
        int(archive_stat.st_size) != int(plan.get("archive_size", archive_stat.st_size))
        or int(archive_stat.st_mtime_ns) != int(plan.get("archive_mtime_ns", archive_stat.st_mtime_ns))
    ):
        raise ValueError("ZIP changed after preview; preview it again before extracting.")
    root = Path(destination_root).resolve()
    root.mkdir(parents=True, exist_ok=True)
    conflicts = _zip_conflicts(plan, root)
    if conflicts:
        raise FileExistsError(conflicts[0])

    created_files: list[Path] = []
    created_dirs: set[Path] = set()
    try:
        with zipfile.ZipFile(archive, "r") as zf:
            infos = zf.infolist()
            for entry in plan.get("entries", []):
                target = (root / entry["relative"]).resolve()
                target.relative_to(root)
                if entry.get("is_dir"):
                    if not target.exists():
                        target.mkdir(parents=True, exist_ok=True)
                        created_dirs.add(target)
                    continue
                parent = target.parent
                missing = []
                cur = parent
                while cur != root and not cur.exists():
                    missing.append(cur)
                    cur = cur.parent
                parent.mkdir(parents=True, exist_ok=True)
                created_dirs.update(missing)
                info = infos[int(entry["index"])]
                if info.filename != entry["member"]:
                    raise ValueError("ZIP contents changed after preview; preview again before extracting.")
                with zf.open(info, "r") as src, target.open("xb") as dst:
                    shutil.copyfileobj(src, dst)
                if target.stat().st_size != int(info.file_size or 0):
                    raise IOError(f"Extracted size verification failed: {target}")
                created_files.append(target)
    except Exception:
        for path in reversed(created_files):
            try:
                path.unlink(missing_ok=True)
            except Exception:
                pass
        for directory in sorted(created_dirs, key=lambda p: len(p.parts), reverse=True):
            try:
                directory.rmdir()
            except Exception:
                pass
        raise
    return created_files



def _zip_scan_extract_plan(zip_path: Path) -> dict:
    """Build the safe extraction plan used by the Scan-time ZIP intake workflow.

    Scan-time extraction always gives each archive its own outer folder beside the
    ZIP.  If the archive already has a single top-level folder with the same name
    as the ZIP, that folder is reused rather than duplicated.
    """
    plan = _zip_preview_plan(Path(zip_path))
    own_outer = _safe_zip_stem(Path(zip_path))
    existing_outer = str(plan.get("outer_folder") or "")

    # _zip_preview_plan already adds the ZIP stem for loose / multi-root archives.
    # For a single-root archive, add the ZIP stem unless the archive's own root is
    # already the same name.  This guarantees that two unrelated ZIPs cannot spill
    # into the same sibling folder during a batch Scan-time extraction.
    if not plan.get("uses_archive_outer_folder") and existing_outer.casefold() != own_outer.casefold():
        for entry in plan.get("entries", []):
            entry["relative"] = str(Path(own_outer) / entry["relative"])
        plan["outer_folder"] = own_outer
        plan["uses_archive_outer_folder"] = True

    plan["scan_outer_folder"] = own_outer
    plan["destination_root"] = Path(zip_path).resolve().parent
    return plan


def _preflight_scan_zip_plans(zip_paths) -> list[dict]:
    """Validate a batch of Scan-time ZIP extraction plans without writing files."""
    plans = []
    planned_targets: dict[str, Path] = {}
    for raw in zip_paths:
        archive = Path(raw).resolve()
        plan = _zip_scan_extract_plan(archive)
        root = Path(plan["destination_root"]).resolve()
        conflicts = _zip_conflicts(plan, root)
        if conflicts:
            raise FileExistsError(
                f"Extraction would overwrite an existing file:\n{conflicts[0]}"
            )
        for entry in plan.get("entries", []):
            if entry.get("is_dir"):
                continue
            target = (root / entry["relative"]).resolve()
            target.relative_to(root)
            key = str(target).casefold()
            previous = planned_targets.get(key)
            if previous is not None:
                raise FileExistsError(
                    "Two ZIP archives would extract to the same file:\n"
                    f"{previous}\n{target}"
                )
            planned_targets[key] = target
        plans.append(plan)
    return plans


def _extract_scan_zip_plans(plans: list[dict]) -> list[Path]:
    """Extract a preflighted Scan-time ZIP batch, rolling back created files on failure."""
    created: list[Path] = []
    try:
        for plan in plans:
            root = Path(plan["destination_root"]).resolve()
            extracted = _extract_zip_plan(plan, root)
            created.extend(extracted)
    except Exception:
        # _extract_zip_plan cleans up the archive that failed.  Roll back files
        # created by earlier archives in this batch as well, then remove empty
        # directories where possible.
        parents = set()
        for path in reversed(created):
            try:
                parents.add(path.parent)
                path.unlink(missing_ok=True)
            except Exception:
                pass
        for parent in sorted(parents, key=lambda p: len(p.parts), reverse=True):
            cur = parent
            while True:
                try:
                    cur.rmdir()
                except Exception:
                    break
                next_parent = cur.parent
                if next_parent == cur:
                    break
                cur = next_parent
        raise
    return created
