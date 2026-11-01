from __future__ import annotations

"""Stable per-user storage for CAS Curator custom themes.

This module is intentionally independent of Qt, the MDB, and scanner code so it
can be unit-tested safely. It keeps .s3theme files outside version-specific app
folders and performs a conservative one-time style migration from legacy
``custom_themes`` folders without deleting or overwriting the originals.
"""

from pathlib import Path
import json
import os
import re
import shutil
import zipfile
from datetime import datetime


def default_custom_theme_dir(edition: str = "DEV") -> Path:
    """Return the stable per-user custom-theme directory for an edition."""
    branch = "Release" if str(edition or "DEV").strip().upper() == "PUBLIC" else "DEV"
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or Path.home())
        return base / "CASCurator" / branch / "custom_themes"
    base = Path(os.environ.get("XDG_DATA_HOME") or (Path.home() / ".local" / "share"))
    return base / "CASCurator" / branch / "custom_themes"


def legacy_custom_theme_dir(app_file: Path | str) -> Path:
    return Path(app_file).resolve().parent / "custom_themes"


def discover_legacy_custom_theme_dirs(app_file: Path | str) -> list[Path]:
    """Find likely pre-v0.4.70 custom-theme folders near the running DEV build.

    Older DEV zips stored themes beside app.py. New extracted builds are commonly
    sibling folders, so scan only nearby CAS-Curator-looking directories rather
    than the user's whole drive.
    """
    app_dir = Path(app_file).resolve().parent
    found: list[Path] = []
    direct = app_dir / "custom_themes"
    if direct.is_dir():
        found.append(direct)
    parent = app_dir.parent
    try:
        siblings = list(parent.iterdir())
    except OSError:
        siblings = []
    for sibling in siblings:
        if not sibling.is_dir() or sibling == app_dir:
            continue
        name = sibling.name.casefold().replace("-", "_")
        if not (name.startswith("cas_curator") or name.startswith("cas_v")):
            continue
        candidate = sibling / "custom_themes"
        if candidate.is_dir() and candidate not in found:
            found.append(candidate)
    return found


def migrate_discovered_legacy_themes(app_file: Path | str, target_dir: Path | str) -> list[Path]:
    copied: list[Path] = []
    for legacy in discover_legacy_custom_theme_dirs(app_file):
        copied.extend(migrate_legacy_custom_themes(legacy, target_dir))
    return copied


def _dedupe_target(target_dir: Path, stem: str) -> Path:
    candidate = target_dir / f"{stem}.s3theme"
    index = 2
    while candidate.exists():
        candidate = target_dir / f"{stem}_legacy_{index}.s3theme"
        index += 1
    return candidate


def migrate_legacy_custom_themes(legacy_dir: Path | str, target_dir: Path | str) -> list[Path]:
    """Copy legacy .s3theme files into stable storage without deleting originals.

    Existing identical files are skipped. Same-name files with different bytes are
    preserved under a deterministic ``_legacy_N`` name rather than overwritten.
    Returns the newly copied destination paths.
    """
    legacy = Path(legacy_dir)
    target = Path(target_dir)
    if not legacy.is_dir():
        return []
    try:
        if legacy.resolve() == target.resolve():
            return []
    except OSError:
        pass
    target.mkdir(parents=True, exist_ok=True)
    copied: list[Path] = []
    for source in sorted(legacy.glob("*.s3theme"), key=lambda p: p.name.casefold()):
        if not source.is_file():
            continue
        destination = target / source.name
        try:
            if destination.exists():
                if destination.read_bytes() == source.read_bytes():
                    continue
                destination = _dedupe_target(target, source.stem)
            shutil.copy2(source, destination)
            copied.append(destination)
        except OSError:
            # A broken legacy file or a read-only destination must not stop launch.
            continue
    return copied


REQUIRED_THEME_KEYS = frozenset({"accent", "accent_soft", "selection", "panel_alt", "thumb_border", "text", "muted"})
_HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


def validate_native_theme_payload(payload: object) -> dict:
    """Validate a portable CAS Curator ``.s3theme`` payload.

    Keep this Qt-free so import/export safety can be regression-tested without
    launching the application. Unknown keys are deliberately allowed so newer
    Theme Studio metadata survives a round-trip through an older compatible DEV
    build.
    """
    if not isinstance(payload, dict):
        raise ValueError("This .s3theme file is not a theme object.")
    name = str(payload.get("name", "") or "").strip()
    if not name or len(name) > 100:
        raise ValueError("This .s3theme file has a missing or invalid theme name.")
    theme = payload.get("theme")
    if not isinstance(theme, dict) or not REQUIRED_THEME_KEYS.issubset(theme):
        raise ValueError("This .s3theme file has incomplete theme settings.")
    for key in REQUIRED_THEME_KEYS:
        value = theme.get(key)
        if not isinstance(value, str) or not _HEX_RE.match(value.strip()):
            raise ValueError(f"This .s3theme file has an invalid colour for '{key}'.")
    for key in ("studio_locks", "studio_section_locks"):
        if key in payload and not isinstance(payload[key], dict):
            raise ValueError(f"This .s3theme file has invalid {key} metadata.")
    if "studio_shuffle_locks" in payload and not isinstance(payload["studio_shuffle_locks"], list):
        raise ValueError("This .s3theme file has invalid studio_shuffle_locks metadata.")
    return payload


def read_native_theme_payload(path: Path | str) -> dict:
    source = Path(path)
    try:
        payload = json.loads(source.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        raise ValueError(f"Could not read this .s3theme file: {exc}") from exc
    return validate_native_theme_payload(payload)


def export_native_theme_file(source: Path | str, destination: Path | str) -> Path:
    """Validate then copy a native theme exactly, preserving all metadata."""
    src = Path(source)
    dst = Path(destination)
    read_native_theme_payload(src)
    if dst.suffix.lower() != ".s3theme":
        dst = dst.with_suffix(".s3theme")
    try:
        if src.resolve() == dst.resolve():
            return dst
    except OSError:
        pass
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return dst



def _safe_backup_filename(name: str, fallback: str = "Custom Theme") -> str:
    """Return a readable Windows-safe filename while preserving the theme name."""
    value = str(name or "").strip() or fallback
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', value).strip().rstrip('.')
    return value or fallback


def export_theme_backup(theme_files: dict[str, Path | str], saved_colour_file: Path | str,
                        destination: Path | str, *, created_at: datetime | None = None) -> Path:
    """Create a portable ZIP backup of all custom themes plus My Colours.

    Each theme is validated before it is copied.  Theme display names are kept in
    both the archive filenames and the readable Theme_Backup_Info.txt manifest.
    The source files are copied only; nothing in the user's live theme storage is
    changed or deleted.
    """
    dst = Path(destination)
    if dst.suffix.lower() != '.zip':
        dst = dst.with_suffix('.zip')
    dst.parent.mkdir(parents=True, exist_ok=True)

    when = created_at or datetime.now()
    entries: list[tuple[str, Path, dict]] = []
    used_names: set[str] = set()
    for display_name, source_value in sorted(theme_files.items(), key=lambda item: str(item[0]).casefold()):
        source = Path(source_value)
        payload = read_native_theme_payload(source)
        theme_name = str(payload.get('name') or display_name or source.stem).strip()
        base = _safe_backup_filename(theme_name, source.stem)
        archive_name = f"{base}.s3theme"
        key = archive_name.casefold()
        index = 2
        while key in used_names:
            archive_name = f"{base}_{index}.s3theme"
            key = archive_name.casefold()
            index += 1
        used_names.add(key)
        entries.append((archive_name, source, payload))

    colour_source = Path(saved_colour_file)
    info = [
        'CAS Curator Theme Backup',
        f'Created: {when.strftime("%Y-%m-%d %H:%M:%S")}',
        f'Custom themes: {len(entries)}',
        '',
        'Themes:',
    ]
    if not entries:
        info.append('- (no custom themes)')
    for archive_name, _source, payload in entries:
        theme = payload.get('theme') if isinstance(payload.get('theme'), dict) else {}
        row_colouring = 'On' if bool(theme.get('_colour_library_rows', True)) else 'Off'
        section_locks = payload.get('studio_section_locks') or payload.get('studio_locks') or {}
        shuffle_locks = payload.get('studio_shuffle_locks') or []
        info.extend([
            f"- {payload.get('name', Path(archive_name).stem)}",
            f"  File: {archive_name}",
            f"  Colour Library rows: {row_colouring}",
            f"  Section locks: {len(section_locks) if isinstance(section_locks, dict) else 0}",
            f"  Shuffle locks: {len(shuffle_locks) if isinstance(shuffle_locks, list) else 0}",
        ])
    info.extend([
        '',
        'My Colours:',
        '- Included as My Colours/my_colours.json' if colour_source.is_file()
            else '- No saved My Colours file existed at backup time.',
        '',
        'This backup only copies your custom theme data. It does not include built-in themes,',
        'CAS Curator program files, packages, scans, or Master Database files.',
    ])

    tmp = dst.with_suffix(dst.suffix + '.tmp')
    try:
        with zipfile.ZipFile(tmp, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
            for archive_name, source, _payload in entries:
                zf.write(source, f'Themes/{archive_name}')
            if colour_source.is_file():
                zf.write(colour_source, 'My Colours/my_colours.json')
            zf.writestr('Theme_Backup_Info.txt', '\n'.join(info) + '\n')
        tmp.replace(dst)
    finally:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
    return dst
