from __future__ import annotations

import shutil
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Optional

SETTINGS_KEY = "temporary_working_folder"
ROOT_MARKER = ".cas_curator_temp_root"
MANAGED_PREFIXES = ("CASCurator_", "cas_curator_")


def default_working_root() -> Path:
    """Return the default CAS Curator large-working-file location."""
    return Path(tempfile.gettempdir()) / "CAS Curator"


def configured_working_root(settings=None) -> Path:
    """Return the custom working root from QSettings, or the safe default."""
    raw = ""
    if settings is not None:
        try:
            raw = str(settings.value(SETTINGS_KEY, "") or "").strip()
        except Exception:
            raw = ""
    return Path(raw).expanduser() if raw else default_working_root()


def ensure_working_root(settings=None, root: Optional[Path] = None) -> Path:
    target = Path(root) if root is not None else configured_working_root(settings)
    target = target.expanduser().resolve()
    target.mkdir(parents=True, exist_ok=True)
    marker = target / ROOT_MARKER
    if not marker.exists():
        marker.write_text(
            "CAS Curator managed temporary working folder.\n"
            "Only CAS Curator-prefixed working items are cleaned automatically.\n",
            encoding="utf-8",
        )
    return target


def probe_working_root(root: Path) -> Path:
    """Create/delete a tiny probe so an unwritable custom location is rejected immediately."""
    target = ensure_working_root(root=Path(root))
    probe = target / "CASCurator_write_test.tmp"
    probe.write_bytes(b"CAS Curator temp folder write test\n")
    probe.unlink()
    return target


@contextmanager
def managed_temporary_directory(settings=None, *, root: Optional[Path] = None, prefix: str = "CASCurator_Work_") -> Iterator[Path]:
    """Create a disposable working directory beneath the configured large-temp root."""
    base = ensure_working_root(settings=settings, root=root)
    path = Path(tempfile.mkdtemp(prefix=prefix, dir=str(base)))
    try:
        yield path
    finally:
        shutil.rmtree(path, ignore_errors=True)


def clear_managed_temporary_items(settings=None, *, root: Optional[Path] = None) -> tuple[int, int]:
    """Remove only CAS Curator-managed temp children, never arbitrary user files.

    Returns (items_removed, failures).
    """
    base = ensure_working_root(settings=settings, root=root)
    removed = 0
    failures = 0
    for child in list(base.iterdir()):
        if child.name == ROOT_MARKER or not child.name.startswith(MANAGED_PREFIXES):
            continue
        try:
            if child.is_symlink() or child.is_file():
                child.unlink(missing_ok=True)
            elif child.is_dir():
                shutil.rmtree(child)
            removed += 1
        except Exception:
            failures += 1
    return removed, failures


def cleanup_stale_managed_temporary_items(settings=None, *, root: Optional[Path] = None, older_than_seconds: int = 24 * 60 * 60) -> tuple[int, int]:
    """Remove only stale CAS Curator temp children, leaving recent/possibly-active sessions alone."""
    base = ensure_working_root(settings=settings, root=root)
    now = time.time()
    removed = 0
    failures = 0
    for child in list(base.iterdir()):
        if child.name == ROOT_MARKER or not child.name.startswith(MANAGED_PREFIXES):
            continue
        try:
            age = max(0.0, now - child.stat().st_mtime)
            if age < max(0, older_than_seconds):
                continue
            if child.is_symlink() or child.is_file():
                child.unlink(missing_ok=True)
            elif child.is_dir():
                shutil.rmtree(child)
            removed += 1
        except Exception:
            failures += 1
    return removed, failures
