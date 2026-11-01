from __future__ import annotations

"""Persistent per-user swatches for CAS Curator Theme Colour Studio.

This is intentionally independent of Qt, the scanner and the MDB.  Saved
colours live beside (not inside) individual .s3theme files so the same swatch
shelf can be reused across every custom theme.
"""

from pathlib import Path
import json
import os
import re

MAX_SAVED_COLOURS = 96
_HEX_RE = re.compile(r"^#?([0-9A-Fa-f]{6})$")


def default_saved_colour_file(edition: str = "DEV") -> Path:
    branch = "Release" if str(edition or "DEV").strip().upper() == "PUBLIC" else "DEV"
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or Path.home())
        return base / "CASCurator" / branch / "my_colours.json"
    base = Path(os.environ.get("XDG_DATA_HOME") or (Path.home() / ".local" / "share"))
    return base / "CASCurator" / branch / "my_colours.json"


def normalise_saved_colour(value: object) -> str:
    match = _HEX_RE.fullmatch(str(value or "").strip())
    return f"#{match.group(1).upper()}" if match else ""


def normalise_saved_colours(values) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values or []:
        colour = normalise_saved_colour(value)
        if colour and colour not in seen:
            seen.add(colour)
            result.append(colour)
        if len(result) >= MAX_SAVED_COLOURS:
            break
    return result


def load_saved_colours(path: Path | str | None = None) -> list[str]:
    source = Path(path) if path is not None else default_saved_colour_file()
    try:
        raw = json.loads(source.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return []
    if isinstance(raw, dict):
        raw = raw.get("colours", [])
    if not isinstance(raw, list):
        return []
    return normalise_saved_colours(raw)


def save_saved_colours(colours, path: Path | str | None = None) -> list[str]:
    target = Path(path) if path is not None else default_saved_colour_file()
    cleaned = normalise_saved_colours(colours)
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = {"schema": 1, "colours": cleaned}
    temp = target.with_suffix(target.suffix + ".tmp")
    try:
        temp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        temp.replace(target)
    finally:
        try:
            if temp.exists():
                temp.unlink()
        except OSError:
            pass
    return cleaned


def add_saved_colour(colours, colour: object) -> tuple[list[str], bool]:
    cleaned = normalise_saved_colours(colours)
    value = normalise_saved_colour(colour)
    if not value:
        raise ValueError("Enter a valid six-digit HEX colour such as #D5A6BD.")
    if value in cleaned:
        return cleaned, False
    if len(cleaned) >= MAX_SAVED_COLOURS:
        raise ValueError(f"My Colours can store up to {MAX_SAVED_COLOURS} colours.")
    return cleaned + [value], True


def remove_saved_colour(colours, colour: object) -> tuple[list[str], bool]:
    cleaned = normalise_saved_colours(colours)
    value = normalise_saved_colour(colour)
    if not value or value not in cleaned:
        return cleaned, False
    return [item for item in cleaned if item != value], True
