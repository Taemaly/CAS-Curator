"""Read scanner and cached SQLite routing previews without dropping first row."""
from pathlib import Path


def parse_preview_rows(content):
    rows = []
    for raw in content.splitlines():
        line = raw.lstrip("\ufeff").strip("\r\n")
        if not line.strip():
            continue
        columns = line.split("\t", 4)
        if len(columns) != 5:
            raise ValueError("Malformed routing preview row: " + line[:160])
        category, duplicate, source, destination, note = columns
        # Old scanner preview includes a column-label header; cached preview
        # deliberately does not. Only skip a recognized header.
        if (category.strip().lower() in ("category", "routing category")
                and source.strip().lower() in ("source", "source path", "source_path")):
            continue
        if not category.strip() or not source.strip() or not destination.strip():
            raise ValueError("Incomplete routing preview row: " + line[:160])
        rows.append((category, duplicate, source, destination, note))
    return rows


def read_preview_rows(path):
    return parse_preview_rows(Path(path).read_text(encoding="utf-8-sig", errors="replace"))
