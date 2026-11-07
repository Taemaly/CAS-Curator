"""SQLite-primary scanner launch bridge for the public CAS Curator port.

This is an opt-in backend entry point. The existing XLSX public scanner runner
continues to handle legacy calls until the GUI and Windows runtime are ported.
SQLite routing is ALWAYS bounded and cached; never fall back to rescanning.
"""
from __future__ import print_function

import argparse
import os
import sqlite3
import sys
from pathlib import Path


BATCH_SIZES = (100, 250, 500, 1000)


def _validate_sqlite_library(path):
    if not path.is_file():
        raise SystemExit("ERROR: open SQLite Library file does not exist: " + str(path))
    if path.suffix.lower() not in (".db", ".sqlite", ".sqlite3"):
        raise SystemExit("ERROR: expected a SQLite Library (.db/.sqlite/.sqlite3)")
    try:
        # mode=ro: this preflight never creates an empty DB by mistake.
        from urllib.parse import quote
        uri = "file:" + quote(str(path), safe="/:\\") + "?mode=ro"
        with sqlite3.connect(uri, uri=True, timeout=10) as conn:
            integrity = conn.execute("PRAGMA quick_check").fetchone()
            meta = conn.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='meta'"
            ).fetchone()
        if not integrity or integrity[0] != "ok" or not meta:
            raise RuntimeError("database integrity or CAS Curator metadata table is invalid")
    except (OSError, sqlite3.Error, RuntimeError) as exc:
        raise SystemExit("ERROR: SQLite Library preflight failed: " + str(exc))


def main(argv=None):
    parser = argparse.ArgumentParser(description="CAS Curator R27 SQLite-primary scanner bridge")
    parser.add_argument("--source", required=True)
    parser.add_argument("--sqlite-db", required=True)
    parser.add_argument("--mdb", default="")
    parser.add_argument("--mode", choices=("INDIVIDUAL", "BIG_REVIEW_BATCH", "THUMBNAIL_BACKFILL"),
                        default="INDIVIDUAL")
    parser.add_argument("--edition", choices=("PUBLIC", "DEV"), default="PUBLIC")
    parser.add_argument("--create-report", choices=("0", "1"), default="1")
    parser.add_argument("--batch-size", type=int, default=0)
    parser.add_argument("--preview-routing", action="store_true")
    parser.add_argument("--real-routing", action="store_true")
    parser.add_argument("--route-current-batch", action="store_true")
    parser.add_argument("--routing-root", default="")
    parser.add_argument("--preview-workdir", default="")
    parser.add_argument("--approved-preview", default="")
    args = parser.parse_args(argv)

    source = Path(args.source).resolve()
    db = Path(args.sqlite_db).resolve()
    if not source.is_dir():
        raise SystemExit("ERROR: source folder does not exist: " + str(source))
    _validate_sqlite_library(db)
    if args.mode == "THUMBNAIL_BACKFILL":
        raise SystemExit("ERROR: SQLite Thumbnail Backfill is not yet supported; use the legacy MDB runner")
    if args.preview_routing and args.real_routing:
        raise SystemExit("ERROR: Preview and Run Routing cannot both be selected")
    routing = args.preview_routing or args.real_routing
    if routing:
        if args.mode != "BIG_REVIEW_BATCH" or not args.route_current_batch:
            raise SystemExit(
                "ERROR: SQLite routing requires the last completed bounded Big Review batch; no fallback scan or unlimited moves"
            )
        if args.batch_size:
            raise SystemExit("ERROR: batch size is for scanning, not routing")
        root = Path(args.routing_root).resolve() if args.routing_root else None
        if root is None or not root.is_dir():
            raise SystemExit("ERROR: selected routing destination does not exist")
        if os.path.commonpath((str(root), str(source))) in (str(root), str(source)):
            raise SystemExit("ERROR: routing source and destination must be separate folders")
        if args.real_routing:
            approved = Path(args.approved_preview).resolve() if args.approved_preview else None
            if approved is None or not approved.is_file() or not Path(str(approved) + ".approved.json").is_file():
                raise SystemExit("ERROR: approved cached Preview missing; no files moved")
    else:
        if args.route_current_batch:
            raise SystemExit("ERROR: --route-current-batch is for cached routing only")
        if args.mode == "BIG_REVIEW_BATCH" and args.batch_size not in BATCH_SIZES:
            raise SystemExit("ERROR: Big Review SQLite scan requires a bounded batch (100/250/500/1000)")
        if args.mode == "INDIVIDUAL" and args.batch_size != 0:
            raise SystemExit("ERROR: individual scans do not use a batch size")

    # All safety validation is complete before the scanner or file-moving
    # code is imported.
    os.environ["SIMS3_CC_SCANNER_EDITION"] = args.edition
    os.environ["S3CC_GUI_MODE"] = "1"
    os.environ["S3CC_SOURCE_FOLDER"] = str(source)
    os.environ["S3CC_SCAN_MODE"] = args.mode
    os.environ["S3CC_SQLITE_DATABASE"] = str(db)
    os.environ["S3CC_SQLITE_LIVE_WRITE"] = "0" if routing else "1"
    os.environ["S3CC_SCAN_BATCH_SIZE"] = str(args.batch_size)
    os.environ["S3CC_ROUTE_CURRENT_BATCH"] = "1" if routing else "0"
    os.environ["S3CC_MASTER_DATABASE"] = str(Path(args.mdb).resolve()) if args.mdb else ""
    os.environ["S3CC_CREATE_SCAN_REPORT"] = args.create_report
    os.environ["S3CC_MOVE_FILES"] = "1" if args.real_routing else "0"
    os.environ["S3CC_SKIP_MDB_UPDATE"] = "1" if routing else "0"
    os.environ["S3CC_ROUTING_PREVIEW"] = "1" if args.preview_routing else "0"
    os.environ["S3CC_SCANS_ROOT"] = str(db.parent)
    os.environ.pop("S3CC_APPROVED_PREVIEW_PATH", None)
    os.environ.pop("S3CC_ROUTING_ROOT", None)
    if routing:
        os.environ["S3CC_ROUTING_ROOT"] = str(root)
    if args.real_routing:
        os.environ["S3CC_APPROVED_PREVIEW_PATH"] = str(approved)

    print("GUI INTEGRATION RUNNER: SQLite-primary public port", flush=True)
    print("GUI INTEGRATION RUNNER: source=" + str(source), flush=True)
    print("GUI INTEGRATION RUNNER: sqlite=" + str(db), flush=True)
    print("GUI INTEGRATION RUNNER: edition=" + args.edition, flush=True)
    print("GUI INTEGRATION RUNNER: mode=" + args.mode, flush=True)
    print("GUI INTEGRATION RUNNER: source moves " + ("ENABLED" if args.real_routing else "DISABLED"), flush=True)
    if args.mode == "BIG_REVIEW_BATCH" and not routing:
        print("GUI INTEGRATION RUNNER: bounded batch=" + str(args.batch_size), flush=True)

    if args.preview_routing:
        from cached_routing_preview import preview
        output = args.preview_workdir or str(db.parent / "routing_preview_runs")
        try:
            preview(str(db), str(source), str(root), output)
        except Exception as exc:
            raise SystemExit("CACHED ROUTING PREVIEW BLOCKED: " + str(exc))
        return 0
    if args.real_routing:
        from cached_routing_move import execute
        try:
            execute(str(db), str(source), str(root), str(approved), str(db.parent))
        except Exception as exc:
            raise SystemExit("CACHED ROUTING BLOCKED: " + str(exc))
        return 0

    core_dir = Path(__file__).resolve().parent / "scanner_core_reference"
    core_file = core_dir / "Sims3_CC_Health_Scanner_R27_SQLITE.py"
    if not core_file.is_file() or "sqlite_live_writer" not in core_file.read_text(encoding="utf-8-sig", errors="replace"):
        raise SystemExit("ERROR: R27 SQLite scanner core not installed; refusing legacy fallback")
    sys.path.insert(0, str(core_dir))
    __import__("Sims3_CC_Health_Scanner_R27_SQLITE")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
