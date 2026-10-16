from __future__ import print_function
import argparse
import os
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description="GUI bridge for the validated Sims 3 CC scanner core")
    default_edition = "PUBLIC" if (Path(__file__).resolve().parent / "PUBLIC_EDITION").exists() else "DEV"
    parser.add_argument("--source", required=True)
    parser.add_argument("--mdb", required=True)
    parser.add_argument("--mode", choices=["INDIVIDUAL", "BIG_REVIEW_BATCH", "THUMBNAIL_BACKFILL"], default="INDIVIDUAL")
    parser.add_argument("--edition", choices=["DEV", "PUBLIC"], default=os.environ.get("CAS_CURATOR_EDITION", default_edition).upper())
    parser.add_argument("--preview-routing", action="store_true")
    parser.add_argument("--real-routing", action="store_true")
    parser.add_argument("--routing-root", default="")
    parser.add_argument("--preview-workdir", default="")
    args = parser.parse_args()

    source = Path(args.source).resolve()
    mdb = Path(args.mdb).resolve()
    if not source.is_dir():
        raise SystemExit("ERROR: source folder does not exist: %s" % source)
    if not mdb.is_file():
        raise SystemExit("ERROR: MDB does not exist: %s" % mdb)

    routing_root = Path(args.routing_root).resolve() if args.routing_root else None
    if args.preview_routing and args.real_routing:
        raise SystemExit("ERROR: preview routing and real routing cannot be enabled together")
    if (args.preview_routing or args.real_routing) and routing_root is None:
        raise SystemExit("ERROR: routing mode requires --routing-root")
    if (args.preview_routing or args.real_routing) and not routing_root.is_dir():
        raise SystemExit("ERROR: routing root does not exist: %s" % routing_root)

    os.environ["SIMS3_CC_SCANNER_EDITION"] = args.edition
    os.environ["S3CC_GUI_MODE"] = "1"
    os.environ["S3CC_SOURCE_FOLDER"] = str(source)
    os.environ["S3CC_SCAN_MODE"] = args.mode
    os.environ["S3CC_MASTER_DATABASE"] = str(mdb)
    os.environ["S3CC_MOVE_FILES"] = "1" if args.real_routing else "0"

    if args.preview_routing:
        preview_workdir = Path(args.preview_workdir).resolve() if args.preview_workdir else (Path(__file__).resolve().parent / "routing_preview_runs")
        preview_workdir.mkdir(parents=True, exist_ok=True)
        os.environ["S3CC_SCANS_ROOT"] = str(preview_workdir)
        os.environ["S3CC_ROUTING_PREVIEW"] = "1"
        os.environ["S3CC_ROUTING_ROOT"] = str(routing_root)
        os.environ["S3CC_SKIP_MDB_UPDATE"] = "1"
    else:
        os.environ["S3CC_SCANS_ROOT"] = str(mdb.parent)
        os.environ["S3CC_ROUTING_PREVIEW"] = "0"
        os.environ["S3CC_SKIP_MDB_UPDATE"] = "0"
        if args.real_routing:
            os.environ["S3CC_ROUTING_ROOT"] = str(routing_root)

    print("GUI INTEGRATION RUNNER: Python process ready")
    print("GUI INTEGRATION RUNNER: source=%s" % source)
    print("GUI INTEGRATION RUNNER: mdb=%s" % mdb)
    print("GUI INTEGRATION RUNNER: mode=%s" % args.mode)
    print("GUI INTEGRATION RUNNER: edition=%s" % ("release" if args.edition == "PUBLIC" else args.edition.lower()))
    if args.edition == "PUBLIC":
        print("GUI INTEGRATION RUNNER: conservative CAS-only guidance enabled; script/gameplay mods are outside supported scan scope")
    print("GUI INTEGRATION RUNNER: source moves ENABLED" if args.real_routing else "GUI INTEGRATION RUNNER: source moves disabled")
    if args.real_routing:
        print("GUI INTEGRATION RUNNER: ROUTING ENABLED")
        print("GUI INTEGRATION RUNNER: routing root=%s" % routing_root)
        print("GUI INTEGRATION RUNNER: MDB write enabled")
    if args.preview_routing:
        print("GUI INTEGRATION RUNNER: ROUTING PREVIEW DRY RUN")
        print("GUI INTEGRATION RUNNER: routing root=%s" % routing_root)
        print("GUI INTEGRATION RUNNER: MDB write disabled")

    core_dir = Path(__file__).resolve().parent / "scanner_core_reference"
    import sys
    sys.path.insert(0, str(core_dir))
    __import__("Sims3_CC_Health_Scanner_CORE_INTEGRATED_TEST")


if __name__ == "__main__":
    main()
