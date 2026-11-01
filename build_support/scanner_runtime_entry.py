from __future__ import print_function
import argparse
import os
from pathlib import Path
import runpy
import sys


def _runtime_root():
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def _normalize_argv(argv):
    args = list(argv)
    # Emulate the old portable Python runtime command line used by CAS Curator.
    if args and args[0] == "-u":
        args = args[1:]
    if args:
        first = str(args[0])
        low = first.lower().replace("/", "\\")
        if low.endswith("cascurator_scannerruntime.exe"):
            args = args[1:]
    return args


def _run_c_mode(args):
    code = args[1] if len(args) > 1 else ""
    sys.argv = ["-c"] + args[2:]
    ns = {"__name__": "__main__", "__package__": None}
    exec(compile(code, "<string>", "exec"), ns, ns)


def _run_script_mode(args):
    script = Path(args[0]).resolve()
    if not script.is_file():
        raise SystemExit("ERROR: Python worker script not found: %s" % script)
    sys.argv = [str(script)] + args[1:]
    runpy.run_path(str(script), run_name="__main__")


def _run_scanner(args):
    parser = argparse.ArgumentParser(description="CAS Curator bundled scanner runtime")
    parser.add_argument("--source", required=True)
    parser.add_argument("--mdb", required=True)
    parser.add_argument("--mode", choices=["INDIVIDUAL", "BIG_REVIEW_BATCH", "THUMBNAIL_BACKFILL"], default="INDIVIDUAL")
    parser.add_argument("--edition", choices=["DEV", "PUBLIC"], default="PUBLIC")
    parser.add_argument("--create-report", choices=["0", "1"], default="1")
    parser.add_argument("--preview-routing", action="store_true")
    parser.add_argument("--real-routing", action="store_true")
    parser.add_argument("--routing-root", default="")
    parser.add_argument("--preview-workdir", default="")
    ns = parser.parse_args(args)

    source = Path(ns.source).resolve()
    mdb = Path(ns.mdb).resolve()
    if not source.is_dir():
        raise SystemExit("ERROR: source folder does not exist: %s" % source)
    if not mdb.is_file():
        raise SystemExit("ERROR: MDB does not exist: %s" % mdb)
    if ns.preview_routing and ns.real_routing:
        raise SystemExit("ERROR: preview routing and real routing cannot both be enabled")
    routing_root = Path(ns.routing_root).resolve() if ns.routing_root else None
    if (ns.preview_routing or ns.real_routing) and (routing_root is None or not routing_root.is_dir()):
        raise SystemExit("ERROR: routing mode requires an existing --routing-root")

    os.environ["SIMS3_CC_SCANNER_EDITION"] = ns.edition
    os.environ["S3CC_GUI_MODE"] = "1"
    os.environ["S3CC_SOURCE_FOLDER"] = str(source)
    os.environ["S3CC_SCAN_MODE"] = ns.mode
    os.environ["S3CC_MASTER_DATABASE"] = str(mdb)
    os.environ["S3CC_MOVE_FILES"] = "1" if ns.real_routing else "0"
    os.environ["S3CC_CREATE_SCAN_REPORT"] = ns.create_report

    if ns.preview_routing:
        preview_workdir = Path(ns.preview_workdir).resolve() if ns.preview_workdir else (_runtime_root() / "routing_preview_runs")
        preview_workdir.mkdir(parents=True, exist_ok=True)
        os.environ["S3CC_SCANS_ROOT"] = str(preview_workdir)
        os.environ["S3CC_ROUTING_PREVIEW"] = "1"
        os.environ["S3CC_ROUTING_ROOT"] = str(routing_root)
        os.environ["S3CC_SKIP_MDB_UPDATE"] = "1"
    else:
        os.environ["S3CC_SCANS_ROOT"] = str(mdb.parent)
        os.environ["S3CC_ROUTING_PREVIEW"] = "0"
        os.environ["S3CC_SKIP_MDB_UPDATE"] = "0"
        if ns.real_routing:
            os.environ["S3CC_ROUTING_ROOT"] = str(routing_root)

    print("GUI INTEGRATION RUNNER: bundled Python 3.8 process ready")
    print("GUI INTEGRATION RUNNER: source=%s" % source)
    print("GUI INTEGRATION RUNNER: mdb=%s" % mdb)
    print("GUI INTEGRATION RUNNER: mode=%s" % ns.mode)
    print("GUI INTEGRATION RUNNER: edition=%s" % ("release" if ns.edition == "PUBLIC" else ns.edition.lower()))
    print("GUI INTEGRATION RUNNER: source moves ENABLED" if ns.real_routing else "GUI INTEGRATION RUNNER: source moves disabled")

    import Sims3_CC_Health_Scanner_HELPER_BOUNDARY  # noqa: F401


def main():
    root = _runtime_root()
    helper_root = root.parent / "helper"
    helper_exe = helper_root / "CASCurator_S3PI_Bridge_Helper.exe"
    helper_s3pi = helper_root / "s3pi"
    ea_file = root / "EA_Reference" / "EA_resources_HASH_TEST_v5_ALL_PACKS.json"
    if not ea_file.is_file():
        raise SystemExit("EA reference database is missing: %s" % ea_file)
    if not helper_exe.is_file():
        raise SystemExit("Separate S3PI helper is missing: %s" % helper_exe)
    if not helper_s3pi.is_dir():
        raise SystemExit("Helper S3PI DLL folder is missing: %s" % helper_s3pi)
    os.environ["CAS_CURATOR_S3PI_BRIDGE_HELPER"] = str(helper_exe)
    os.environ["CAS_CURATOR_S3PI_BRIDGE_DLL_DIR"] = str(helper_s3pi)

    args = _normalize_argv(sys.argv[1:])
    if args and args[0] == "-c":
        return _run_c_mode(args)
    if args and str(args[0]).lower().endswith(".py"):
        return _run_script_mode(args)
    return _run_scanner(args)


if __name__ == "__main__":
    main()
