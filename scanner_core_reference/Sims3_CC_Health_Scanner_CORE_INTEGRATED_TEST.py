import os
import re
import openpyxl
from openpyxl.styles import Alignment
from openpyxl import Workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
import clr
import sys
import json
import shutil
import psutil
import hashlib
import struct
import zlib
import xml.etree.ElementTree as ET

# -------------------------
# BUILD EDITION / SHARED CORE
# -------------------------
# DEV and PUBLIC launchers both execute this exact scanner core.  The edition
# changes presentation/diagnostic access only; health checks, classification,
# routing, Master Database schema and extracted fields stay shared.
SCANNER_EDITION = str(os.environ.get("SIMS3_CC_SCANNER_EDITION", "DEV") or "DEV").strip().upper()
if SCANNER_EDITION not in {"DEV", "PUBLIC"}:
    raise RuntimeError(
        "SIMS3_CC_SCANNER_EDITION must be DEV or PUBLIC, got: "
        + repr(SCANNER_EDITION)
    )

IS_DEV_BUILD = SCANNER_EDITION == "DEV"
IS_PUBLIC_BUILD = SCANNER_EDITION == "PUBLIC"

# Both editions always retain the complete permanent MDB record.  PUBLIC may
# show fewer fields by default in the future GUI, but it must never discard or
# stop writing technical MDB fields.
MDB_FIELD_POLICY = "FULL"

PRIMARY_BUILD_MARKER = (
    "BUILD: DEV/PUBLIC SCRIPTCLASS TYPEDEF MATCH PERMANENT - 2026-09-12"
)

# Developer diagnostics can only be enabled in the DEV launcher.  They remain
# off by default so normal DEV validation scans keep the current quiet behaviour.
def _env_flag(name, default=False):
    value = os.environ.get(name)
    if value is None:
        return bool(default)
    return str(value).strip().lower() in {"1", "true", "yes", "on"}

DEBUG = IS_DEV_BUILD and _env_flag("SIMS3_CC_DEBUG", False)
DEBUG_RESOURCES = IS_DEV_BUILD and _env_flag("SIMS3_CC_DEBUG_RESOURCES", False)
DEBUG_BUILD = IS_DEV_BUILD and _env_flag("SIMS3_CC_DEBUG_BUILD", False)
ALLOW_EXPERIMENTAL_FEATURES = IS_DEV_BUILD and _env_flag(
    "SIMS3_CC_EXPERIMENTAL", False
)

# -------------------------
# CONSOLE TXT LOG — QUIET FULL-SCAN MODE
# -------------------------
# Output-only optimisation:
# - routine per-resource success chatter is suppressed before terminal/file I/O
# - useful package progress, routing, classifications, summaries and problems remain
# - log-file writes are buffered and flushed periodically instead of on every
#   fragment of every print() call
#
# Scanner health checks, risk logic, resource parsing, classification and
# Master Database behaviour are unchanged.  PUBLIC is always quiet; DEV can
# temporarily opt out by setting SIMS3_CC_QUIET=0.
QUIET_FULL_SCAN = True if IS_PUBLIC_BUILD else _env_flag("SIMS3_CC_QUIET", True)


class _ScannerConsoleLogger:
    def __init__(self, terminal):
        self.terminal = terminal
        self.buffer = []
        self.log_file = None
        self.pending = ""
        self.kept_since_flush = 0
        self.flush_every = 32

    def _keep_quiet_line(self, line):
        if not QUIET_FULL_SCAN:
            return True

        text = line.strip()

        if not text:
            return False

        # One compact build identity replaces the long historical BUILD list.
        # Keep the shared-core build marker, not an older historical QUIET line.
        if text.startswith("BUILD:"):
            return text == PRIMARY_BUILD_MARKER

        # DEV/PUBLIC identity must always survive quiet logging so uploaded logs
        # tell us exactly which launcher/edition produced the scan.
        if text.startswith("SCANNER EDITION:"):
            return True

        # Scan identity / progress / destinations.
        if text.startswith((
            "LOCK LOCATION:",
            "EA HASH REFERENCE:",
            "SOURCE PATH MANIFEST:",
            "SCAN SOURCE:",
            "SCAN SESSION:",
            "TOTAL PACKAGES AVAILABLE:",
            "BATCH MODE:",
            "TOTAL PACKAGES TO SCAN:",
            "Scanning ",
            "MOVING FILES...",
            "MOVED:",
            "FINAL SUMMARY:",
            "CASP RECORDS FOUND:",
            "Finished! Report saved:",
            "Console log saved:",
            "Scan time:",
            "MASTER DB:",
            "THUMBNAIL SUMMARY:",
            "THUMBNAIL EXTRACTION ERROR:",
            "THUMBNAIL BACKFILL MODE",
            "THUMBNAIL BACKFILL:",
            "THUMBNAIL BACKFILL PROGRESS:",
            "THUMBNAIL BACKFILL SUMMARY:",
            "Packages supplied:",
            "Packages matched to existing MDB SHA256:",
            "New thumbnails extracted:",
            "Existing thumbnails reused:",
            "No dedicated preview available:",
            "Duplicate input hashes reused in this run:",
            "Not already in MDB (left untouched):",
            "Sims 2 / Sims 4 packages skipped:",
            "Empty/hash failures skipped:",
            "Thumbnail errors:",
            "MDB Image cells updated:",
            "Thumbnail folder:",
            "Backfill time:",
            "Full health/classification/risk scan:",
            "File moving/routing:",
            "Master DB changes:",
            "Identity match:",
            "ROUTING SUMMARY:",
            "Sims3Pack files found:",
            "Archive files found:",
            "Image files found:",
            "Non-package files were moved intact; no archives were extracted.",
            "ROUTING PREVIEW:",
            "ROUTING PREVIEW ROOT:",
            "ROUTING PREVIEW SOURCE RECORDS:",
            "ROUTING PREVIEW ITEM\t",
            "Routing preview saved:",
            "ROUTING PREVIEW COUNT:",
        )):
            return True

        # The session folder is printed on the line after SCAN SESSION.
        if "Health_Scan_Result_" in text:
            return True

        # Keep actual category decisions; suppress the many lower-level marker
        # lines that merely feed those decisions.
        if " CLASSIFICATION:" in text:
            return True

        # Custom Content is the common case and is already reflected by routing.
        # Keep EA Match / Override / Candidate / other unusual content results.
        if text.startswith("EA CONTENT CHECK:"):
            return "Result= Custom Content" not in text

        # Rare/important diagnostics should never disappear in quiet mode.
        important_words = (
            "FAILURE",
            "FAILED",
            "BROKEN",
            "ERROR",
            "WARNING",
            "UNSUPPORTED",
            "CORRUPT",
            "INVALID",
            "LOCKED",
            "RETRYING",
            "STALE LOCK",
            "SKIPPED DUPLICATE",
            "MOVE FAILED",
            "CANCELLED",
            "DUPLICATE",
            "UNKNOWN RESOURCE TYPES",
            "EXTREME",
            "VERY HEAVY",
            "HEAVY",
            "4K TEXTURE",
        )

        upper = text.upper()

        if any(word in upper for word in important_words):
            return True

        # In quiet mode the first FAILURE/ERROR line already contains the
        # useful resource ID and exception summary. Suppress repetitive CLR
        # stack-frame continuation lines ("at s3pi...", "at System...").
        # Python traceback headers are retained for an actual scanner crash.
        if (
            text.startswith("Traceback ")
            or text.startswith('File "')
        ):
            return True

        return False

    def _emit_line(self, line):
        if not self._keep_quiet_line(line):
            return

        self.terminal.write(line)

        if self.log_file is not None:
            self.log_file.write(line)
            self.kept_since_flush += 1

            if self.kept_since_flush >= self.flush_every:
                self.log_file.flush()
                self.kept_since_flush = 0
        else:
            self.buffer.append(line)

    def write(self, message):
        self.pending += str(message)

        while "\n" in self.pending:
            line, self.pending = self.pending.split("\n", 1)
            self._emit_line(line + "\n")

        return len(message)

    def flush(self):
        # Do not normally force out an incomplete line; print() completes with
        # a newline. If a third-party writer leaves a partial diagnostic, retain
        # it when an explicit flush occurs.
        if self.pending:
            partial = self.pending
            self.pending = ""

            if self._keep_quiet_line(partial):
                self.terminal.write(partial)

                if self.log_file is not None:
                    self.log_file.write(partial)
                else:
                    self.buffer.append(partial)

        try:
            self.terminal.flush()
        except Exception:
            pass

        if self.log_file is not None:
            try:
                self.log_file.flush()
                self.kept_since_flush = 0
            except Exception:
                pass

    def isatty(self):
        try:
            return self.terminal.isatty()
        except Exception:
            return False

    def start_log(self, path):
        if self.log_file is not None:
            return

        self.log_file = open(path, "w", encoding="utf-8")

        if self.buffer:
            self.log_file.write("".join(self.buffer))
            self.log_file.flush()
            self.buffer = []

    def close(self):
        self.flush()

        if self.log_file is not None:
            try:
                self.log_file.close()
            except Exception:
                pass
            self.log_file = None


_scanner_console_logger = _ScannerConsoleLogger(sys.stdout)
sys.stdout = _scanner_console_logger

SCANNER_VERSION = f"1.0 {SCANNER_EDITION} - SHARED CORE"
print(f"SCANNER EDITION: {SCANNER_EDITION} | MDB fields: {MDB_FIELD_POLICY}")
print(PRIMARY_BUILD_MARKER)
print("BUILD: COLLECTION + NMAP/STBL + IMG-ONLY TEXTURE CLASSIFIERS + MALFORMED HEALTH TEST - 2026-09-05")
print("BUILD: QUIET SCAN SUMMARY 3-SHEET REPORT + CC PRIORITY/RISK/EA + NON-PACKAGE ROUTER TEST - 2026-09-06")
print("BUILD: CONSOLE TXT LOG RESTORED - 2026-08-21")
print("BUILD: PNG IMAG HEALTH PERMANENT - 2026-08-16")
print("BUILD: TXTC HEALTH PERMANENT - 2026-08-16")
print("BUILD: PTRN HEALTH PERMANENT - 2026-08-16")
print("BUILD: SIMO HEALTH PERMANENT - 2026-08-16")
print("BUILD: TXTF HEALTH PERMANENT - 2026-08-16")
print("BUILD: FAMD HEALTH PERMANENT - 2026-08-16")
print("BUILD: OBJD HEALTH PERMANENT - 2026-08-16")
print("BUILD: OBJK HEALTH PERMANENT - 2026-08-16")
print("BUILD: XML HEALTH PERMANENT - 2026-08-16")
print("BUILD: STBL HEALTH PERMANENT - 2026-08-16")
print("BUILD: NMAP HEALTH PERMANENT - 2026-08-16")
print("BUILD: RSLT HEALTH PERMANENT - 2026-08-16")
print("BUILD: FTPT HEALTH PERMANENT - 2026-08-16")
print("BUILD: RIG HEALTH PERMANENT - 2026-08-16")
print("BUILD: JAZZ HEALTH PERMANENT - 2026-08-16")
print("BUILD: S3SA HEALTH PERMANENT - 2026-08-16")
print("BUILD: LITE HEALTH PERMANENT - 2026-08-17")
print("BUILD: CLIP HEALTH PERMANENT - 2026-08-17")
print("BUILD: MATD HEALTH PERMANENT - 2026-08-17")
print("BUILD: MTST HEALTH PERMANENT - 2026-08-17")
print("BUILD: MDLR HEALTH PERMANENT - 2026-08-17")
print("BUILD: CWAL HEALTH PERMANENT - 2026-08-17")
print("BUILD: CFEN HEALTH PERMANENT - 2026-08-17")
print("BUILD: CTPT HEALTH PERMANENT - 2026-08-17")
print("BUILD: NGMP HEALTH PERMANENT - 2026-08-17")
print("BUILD: CRMT HEALTH PERMANENT - 2026-08-17")
print("BUILD: SPT2 HEALTH PERMANENT - 2026-08-17")
print("BUILD: _SPT HEALTH PERMANENT - 2026-08-17")
print("BUILD: CRAL HEALTH PERMANENT - 2026-08-17")
print("BUILD: LAYO HEALTH PERMANENT - 2026-08-17")
print("BUILD: HINF HEALTH PERMANENT - 2026-08-17")
print("BUILD: 2ARY HEALTH PERMANENT - 2026-08-17")
print("BUILD: COMPLATE HEALTH PERMANENT - 2026-08-17")
print("BUILD: 0x63A33EA7 HEALTH PERMANENT - 2026-08-17")
print("BUILD: PACKAGE STORAGE HEALTH PERMANENT - 2026-08-17")
print("BUILD: FIREPLACE 0x04F3CC01 HEALTH PERMANENT - 2026-08-17")
print("BUILD: THUM 0x626F60CD MIXED PNG/JPEG/BMP HEALTH PERMANENT - 2026-08-22")
print("BUILD: 0x626F60CE MIXED PNG/JPEG/BMP HEALTH PERMANENT - 2026-08-22")
print("BUILD: ICON 0x2E75C765 MIXED PNG/JPEG HEALTH PERMANENT - 2026-08-22")
print("BUILD: THUM 0x0580A2B5 MIXED PNG/JPEG HEALTH PERMANENT - 2026-08-22")
print("BUILD: THUM 0x0580A2B4 MIXED PNG/JPEG HEALTH PERMANENT - 2026-08-22")
print("BUILD: THUM 0x0580A2B6 MIXED PNG/JPEG HEALTH PERMANENT - 2026-08-22")
print("BUILD: SNAP 0x0580A2CD PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: SNAP 0x0580A2CF PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: ICON 0x2E75C764 PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: ICON 0x2E75C766 PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: ICON 0x2E75C767 PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: THUM 0x626F60CC MIXED PNG/JPEG HEALTH PERMANENT - 2026-08-19")
print("BUILD: THUM 0x05B1B524 PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: THUM 0x05B1B525 PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: THUM 0x05B1B526 PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: THUM 0x2653E3C9 FENCE PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: THUM 0xFCEAB65B CUSTOM COLOURSWATCH PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: ICON 0xD84E7FC5 SMALL LOT PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: ICON 0xD84E7FC6 MEDIUM LOT PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: ICON 0xD84E7FC7 LARGE LOT PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: TWNI 0x0668F635 WORLD DESCRIPTION IMAGE PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: TWNI 0x0668F635 REPORT HEADER FIX - 2026-08-19")
print("BUILD: THUM 0x2D4284F1 RAILING PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: MANIFEST 0x73E93EEB XML HEALTH PERMANENT - 2026-08-19")
print("BUILD: 0x8F78F74B PNG HEALTH PERMANENT - 2026-08-19")
print("BUILD: CASP HEALTH PERMANENT - 2026-08-19")
print("BUILD: GEOM HEALTH PERMANENT - 2026-08-19")
print("BUILD: GEOM LEGACY MULTI-GROUP COMPATIBILITY PERMANENT - 2026-08-21")
print("BUILD: IMG DDS HEALTH PERMANENT - 2026-08-19")
print("BUILD: VPXY HEALTH PERMANENT - 2026-08-19")
print("BUILD: MODL HEALTH PERMANENT - 2026-08-19")
print("BUILD: MLOD HEALTH PERMANENT - 2026-08-19")
print("BUILD: BONE HEALTH PERMANENT - 2026-08-19")
print("BUILD: BOND HEALTH PERMANENT - 2026-08-19")
print("BUILD: BGEO HEALTH PERMANENT - 2026-08-19")
print("BUILD: FACE HEALTH PERMANENT - 2026-08-19")
print("BUILD: FBLN HEALTH PERMANENT - 2026-08-19")
print("BUILD: PRESET HEALTH PERMANENT - 2026-08-19")
print("BUILD: BBLN HEALTH PERMANENT - 2026-08-19")
print("BUILD: TXTC COMPATIBILITY SAFEGUARD PRECISE - 2026-08-16")


def calculate_file_hash(filename):

    digest = hashlib.sha256()

    with open(filename, "rb") as f:

        while True:

            chunk = f.read(1024 * 1024)

            if not chunk:
                break

            digest.update(chunk)

    return digest.hexdigest()
# -------------------------
# SINGLE INSTANCE LOCK
# -------------------------

LOCK_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "Sims3_CC_Health_Scanner.lock"
)

print("BUILD: TONE HEALTH PERMANENT - 2026-08-16")
print("BUILD: FACE LEGACY TGI-SIZE COMPATIBILITY PERMANENT - 2026-08-23")
print("BUILD: BOND EXTENDED-KEY COMPATIBILITY PERMANENT - 2026-08-23")
print("BUILD: MASTER DATABASE INDEXED FAST UPDATE TEST - 2026-08-24")
print("BUILD: CUSTOM SIM PERFORMANCE-RISK EXEMPTION PERMANENT - 2026-08-26")
print("BUILD: SIM PACKAGE STRUCTURAL CLASSIFICATION PERMANENT - 2026-08-26")
print("BUILD: GEOM RISK THRESHOLD LEGACY PARITY PERMANENT - 2026-08-26")
print("BUILD: MASTER DB SUBTYPE + PACKAGE FORMAT MIRROR PERMANENT - 2026-08-27")
print("BUILD: ORIGINAL SOURCE PATH MANIFEST PERMANENT - 2026-08-27")
print("BUILD: MASTER DB FILTER RANGE PERMANENT - 2026-08-28")
print("BUILD: EA CONTENT-HASH MATCHING V4.3 NMAP AUXILIARY TEST - 2026-08-28")
print("BUILD: EA CONTENT-HASH MATCHING V4.4 MANIFEST AUXILIARY TEST - 2026-08-28")
print("BUILD: DEFAULT REPLACEMENT AUXILIARY PREVIEW EA-OVERRIDE FIX TEST - 2026-09-07")
print("BUILD: EA HASH REFERENCE V5 ALL INSTALLED PACKS TEST - 2026-08-28")
print("BUILD: EA CONTENT-HASH ALL-PACKS CLASSIFICATION PERMANENT - 2026-08-28")
print("BUILD: CAREER CLASSIFICATION CAREERFILES XML PERMANENT - 2026-08-28")
print("BUILD: TUNING MOD CLASSIFICATION CURRENT_TUNING XML PERMANENT - 2026-08-28")
print("BUILD: UI RESOURCE COMPATIBILITY JPEG-IN-PNG + EMPTY LAYO METADATA PERMANENT - 2026-08-28")
print("BUILD: SCAN SUMMARY UNIDENTIFIED CC COUNT PERMANENT - 2026-08-28")
print("BUILD: RANDOM EVENT MODULE CLASSIFICATION RANDOM_EVENTS XML PERMANENT - 2026-08-29")
print("BUILD: INTERACTION TUNING CLASSIFICATION ITUN XML PERMANENT - 2026-08-29")
print("BUILD: SCRIPT MOD CLASSIFICATION VALID S3SA PERMANENT - 2026-08-29")
print("BUILD: UI LAYOUT MOD CLASSIFICATION VALID LAYO PERMANENT - 2026-08-29")
print("BUILD: PET MARKING CASP PELTLAYER 0x35 CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: INI MOD CLASSIFICATION _INI 0x1F886EAD PERMANENT - 2026-08-30")
print("BUILD: HAIR COLOR PRESET CLASSIFICATION VALID HINF PERMANENT - 2026-08-30")
print("BUILD: THUMBNAIL MOD PURE 0x626F60CD/0x626F60CE CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: FONT MOD NMAP + _CSS 0x025C90A6 + _TTF 0x062E9EE0 CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: LOT CLASSIFICATION LDES 0xD063545B + LOT PREVIEW PERMANENT - 2026-08-30")
print("BUILD: THUMBNAIL MOD V2 0xAD366F95 + NMAP AUXILIARY PERMANENT - 2026-08-30")
print("BUILD: POSE CLASSIFICATION PURE NMAP + ONE-FRAME CLIP PERMANENT - 2026-08-30")
print("BUILD: ANIMATION CLASSIFICATION PURE NMAP + MULTI-FRAME CLIP PERMANENT - 2026-09-02")
print("BUILD: CAREER TUNING V2 ACTIVE/SKILL-BASED XML PERMANENT - 2026-09-02")
print("BUILD: STRING TABLE MOD PURE STBL 0x220557DA CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: FOOTPRINT MOD PURE FTPT 0xD382BF57 CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: OPPORTUNITY TUNING XML + NMAP CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: JAZZ MOD NMAP + JAZZ 0x02D5DF13 CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: ACADEMIC DEGREE TUNING ACADEMICDEGREES XML CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: CAREER TUNING CAREERS XML + NMAP CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: BOOK TUNING BOOKS XML + NMAP CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: DUMPSTER DIVING TUNING DUMSTERDIVEDATA XML + NMAP CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: DREAMS AND PROMISES TUNING DREAMSANDPROMISESNODES XML CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: SCORING TUNING SCORING XML + NMAP CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: BOARDING SCHOOL TUNING BOARDINGSCHOOL XML + NMAP CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: BUFF TUNING BUFFS XML + NMAP CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: SOCIAL TUNING SOCIAL XML + NMAP CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: SKILL TUNING SKILLS XML + NMAP CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: TEXTURE MOD PURE NMAP + DDS IMG CLASSIFICATION PERMANENT - 2026-08-30")
print("BUILD: FISHING TUNING FISHING XML + NMAP + STBL CLASSIFICATION V2 PERMANENT - 2026-08-31")
print("BUILD: COOKING TUNING MULTI-TABLE XML + STORE INGREDIENTS CLASSIFICATION PERMANENT - 2026-08-31")
print("BUILD: SMART SOURCE ROUTING + BATCH500 + TERRAIN PAINT CLASSIFICATION PERMANENT - 2026-08-31")
print("BUILD: FINAL PRE-PUBLIC CLASSIFICATION PERMANENT - EMPTY PACKAGE + SIME/SIMO CUSTOM SIM + EXPANDED TERRAIN PAINT - 2026-09-01")
print("BUILD: QUIET FULL-SCAN + EXACT SHA256 DUPLICATE ROUTING + MASTER DB REAL-PACKAGE APPEND TEST - 2026-09-02")

print("BUILD: SCRIPT MOD EMBEDDED CAS PRIMARY CLASSIFICATION TEST - 2026-09-07")
print("BUILD: HYBRID SCRIPTED OBJECT + PATTERN PRIORITY CLASSIFICATION TEST - 2026-09-07")

print("LOCK LOCATION:", LOCK_FILE)


if os.path.exists(LOCK_FILE):

    try:

        with open(LOCK_FILE, "r") as f:
            old_pid = int(f.read().strip())


        # Check if process still exists

        if psutil.pid_exists(old_pid):

            print("")
            print("==========================================")
            print(" Sims 3 CC Health Scanner")
            print("==========================================")
            print("")
            print("ERROR:")
            print("Scanner is already running.")
            print("")
            print("PID:", old_pid)
            print("")

            if os.environ.get("S3CC_GUI_MODE", "0") != "1":
                input("Press Enter to exit...")
            sys.exit(11)


        else:

            print("")
            print("STALE LOCK DETECTED")
            print("Removing old lock...")
            print("")

            os.remove(LOCK_FILE)


    except Exception as e:

        print("")
        print("LOCK CHECK ERROR:")
        print(e)
        print("")

        if os.environ.get("S3CC_GUI_MODE", "0") != "1":
            input("Press Enter to exit...")
        sys.exit(11)



# create lock

with open(LOCK_FILE, "w") as f:
    f.write(str(os.getpid()))


# Ensure the lock is removed when Python exits normally,
# including after an unhandled Python exception.
import atexit

def cleanup_lock_file():

    try:
        if os.path.exists(LOCK_FILE):

            with open(LOCK_FILE, "r") as f:
                lock_pid = f.read().strip()

            if lock_pid == str(os.getpid()):
                os.remove(LOCK_FILE)

    except Exception as e:
        print("LOCK CLEANUP ERROR:", e)


atexit.register(cleanup_lock_file)


S3PE_PATH = r"G:\Game Tools\Sims\Sims 3 tools\s3pe"

sys.path.append(S3PE_PATH)

clr.AddReference(S3PE_PATH + r"\s3pi.Interfaces.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.Package.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.WrapperDealer.dll")

clr.AddReference(S3PE_PATH + r"\s3pi.CASPartResource.dll")

clr.AddReference(S3PE_PATH + r"\s3pi.DefaultResource.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.GenericRCOLResource.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.RCOLChunks.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.RigResource.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.JazzResource.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.ScriptResource.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.MeshChunks.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.ModularResource.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.NGMPHashMapResource.dll")
clr.AddReference(S3PE_PATH + r"\s3pi.CatalogResource.dll")
from CASPartResource import CASPartResourceHandler

CASPartResourceHandler()

EA_DATABASE = r"G:\Game Tools\Sims\Sims 3 tools\Sims3_Texture_Scan\EA_Reference\EA_resources_HASH_TEST_v5_ALL_PACKS.json"

with open(EA_DATABASE, "r", encoding="utf-8") as f:
    EA_RESOURCES = json.load(f)

EA_LOOKUP = set()
EA_HASH_LOOKUP = {}
EA_UNVERIFIABLE_TGIS = set()
EA_REFERENCE_INVALID_RECORDS = 0

for r in EA_RESOURCES:
    key = (
        int(r["type"], 16),
        int(r["group"], 16),
        int(r["instance"], 16)
    )

    EA_LOOKUP.add(key)

    hash_method = str(r.get("hash_method", "raw") or "raw").strip().lower()
    digest = str(r.get("sha256", "") or "").strip().lower()

    if hash_method == "unverifiable":
        EA_UNVERIFIABLE_TGIS.add(key)
        continue

    if hash_method not in ("raw", "wrapper") or len(digest) != 64:
        EA_REFERENCE_INVALID_RECORDS += 1
        continue

    EA_HASH_LOOKUP.setdefault(key, {}).setdefault(hash_method, set()).add(digest)

if EA_REFERENCE_INVALID_RECORDS:
    raise RuntimeError(
        "EA hash reference contains %s invalid hash records"
        % EA_REFERENCE_INVALID_RECORDS
    )

print(
    "EA HASH REFERENCE: TGIs=%s | exact variants=%s | unverifiable TGIs=%s"
    % (
        format(len(EA_LOOKUP), ","),
        format(
            sum(
                len(hashes)
                for methods in EA_HASH_LOOKUP.values()
                for hashes in methods.values()
            ),
            ","
        ),
        format(len(EA_UNVERIFIABLE_TGIS), ","),
    )
)

from s3pi.Package import *
from s3pi.WrapperDealer import *


def _get_raw_resource_bytes(package, entry):
    """
    Read the unparsed/decompressed resource bytes directly from the package.
    Used only as a conservative fallback when the s3pi GEOM wrapper rejects
    a documented legacy multi-face-group GEOM before it can expose the mesh.
    """
    import System
    from System.Reflection import BindingFlags

    package_type = package.GetType()
    get_resource_method = None
    fallback_method = None

    methods = package_type.GetMethods(
        BindingFlags.Instance |
        BindingFlags.Public |
        BindingFlags.NonPublic
    )

    for method in methods:
        try:
            if method.Name != "GetResource":
                continue

            params = method.GetParameters()
            if len(params) != 1:
                continue

            if fallback_method is None:
                fallback_method = method

            param_type_name = str(
                params[0].ParameterType.FullName or ""
            )

            if "IResourceIndexEntry" in param_type_name:
                get_resource_method = method
                break
        except Exception:
            continue

    if get_resource_method is None:
        get_resource_method = fallback_method

    if get_resource_method is None:
        raise RuntimeError("Could not resolve package GetResource method")

    stream = None

    try:
        args = System.Array[System.Object]([entry])
        stream = get_resource_method.Invoke(package, args)

        if stream is None or not stream.CanRead:
            raise RuntimeError("Package returned an unreadable resource stream")

        if stream.CanSeek:
            stream.Position = 0

        data = bytearray()
        buffer = System.Array.CreateInstance(System.Byte, 65536)

        while True:
            count = stream.Read(buffer, 0, buffer.Length)
            if count <= 0:
                break

            data.extend(int(buffer[i]) for i in range(count))

        return bytes(data)

    finally:
        try:
            if stream is not None:
                stream.Dispose()
        except Exception:
            try:
                if stream is not None:
                    stream.Close()
            except Exception:
                pass


def _get_resource_sha256(package, entry):
    """Return SHA-256 of the same unparsed/decompressed resource bytes used by
    the scanner's conservative raw-resource fallbacks."""
    raw = _get_raw_resource_bytes(package, entry)
    return hashlib.sha256(raw).hexdigest()


def _get_resource_sha256_wrapper(package, entry):
    """Hash WrapperDealer.AsBytes for EA reference records that had to use
    that same fallback while the EA reference was built."""
    resource = None
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        if resource is None:
            raise RuntimeError("WrapperDealer returned no resource")
        raw = bytes(resource.AsBytes)
        return hashlib.sha256(raw).hexdigest()
    finally:
        resource = None


def _get_resource_sha256_by_method(package, entry, method):
    if method == "raw":
        return _get_resource_sha256(package, entry)
    if method == "wrapper":
        return _get_resource_sha256_wrapper(package, entry)
    raise RuntimeError("Unsupported EA hash method: " + str(method))


def _validate_legacy_multigroup_geom(raw):
    """
    Validate the documented Sims 3 GEOM variant where the face/index section
    contains more than one group.

    Old s3pi GEOM code hard-codes the group count to 1 and throws before
    reading these otherwise valid meshes. This fallback does NOT simply ignore
    that exception: it independently validates the whole legacy layout.

    Conservative scope:
      - exactly one GEOM chunk in this resource
      - GEOM version 5 only
      - more than one face group
      - every face point is a 2-byte index
      - every group count is triangle-aligned
      - every face index is within the vertex table
      - skin index + bone hash list land exactly at the declared TGI block
      - declared TGI block lands exactly at the end of the resource
    """
    def need(pos, size, label):
        if pos < 0 or size < 0 or pos + size > len(raw):
            raise ValueError(
                f"{label} truncated/out of bounds at 0x{pos:X}"
            )

    def u32(pos, label):
        need(pos, 4, label)
        return struct.unpack_from("<I", raw, pos)[0]

    def i32(pos, label):
        need(pos, 4, label)
        return struct.unpack_from("<i", raw, pos)[0]

    geom_start = raw.find(b"GEOM")

    if geom_start < 0:
        return False, {}, "GEOM tag not found in raw resource"

    if raw.find(b"GEOM", geom_start + 4) >= 0:
        return (
            False,
            {},
            "multiple GEOM tags found; compatibility fallback refuses ambiguity"
        )

    try:
        version = u32(geom_start + 4, "GEOM version")

        if version != 0x00000005:
            return (
                False,
                {},
                f"legacy multi-group fallback supports GEOM version 5 only; found 0x{version:08X}"
            )

        tgi_offset = u32(geom_start + 8, "TGI offset")
        tgi_size = u32(geom_start + 12, "TGI size")
        tgi_pos = geom_start + 12 + tgi_offset
        tgi_end = tgi_pos + tgi_size

        if tgi_pos <= geom_start + 20:
            return False, {}, "declared TGI block overlaps GEOM header"

        if tgi_end != len(raw):
            return (
                False,
                {},
                f"declared TGI block does not end at resource end "
                f"(TGI end={tgi_end}, resource={len(raw)})"
            )

        shader = u32(geom_start + 16, "shader")
        pos = geom_start + 20

        if shader != 0:
            mtnf_size = u32(pos, "MTNF size")
            pos += 4
            need(pos, mtnf_size, "MTNF")

            if mtnf_size < 4 or raw[pos:pos + 4] != b"MTNF":
                return False, {}, "embedded material does not begin with MTNF"

            pos += mtnf_size

        need(pos, 12, "GEOM mesh header")
        merge_group = u32(pos, "merge group")
        sort_order = u32(pos + 4, "sort order")
        num_vertices = i32(pos + 8, "vertex count")
        pos += 12

        if num_vertices <= 0 or num_vertices > 10000000:
            return False, {}, f"unreasonable vertex count: {num_vertices}"

        format_count = u32(pos, "vertex format count")
        pos += 4

        if format_count == 0 or format_count > 64:
            return False, {}, f"unreasonable vertex format count: {format_count}"

        stride = 0

        for index in range(format_count):
            need(pos, 9, f"vertex format {index}")
            usage = u32(pos, f"vertex format {index} usage")
            subtype = u32(pos + 4, f"vertex format {index} subtype")
            bytes_per_element = raw[pos + 8]

            if bytes_per_element == 0 or bytes_per_element > 64:
                return (
                    False,
                    {},
                    f"invalid vertex element size {bytes_per_element} "
                    f"in format {index}"
                )

            stride += bytes_per_element
            pos += 9

        if stride <= 0 or stride > 1024:
            return False, {}, f"unreasonable vertex stride: {stride}"

        vertex_bytes = num_vertices * stride

        if pos + vertex_bytes > tgi_pos:
            return False, {}, "vertex data overruns declared TGI block"

        pos += vertex_bytes

        group_count = u32(pos, "face group count")
        pos += 4

        if group_count <= 1 or group_count > 256:
            return (
                False,
                {},
                f"not a valid legacy multi-group count: {group_count}"
            )

        total_face_points = 0
        highest_index = -1
        group_face_points = []

        for group_index in range(group_count):
            need(pos, 5, f"face group {group_index} header")

            point_size = raw[pos]
            point_count = u32(
                pos + 1,
                f"face group {group_index} point count"
            )
            pos += 5

            if point_size != 2:
                return (
                    False,
                    {},
                    f"face group {group_index} uses unsupported "
                    f"{point_size}-byte indices"
                )

            if point_count % 3 != 0:
                return (
                    False,
                    {},
                    f"face group {group_index} point count "
                    f"{point_count} is not triangle-aligned"
                )

            group_bytes = point_count * point_size
            need(pos, group_bytes, f"face group {group_index} indices")

            end = pos + group_bytes

            for offset in range(pos, end, 2):
                vertex_index = struct.unpack_from("<H", raw, offset)[0]

                if vertex_index >= num_vertices:
                    return (
                        False,
                        {},
                        f"face group {group_index} references vertex "
                        f"{vertex_index}, but vertex count is {num_vertices}"
                    )

                if vertex_index > highest_index:
                    highest_index = vertex_index

            total_face_points += point_count
            group_face_points.append(point_count)
            pos = end

        # Version 5 stores the skin controller index here.
        skin_index = i32(pos, "skin index")
        pos += 4

        bone_count = u32(pos, "bone hash count")
        pos += 4

        if bone_count > 4096:
            return False, {}, f"unreasonable bone hash count: {bone_count}"

        bone_bytes = bone_count * 4
        need(pos, bone_bytes, "bone hash list")
        pos += bone_bytes

        if pos != tgi_pos:
            return (
                False,
                {},
                f"legacy GEOM body does not land on declared TGI block "
                f"(body=0x{pos:X}, TGI=0x{tgi_pos:X})"
            )

        details = {
            "version": version,
            "vertices": num_vertices,
            "groups": group_count,
            "face_points": total_face_points,
            "faces": total_face_points // 3,
            "highest_index": highest_index,
            "skin_index": skin_index,
            "bone_count": bone_count,
            "bytes": len(raw),
            "merge_group": merge_group,
            "sort_order": sort_order,
            "group_face_points": group_face_points,
        }

        return True, details, ""

    except Exception as exc:
        return False, {}, f"{type(exc).__name__}: {exc}"



def get_geom_stats(g):

    vertices = 0
    faces = 0

    try:
        if not hasattr(g, "VertexData"):
            return 0, 0

        vertices = g.VertexData.Count
        faces = g.Faces.Count

        return vertices, faces

    except Exception:

        return 0, 0

def get_geom_poly_count(package, entry, resource_objects):

    highest_vertices = 0
    highest_faces = 0
    geom_id = f"0x{int(entry.Instance):016X}"

    try:

        geom = WrapperDealer.GetResource(
            0,
            package,
            entry
        )

        resource_objects.append(geom)

        chunks = geom["ChunkEntries"].Value

        for i in range(chunks.Count):

            g = chunks[i].RCOLBlock

            vertices, faces = get_geom_stats(g)

            if vertices > 0 and faces > 0:

                if faces > highest_faces:
                    highest_faces = faces
                    highest_vertices = vertices

        # GEOM DIAGNOSTIC: report successful parse details without adding
        # any new health rule. Existing parser failure remains the structural
        # GEOM health failure condition.
        try:
            geom_byte_count = len(geom.AsBytes)
        except Exception:
            try:
                geom_byte_count = int(entry.Filesize)
            except Exception:
                geom_byte_count = 0

        try:
            chunk_count = int(chunks.Count)
        except Exception:
            chunk_count = 0

        print(
            f"GEOM CHECK: {geom_id} | "
            f"Group=0x{int(entry.ResourceGroup):08X} | "
            f"Bytes={geom_byte_count:,} | Chunks={chunk_count} | "
            f"HighestFaces={highest_faces:,} | "
            f"Wrapper={type(geom).__name__} | Status=OK"
        )

        # A GEOM that parses successfully is structurally readable.
        # Do not mark tiny/zero-poly meshes broken here; legitimate CC can
        # contain very small GEOMs. Health failure is reserved for an actual
        # parser exception.
        return highest_vertices, highest_faces, True, ""


    except Exception as e:

        error_text = f"{type(e).__name__}: {e}"

        # LEGACY MULTI-GROUP GEOM COMPATIBILITY TEST
        #
        # s3pi's GEOM wrapper hard-codes the face-group count to 1.
        # Real old Sims 3 meshes can contain multiple structurally valid
        # face/index groups. Only invoke the manual fallback for that exact
        # known compatibility exception; every other parser exception remains
        # a normal structural GEOM failure.
        if "Expected number of face point sizes to be 1, read" in error_text:
            try:
                raw = _get_raw_resource_bytes(package, entry)

                (
                    legacy_ok,
                    legacy_details,
                    legacy_reason
                ) = _validate_legacy_multigroup_geom(raw)

                if legacy_ok:
                    print(
                        f"GEOM LEGACY MULTI-GROUP CHECK: {geom_id} | "
                        f"Group=0x{int(entry.ResourceGroup):08X} | "
                        f"Bytes={legacy_details['bytes']:,} | "
                        f"Groups={legacy_details['groups']} | "
                        f"Vertices={legacy_details['vertices']:,} | "
                        f"Faces={legacy_details['faces']:,} | "
                        f"HighestIndex={legacy_details['highest_index']:,} | "
                        f"Status=OK"
                    )

                    return (
                        legacy_details["vertices"],
                        legacy_details["faces"],
                        True,
                        ""
                    )

                error_text = (
                    f"{error_text} | "
                    f"Legacy multi-group validation failed: {legacy_reason}"
                )

            except Exception as legacy_exc:
                error_text = (
                    f"{error_text} | "
                    f"Legacy multi-group fallback error: "
                    f"{type(legacy_exc).__name__}: {legacy_exc}"
                )

        print(
            f"GEOM HEALTH FAILURE: {geom_id} | {error_text}"
        )

        return 0, 0, False, error_text
from datetime import datetime

scan_folder_name = "Health_Scan_Result_" + datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

scans_root = os.environ.get(
    "S3CC_SCANS_ROOT",
    r"D:\Documents\CCMagicCheck\Sims3_Sorted_CC"
)

scan_session_folder = os.path.join(
    scans_root,
    scan_folder_name
)

os.makedirs(
    scan_session_folder,
    exist_ok=True
)

LOG_OUTPUT = os.path.join(
    scan_session_folder,
    "Health_Scan_Log.txt"
)

_scanner_console_logger.start_log(LOG_OUTPUT)

OUTPUT = os.path.join(
    scan_session_folder,
    "Scan_Summary.xlsx"
)
MASTER_DATABASE = os.environ.get(
    "S3CC_MASTER_DATABASE",
    os.path.join(scans_root, "Sims3_CC_Master_Database.xlsx")
)
MOVE_FILES = os.environ.get("S3CC_MOVE_FILES", "1").strip().lower() not in ("0", "false", "no", "off")
ROUTING_PREVIEW = os.environ.get("S3CC_ROUTING_PREVIEW", "0").strip().lower() in ("1", "true", "yes", "on")
ROUTING_ACTIVE = MOVE_FILES or ROUTING_PREVIEW
SKIP_MDB_UPDATE = os.environ.get("S3CC_SKIP_MDB_UPDATE", "0").strip().lower() in ("1", "true", "yes", "on")
_routing_root_override = os.environ.get("S3CC_ROUTING_ROOT", "").strip()
ROUTING_ROOT = os.path.abspath(_routing_root_override) if _routing_root_override else ""
MOVE_FOLDER = ROUTING_ROOT if ROUTING_ROOT else scan_session_folder
_passthrough_root = ROUTING_ROOT if ROUTING_ROOT else scans_root
DUPLICATES_FOLDER = os.path.join(
    _passthrough_root,
    "DUPLICATES"
)
SIMS3PACK_FOLDER = os.path.join(
    _passthrough_root,
    "SIMS3PACK"
)
ARCHIVES_FOLDER = os.path.join(
    _passthrough_root,
    "ARCHIVES"
)
IMAGES_FOLDER = os.path.join(
    _passthrough_root,
    "IMAGES"
)

# Dedicated external preview storage.  Keep thumbnails outside the XLSX so the
# Master Database does not become enormous or slow.  The Image cells link to
# these files instead of embedding thousands of pictures.
THUMBNAILS_FOLDER = os.path.join(
    scans_root,
    "Thumbnails"
)
os.makedirs(THUMBNAILS_FOLDER, exist_ok=True)

# Big_Review scans are unlimited: when the manual/individual folder is empty,
# scan every available .package under Big_Review in the same run.
# Sims3_CC_for_Review still has priority for manual/individual scans.
PACKAGE_BATCH_LIMIT = None


IMG_TYPE = 0x00B2D882

# IMAG resources (separate from DDS IMG):
# 0x2F7D0002 = JPEG
# 0x2F7D0004 = PNG
# 0x2F7D0006 = TGA
IMAG_JPEG_TYPE = 0x2F7D0002
IMAG_PNG_TYPE = 0x2F7D0004
IMAG_TGA_TYPE = 0x2F7D0006
IMAG_TYPES = {
    IMAG_JPEG_TYPE: "JPEG",
    IMAG_PNG_TYPE: "PNG",
    IMAG_TGA_TYPE: "TGA",
}

CASP_TYPES = {
    0x034AEECB
}

GEOM_TYPE = 0x015A1849

VPXY_TYPE = 0x736884F1

BONE_TYPE = 0x00AE6C67
BOND_TYPE = 0x0355E0A6
RIG_TYPE = 0x8EAF13DE
JAZZ_TYPE = 0x02D5DF13
S3SA_TYPE = 0x073FAA07
LITE_TYPE = 0x03B4C61D
CLIP_TYPE = 0x6B20C4F3
MATD_TYPE = 0x01D0E75D
MTST_TYPE = 0x02019972
MDLR_TYPE = 0xCF9A4ACE
CWAL_TYPE = 0x515CA4CD
CFEN_TYPE = 0x0418FE2A
CTPT_TYPE = 0x04ED4BB2
NGMP_TYPE = 0xF3A38370
CRMT_TYPE = 0xF1EDBD86
SPT2_TYPE = 0x021D7E8C
SPT_TYPE = 0x00B552EA
CRAL_TYPE = 0x04C58103
LAYO_TYPE = 0x025C95B6
HINF_TYPE = 0x063261DA
TWOARY_TYPE = 0x05FF6BA4
COMPLATE_TYPE = 0x044AE110
TYPE_63A33EA7 = 0x63A33EA7
FIREPLACE_TYPE = 0x04F3CC01
THUM_TYPE = 0x626F60CD
THUM_AD366F95_TYPE = 0xAD366F95  # additional Sims 3 THUM image resource type
THUM_626F60CC_TYPE = 0x626F60CC
THUM_05B1B524_TYPE = 0x05B1B524
THUM_05B1B525_TYPE = 0x05B1B525
THUM_05B1B526_TYPE = 0x05B1B526
THUM_2653E3C9_TYPE = 0x2653E3C9
THUM_FCEAB65B_TYPE = 0xFCEAB65B
ICON_D84E7FC5_TYPE = 0xD84E7FC5
ICON_D84E7FC6_TYPE = 0xD84E7FC6
ICON_D84E7FC7_TYPE = 0xD84E7FC7
TWNI_TYPE = 0x0668F635
THUM_2D4284F1_TYPE = 0x2D4284F1
MANIFEST_TYPE = 0x73E93EEB
TYPE_8F78F74B = 0x8F78F74B
PNG_626F60CE_TYPE = 0x626F60CE
ICON_TYPE = 0x2E75C765
PNG_0580A2B5_TYPE = 0x0580A2B5
THUM_0580A2B4_TYPE = 0x0580A2B4
THUM_0580A2B6_TYPE = 0x0580A2B6
SNAP_0580A2CD_TYPE = 0x0580A2CD
SNAP_0580A2CF_TYPE = 0x0580A2CF
ICON_2E75C764_TYPE = 0x2E75C764
ICON_2E75C766_TYPE = 0x2E75C766
ICON_2E75C767_TYPE = 0x2E75C767

# -------------------------
# EXTERNAL THUMBNAIL EXTRACTION (TEST)
# -------------------------
# Only dedicated Sims 3 preview/icon/snapshot resource families qualify.
# Ordinary DDS IMG textures are deliberately excluded: using them as fallback
# can show diffuse/normal/specular textures rather than the CC itself.
# Higher values win first; pixel area and payload size break ties.
THUMBNAIL_PREVIEW_PRIORITY = {
    SNAP_0580A2CF_TYPE: 600,
    SNAP_0580A2CD_TYPE: 590,
    ICON_D84E7FC7_TYPE: 580,
    ICON_D84E7FC6_TYPE: 570,
    ICON_D84E7FC5_TYPE: 560,
    THUM_TYPE: 550,
    PNG_626F60CE_TYPE: 545,
    THUM_AD366F95_TYPE: 540,
    THUM_626F60CC_TYPE: 535,
    ICON_TYPE: 520,
    ICON_2E75C767_TYPE: 515,
    ICON_2E75C766_TYPE: 510,
    ICON_2E75C764_TYPE: 505,
    PNG_0580A2B5_TYPE: 490,
    THUM_0580A2B6_TYPE: 485,
    THUM_0580A2B4_TYPE: 475,
    THUM_05B1B526_TYPE: 470,
    THUM_05B1B525_TYPE: 465,
    THUM_05B1B524_TYPE: 460,
    THUM_2653E3C9_TYPE: 455,
    THUM_2D4284F1_TYPE: 450,
    THUM_FCEAB65B_TYPE: 445,
}


def _png_has_terminal_iend(raw):
    return bool(
        raw
        and len(raw) >= 12
        and raw[-12:-8] == b"\x00\x00\x00\x00"
        and raw[-8:-4] == b"IEND"
        and raw[-4:] == b"\xAE\x42\x60\x82"
    )


def _validate_png_missing_iend_only(raw):
    """
    Return (True, width, height) only when a PNG is structurally complete and
    decodable except for the final 12-byte IEND chunk.

    This is deliberately strict because it is used to avoid a THUM false
    positive. It validates chunk boundaries/CRCs, IHDR fields, IDAT zlib
    completion and the expected decompressed scanline byte count. A genuinely
    truncated/corrupt PNG still fails health.
    """
    png_sig = b"\x89PNG\r\n\x1a\n"
    if not raw or not raw.startswith(png_sig) or len(raw) < 33:
        return False, 0, 0
    if _png_has_terminal_iend(raw):
        return False, 0, 0

    pos = 8
    chunks = []
    idat_parts = []
    saw_ihdr = False
    saw_idat = False
    idat_closed = False
    width = height = 0
    bit_depth = color_type = compression = filter_method = interlace = None

    while pos < len(raw):
        if len(raw) - pos < 12:
            return False, 0, 0

        length = int.from_bytes(raw[pos:pos + 4], "big", signed=False)
        ctype = raw[pos + 4:pos + 8]
        end = pos + 12 + length
        if end > len(raw):
            return False, 0, 0

        data = raw[pos + 8:pos + 8 + length]
        stored_crc = int.from_bytes(raw[pos + 8 + length:end], "big", signed=False)
        calc_crc = zlib.crc32(ctype)
        calc_crc = zlib.crc32(data, calc_crc) & 0xFFFFFFFF
        if stored_crc != calc_crc:
            return False, 0, 0

        if not all((65 <= b <= 90) or (97 <= b <= 122) for b in ctype):
            return False, 0, 0

        if not chunks:
            if ctype != b"IHDR" or length != 13:
                return False, 0, 0

        if ctype == b"IHDR":
            if saw_ihdr or length != 13 or chunks:
                return False, 0, 0
            saw_ihdr = True
            width = int.from_bytes(data[0:4], "big", signed=False)
            height = int.from_bytes(data[4:8], "big", signed=False)
            bit_depth = data[8]
            color_type = data[9]
            compression = data[10]
            filter_method = data[11]
            interlace = data[12]
            if width <= 0 or height <= 0:
                return False, 0, 0
            valid_depths = {
                0: {1, 2, 4, 8, 16},
                2: {8, 16},
                3: {1, 2, 4, 8},
                4: {8, 16},
                6: {8, 16},
            }
            if color_type not in valid_depths or bit_depth not in valid_depths[color_type]:
                return False, 0, 0
            if compression != 0 or filter_method != 0 or interlace not in (0, 1):
                return False, 0, 0

        elif ctype == b"IDAT":
            if not saw_ihdr or idat_closed:
                return False, 0, 0
            saw_idat = True
            idat_parts.append(data)

        elif ctype == b"IEND":
            # This helper is ONLY for a PNG whose sole defect is absent IEND.
            return False, 0, 0

        else:
            if saw_idat:
                idat_closed = True

        chunks.append(ctype)
        pos = end

    if pos != len(raw) or not saw_ihdr or not saw_idat or not idat_parts:
        return False, 0, 0

    # Indexed-colour PNGs require a PLTE before IDAT.
    if color_type == 3:
        try:
            if chunks.index(b"PLTE") > chunks.index(b"IDAT"):
                return False, 0, 0
        except ValueError:
            return False, 0, 0

    try:
        inflater = zlib.decompressobj()
        decoded = inflater.decompress(b"".join(idat_parts))
        decoded += inflater.flush()
        if not inflater.eof or inflater.unused_data or inflater.unconsumed_tail:
            return False, 0, 0
    except Exception:
        return False, 0, 0

    channels = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[color_type]

    def row_bytes(pixel_width):
        return (pixel_width * channels * bit_depth + 7) // 8

    if interlace == 0:
        expected = height * (1 + row_bytes(width))
    else:
        expected = 0
        # Adam7: x-start, y-start, x-step, y-step.
        for xs, ys, dx, dy in (
            (0, 0, 8, 8),
            (4, 0, 8, 8),
            (0, 4, 4, 8),
            (2, 0, 4, 4),
            (0, 2, 2, 4),
            (1, 0, 2, 2),
            (0, 1, 1, 2),
        ):
            if width <= xs or height <= ys:
                continue
            pw = (width - xs + dx - 1) // dx
            ph = (height - ys + dy - 1) // dy
            expected += ph * (1 + row_bytes(pw))

    if len(decoded) != expected:
        return False, 0, 0

    return True, width, height


def _repair_png_missing_iend(raw):
    ok, _width, _height = _validate_png_missing_iend_only(raw)
    if not ok:
        return raw
    # Canonical empty IEND chunk: length + type + CRC.
    return raw + b"\x00\x00\x00\x00IEND\xAE\x42\x60\x82"


def _thumbnail_payload_info(raw):
    """Return (extension, width, height) only for a complete known image."""
    if not raw:
        return None

    png_sig = b"\x89PNG\r\n\x1a\n"
    jpeg_soi = b"\xFF\xD8"
    jpeg_eoi = b"\xFF\xD9"

    if raw.startswith(png_sig):
        if len(raw) < 24:
            return None
        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)
        if width <= 0 or height <= 0:
            return None
        if not _png_has_terminal_iend(raw):
            recoverable, repaired_width, repaired_height = _validate_png_missing_iend_only(raw)
            if not recoverable:
                return None
            width = repaired_width
            height = repaired_height
        return "png", width, height

    if raw.startswith(jpeg_soi):
        if len(raw) < 4 or not raw.endswith(jpeg_eoi):
            return None
        width = None
        height = None
        try:
            width, height = _diagnose_jpeg_dimensions_626f60cc(raw)
        except Exception:
            pass
        return "jpg", int(width or 0), int(height or 0)

    if raw.startswith(b"BM"):
        try:
            ok, _reason, width, height = _validate_known_bmp_thumbnail(raw)
            if ok:
                return "bmp", int(width or 0), int(height or 0)
        except Exception:
            pass

    return None



def _validate_626f60cd_preset_metadata_variant(raw):
    """
    Recognise the real-world non-image 0x626F60CD variant used by some CAS
    packages for binary preset metadata.

    Evidence from [sls]eyebrows82.package:
      - the DBPF contains ordinary PNG 0x626F60CD/0x626F60CE thumbnails AND a
        second 0x626F60CD resource whose decompressed payload begins with a
        short binary header followed by UTF-16LE <preset><complate ...> XML.
      - the XML block is well-formed and describes a CAS preset/complate.

    This validator is deliberately narrow. It does not accept arbitrary
    non-image bytes merely because the resource type is THUM.
    """
    if not raw or len(raw) < 32:
        return False

    preset_open = "<preset>".encode("utf-16-le")
    preset_close = "</preset>".encode("utf-16-le")

    # Known samples place the binary header before the XML. Keep the search
    # close to the front so random embedded text later in a damaged resource
    # cannot turn a true corruption into OK.
    start = raw.find(preset_open, 0, min(len(raw), 96))
    if start < 8:
        return False

    close = raw.find(preset_close, start)
    if close < 0:
        return False

    xml_end = close + len(preset_close)
    xml_raw = raw[start:xml_end]
    if len(xml_raw) % 2:
        return False

    try:
        text = xml_raw.decode("utf-16-le")
        root = ET.fromstring(text)
    except Exception:
        return False

    def local_name(tag):
        return str(tag).rsplit("}", 1)[-1].lower()

    if local_name(root.tag) != "preset":
        return False

    # Require the CAS preset structure rather than accepting any XML root.
    if not any(local_name(node.tag) == "complate" for node in root.iter()):
        return False

    return True

def _existing_thumbnail_for_hash(package_hash):
    key = str(package_hash or "").strip().lower()
    if not key:
        return ""
    for ext in ("png", "jpg", "bmp"):
        candidate = os.path.join(THUMBNAILS_FOLDER, f"{key}.{ext}")
        if os.path.isfile(candidate):
            return candidate
    return ""


def extract_best_package_thumbnail(filename, package_hash):
    """
    Extract the best dedicated preview resource from one valid Sims 3 package.

    Returns (path, status), where status is one of:
      extracted / existing / none / error

    This is deliberately separate from all health check_* functions.
    """
    existing = _existing_thumbnail_for_hash(package_hash)
    if existing:
        return existing, "existing"

    package = None
    resources = None

    try:
        package = Package.OpenPackage(0, filename)
        resources = package.GetResourceList
        best = None

        for i in range(resources.Count):
            entry = resources[i]
            rtype = int(entry.ResourceType)
            priority = THUMBNAIL_PREVIEW_PRIORITY.get(rtype)
            if priority is None:
                continue

            resource = None
            try:
                resource = WrapperDealer.GetResource(0, package, entry)
                raw = bytes(resource.AsBytes)
                info = _thumbnail_payload_info(raw)
                if info is None:
                    continue

                ext, width, height = info
                pixel_area = int(width or 0) * int(height or 0)
                score = (priority, pixel_area, len(raw))

                if best is None or score > best[0]:
                    best = (
                        score,
                        ext,
                        raw,
                        rtype,
                        int(entry.ResourceGroup),
                        int(entry.Instance),
                        width,
                        height,
                    )

            except Exception:
                # Extraction must never change health/classification outcomes.
                # A bad candidate is simply ignored; its dedicated health check
                # remains the authoritative place for corruption reporting.
                continue
            finally:
                try:
                    if resource is not None and hasattr(resource, "Stream"):
                        stream = resource.Stream
                        if stream is not None:
                            stream.Dispose()
                            stream.Close()
                except Exception:
                    pass

        if best is None:
            return "", "none"

        _score, ext, raw, rtype, group, instance, width, height = best
        hash_key = str(package_hash or "").strip().lower()
        if not hash_key:
            hash_key = hashlib.sha256(os.path.abspath(filename).encode("utf-8", "ignore")).hexdigest()

        output_path = os.path.join(THUMBNAILS_FOLDER, f"{hash_key}.{ext}")
        if ext == "png" and not _png_has_terminal_iend(raw):
            raw = _repair_png_missing_iend(raw)
        with open(output_path, "wb") as preview_file:
            preview_file.write(raw)

        if DEBUG:
            print(
                "THUMBNAIL EXTRACTED:",
                os.path.basename(filename),
                "| Type=0x%08X" % rtype,
                "| Group=0x%08X" % group,
                "| Instance=0x%016X" % instance,
                "| Size=%sx%s" % (width or "?", height or "?"),
                "|",
                output_path,
            )

        return output_path, "extracted"

    except Exception as e:
        print(
            "THUMBNAIL EXTRACTION ERROR:",
            os.path.basename(filename),
            "|",
            f"{type(e).__name__}: {e}"
        )
        return "", "error"

    finally:
        resources = None
        try:
            if package is not None:
                Package.ClosePackage(0, package)
        except Exception:
            pass
        package = None


def _thumbnail_file_uri(path):
    # Excel/LibreOffice-friendly absolute local-file URI.
    text = os.path.abspath(path).replace("\\", "/")
    from urllib.parse import quote
    return "file:///" + quote(text, safe="/:")


def _set_thumbnail_link(cell, thumbnail_path):
    """Set a clickable Image cell without overwriting manual text."""
    if not thumbnail_path:
        return

    existing_text = str(cell.value or "").strip()
    if existing_text and existing_text not in ("View thumbnail", "Thumbnail"):
        return

    cell.value = "View thumbnail"
    cell.hyperlink = _thumbnail_file_uri(thumbnail_path)
    cell.font = Font(color="0563C1", underline="single")


BGEO_TYPE = 0x067CAA11
BBLN_TYPE = 0x062C8204
FACE_TYPE = 0x0358B08A
FBLN_TYPE = 0xB52F5055

TXTC_TYPE = 0x033A1435
TXTF_TYPE = 0x0341ACC9
PTRN_TYPE = 0xD4D9FBE5
XML_TYPE = 0x0333406C
STORE_INGREDIENTS_TYPE = 0xE5105066  # Sims 3 Store Ingredients XML resource; classification-only here
INI_TYPE = 0x1F886EAD  # _INI / Sims 3 configuration resource
CSS_TYPE = 0x025C90A6  # _CSS / UI text styles
TTF_TYPE = 0x062E9EE0  # _TTF / embedded TrueType font
LDES_TYPE = 0xD063545B  # LDES / Lot Description / lot file
ITUN_TYPE = 0x03B33DDF

MODL_TYPE = 0x01661233
MLOD_TYPE = 0x01D10F34
RSLT_TYPE = 0xD3044521
FTPT_TYPE = 0xD382BF57
OBJK_TYPE = 0x02DC343F
OBJD_TYPE = 0x319E4F1D

PRESET_TYPE = 0x051DF2DD
TONE_TYPE = 0x0354796A
STBL_TYPE = 0x220557DA
COLL_TYPE = 0x0C07456D  # COLL / Sims 3 Collection file resource

# Sim / household package resources
SIMO_TYPE = 0x025ED6F4
SIME_TYPE = 0x04F88964  # Sim Export / saved-Sim information (classification-only)
FAMD_TYPE = 0x062853A8

# Name Map / _KEY
KEY_TYPE = 0x0166038C  # historical scanner name
NMAP_TYPE = KEY_TYPE  # actual resource: NMAP / Name Map

CAS_TYPE_NAMES = {

    0x00000001: "Hair",
    0x00000004: "Body",
    0x00000005: "Top",
    0x00000006: "Bottom",
    0x00000007: "Shoes",

    0x00000009: "Accessory",
    0x0000000B: "Earrings",
    0x0000000C: "Glasses_F",
    0x0000000D: "Bracelets",
    0x0000000E: "Ring_Lt",
    0x00000010: "Beard",

    0x00000011: "Lipstick",
    0x00000012: "Eyeshadow",
    0x00000013: "Eyeliner",
    0x00000014: "Blush",
    0x00000015: "Makeup",

    0x00000016: "Eyebrow",
    0x00000018: "Glove",
    0x00000019: "Socks",
    0x0000001B: "Moles",
    0x0000001E: "Earrings",
    0x0000001F: "Earrings",
    0x00000020: "ArmBand",
    0x00000022: "Tattoo",

    # Pets EP CASP ClothingType 0x35 = PeltLayer (pet coat marking).
    0x00000035: "Pet Marking"
}


def get_texture_size(
    package,
    entry,
    resource_objects,
    broken_img_ids=None,
    broken_img_errors=None
):

    img_id = f"0x{int(entry.Instance):016X}"

    if broken_img_ids is None:
        broken_img_ids = []

    if broken_img_errors is None:
        broken_img_errors = []

    def record_img_failure(reason):
        if img_id not in broken_img_ids:
            broken_img_ids.append(img_id)

        broken_img_errors.append(f"{img_id}: {reason}")
        print(f"IMG HEALTH FAILURE: {img_id} | {reason}")

    try:

        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )

        resource_objects.append(resource)

        data = resource.AsBytes
        data_len = len(data)

        # IMG is a container type and not every IMG resource is DDS.
        # Only data that begins with the DDS magic is validated as DDS.
        # Non-DDS IMG resources are left alone.
        if data_len < 4:
            return None

        is_dds = (
            data[0] == 0x44
            and data[1] == 0x44
            and data[2] == 0x53
            and data[3] == 0x20
        )

        if not is_dds:
            return None

        # A standard DDS header is 128 bytes including the 4-byte magic.
        if data_len < 128:
            record_img_failure(
                f"DDS TRUNCATED ({data_len:,} bytes; minimum header is 128 bytes)"
            )
            return None

        header_size = (
            data[4]
            | (data[5] << 8)
            | (data[6] << 16)
            | (data[7] << 24)
        )

        if header_size != 124:
            record_img_failure(
                f"DDS BAD HEADER (header size {header_size}; expected 124)"
            )
            return None

        height = (
            data[12]
            | (data[13] << 8)
            | (data[14] << 16)
            | (data[15] << 24)
        )
        width = (
            data[16]
            | (data[17] << 8)
            | (data[18] << 16)
            | (data[19] << 24)
        )

        if width <= 0 or height <= 0:
            record_img_failure(
                f"DDS BAD DIMENSIONS ({width}x{height})"
            )
            return None

        try:
            group_id = f"0x{int(entry.ResourceGroup):08X}"
        except Exception:
            group_id = "UNKNOWN"

        print(
            f"IMG DDS CHECK: {img_id} | Group={group_id} | Bytes={data_len:,} | "
            f"Size={width}x{height} | Wrapper={type(resource).__name__} | Status=OK"
        )

        return width, height

    except Exception as e:

        record_img_failure(
            f"RESOURCE PARSE FAILED - {type(e).__name__}: {e}"
        )

        return None

def detect_package_format(filename):

    raw_name = os.path.splitext(
        os.path.basename(filename)
    )[0]

    name_parts = []

    for index, char in enumerate(raw_name):

        if (
            index > 0
            and char.isupper()
            and raw_name[index - 1].islower()
        ):
            name_parts.append(" ")

        name_parts.append(char)

    name = "".join(name_parts).lower()

    clean_name = "".join(
        char if char.isalnum() else " "
        for char in name
    )

    words = set(clean_name.split())

    if (
        "merged" in words
        or "merge" in words
    ):
        return "Merged Collection"

    elif (
        "collection" in words
        or "set" in words
        or "bundle" in words
        or "pack" in words
        or "suite" in words
    ):
        return "Collection"

    else:
        return "Single Item"
        
def detect_subtype(filename, cc_type, package_format, explicit_subtype=""):

    if explicit_subtype:
        return str(explicit_subtype).strip()

    raw_name = os.path.splitext(
        os.path.basename(filename)
    )[0]

    name_parts = []

    for index, char in enumerate(raw_name):

        if (
            index > 0
            and char.isupper()
            and raw_name[index - 1].islower()
        ):
            name_parts.append(" ")

        name_parts.append(char)

    name = "".join(
        char.lower() if char.isalnum() else " "
        for char in "".join(name_parts)
    )

    name_tokens = set(name.split())

    # -------------------------
    # BUILD / BUY SUBTYPES

    if cc_type in [
        "Build/Buy",
        "Pattern Collection"
    ]:

        if package_format == "Merged Collection":
            return "Merged Collection"

        elif "pattern" in name_tokens:
            return "Pattern"

        elif "wallpaper" in name_tokens:
            return "Wallpaper"

        elif (
            "floor" in name_tokens
            or "floortile" in name_tokens
        ):
            return "Floor Tile"

        elif (
            "light" in name_tokens
            or "lighting" in name_tokens
        ):
            return "Lighting"

        elif "wall" in name_tokens:
            return "Wall"

        elif "rug" in name_tokens:
            return "Rug"

        elif "mirror" in name_tokens:
            return "Mirror"

        elif "poster" in name_tokens:
            return "Poster"

        elif "painting" in name_tokens:
            return "Painting"

    return ""


def detect_special_cc(
    filename,
    casp,
    geom,
    presets,
    cc_type="",
    objd=0,
    objk=0,
    modl=0,
    mlod=0,
    simo=0,
    famd=0
):


    # Strong Sim / household package evidence
    if simo > 0 and famd > 0:
        return "Custom Sim"

    # -------------------------
    # BUILD / BUY DETECTION

    # A real PTRN/pattern classification is stronger than generic object-model
    # evidence. Pattern packages can legitimately contain MODL/MLOD/OBJD/OBJK
    # helpers, so do not collapse them back to generic Build/Buy.
    if cc_type == "Pattern":
        return None

    if (
        casp == 0
        and (
            objd > 0
            or objk > 0
            or modl > 0
            or mlod > 0
        )
    ):
        return "Build/Buy"

    # Never override recognised CAS types
    if (
        cc_type in CAS_TYPE_NAMES.values()
        or cc_type == "Mixed CAS"
    ):
        return None


    # Presets
    # Real PRESET resource evidence outranks filename guesses.
    if presets > 0 and casp == 0:

        raw_preset_name = os.path.splitext(
            os.path.basename(filename)
        )[0]

        preset_parts = []

        for index, char in enumerate(raw_preset_name):

            if (
                index > 0
                and char.isupper()
                and raw_preset_name[index - 1].islower()
            ):
                preset_parts.append(" ")

            preset_parts.append(char)

        preset_name = "".join(
            char.lower() if char.isalnum() else " "
            for char in "".join(preset_parts)
        )

        preset_tokens = set(
            preset_name.split()
        )

        types = []

        if (
            "nose" in preset_tokens
            or "noses" in preset_tokens
        ):
            types.append("Nose Preset")
        if (
            "lip" in preset_tokens
            or "lips" in preset_tokens
        ):
            types.append("Lip Preset")

        if (
            "eye" in preset_tokens
            or "eyes" in preset_tokens
        ):
            types.append("Eye Preset")

        if not types:
            types.append("Face Preset")

        return ", ".join(sorted(types))


    # Never let filename heuristics override
    # established resource-based classifications.
    if cc_type in [
        "Object",
        "Morph Slider"
    ]:
        return None


    # Skin detection
    raw_skin_name = os.path.splitext(
        os.path.basename(filename)
    )[0]

    skin_parts = []

    for index, char in enumerate(raw_skin_name):

        if (
            index > 0
            and char.isupper()
            and raw_skin_name[index - 1].islower()
        ):
            skin_parts.append(" ")

        skin_parts.append(char)

    skin_name = "".join(
        char.lower() if char.isalnum() else " "
        for char in "".join(skin_parts)
    )

    skin_tokens = set(skin_name.split())

    if (
        casp == 0
        and geom == 0
        and (
            "skin" in skin_tokens
            or "skinblend" in skin_tokens
            or "defaultskin" in skin_tokens
        )
    ):
        return "Skin"


    return None
    
def calculate_risk(hq_warning, poly_warning, geom_bytes, fourk, vpxy, cc_type):
    # Confirmed Custom Sims are containers, not individual CAS items.
    # Performance-risk scoring is suppressed, but structural health checks
    # are completely separate and still run normally.
    if cc_type == "Custom Sim":
        return "0/10", "Safe"

    score = 0


    # Texture risk
    if hq_warning == "4K TEXTURE":
        score += 5

        if fourk >= 3:
            score += 1

    elif hq_warning == "2048 HIGH":
        score += 2


    # Mesh risk
    if poly_warning == "VERY HEAVY":
        score += 5

    elif poly_warning == "HEAVY":
        score += 3


    # GEOM size risk

    geom_mb = geom_bytes / (1024 * 1024)

    if geom_mb >= 2:
        score += 2

    elif geom_mb >= 1:
        score += 1

    # VPXY Complexity Risk

    if vpxy >= 20:
        score += 2

    elif vpxy >= 10:
        score += 1    

    # Cap score
    if score > 10:
        score = 10


    # Final label
    if score <= 1:
        level = "Safe"

    elif score <= 3:
        level = "Moderate"

    elif score <= 5:
        level = "High"

    else:
        level = "Extreme"


    return f"{score}/10", level
    
def normalize_package_name(name):

    name = os.path.splitext(name)[0]

    name = name.lower()

    name = name.replace(" ", "_")

    name = name.replace("-", "_")

    name = name.replace("[", "")
    name = name.replace("]", "")

    return name
    
def get_risk_reasons(
    hq_warning,
    poly_warning,
    geom_bytes,
    fourk,
    highest_poly,
    cc_type,
    vpxy
):

    reasons = []

    if cc_type == "Custom Sim":
        return "None"

    if hq_warning == "4K TEXTURE":
        reasons.append("4K Texture")

    elif hq_warning == "2048 HIGH":
        reasons.append("High Resolution Texture")

    if poly_warning == "VERY HEAVY":
        reasons.append(
            f"{cc_type} mesh: {highest_poly:,} polys (Very Heavy)"
        )

    elif poly_warning == "HEAVY":
        reasons.append(
            f"{cc_type} mesh: {highest_poly:,} polys (Heavy)"
        )

    geom_mb = geom_bytes / (1024 * 1024)

    if geom_mb >= 2:
        reasons.append(
            f"GEOM size: {geom_mb:.2f} MB (Very Heavy)"
        )

    elif geom_mb >= 1:
        reasons.append(
            f"GEOM size: {geom_mb:.2f} MB (Heavy)"
        )

    elif geom_mb >= 0.5:
        reasons.append(
            f"GEOM size: {geom_mb:.2f} MB (Moderate)"
        )
    # VPXY Complexity

    if vpxy >= 20:
        reasons.append(f"Very High VPXY Count: {vpxy}")

    elif vpxy >= 10:
        reasons.append(f"High VPXY Count: {vpxy}")

    if fourk > 1:
        reasons.append(f"{fourk}x 4K Textures")

    return ", ".join(reasons) if reasons else "None"


def get_poly_warning(poly_count, cc_type):

    # Shoes - very strict
    if cc_type == "Shoes":

        if poly_count >= 8000:
            return "VERY HEAVY"

        elif poly_count >= 4000:
            return "HEAVY"

        return "Normal"


    # Tops - stricter because they combine with bottoms
    elif cc_type == "Top":

        if poly_count >= 12000:
            return "VERY HEAVY"

        elif poly_count >= 7000:
            return "HEAVY"

        return "Normal"


    # Bottoms - stricter because they combine with tops
    elif cc_type == "Bottom":

        if poly_count >= 10000:
            return "VERY HEAVY"

        elif poly_count >= 5000:
            return "HEAVY"

        return "Normal"


    # Full body outfits
    elif cc_type == "Body":

        if poly_count >= 15000:
            return "VERY HEAVY"

        elif poly_count >= 7500:
            return "HEAVY"

        return "Normal"


    # Hair gets more allowance
    elif cc_type == "Hair":

        if poly_count >= 30000:
            return "VERY HEAVY"

        elif poly_count >= 20000:
            return "HEAVY"

        return "Normal"


    # Accessories / small CAS parts
    elif cc_type in [
        "Accessory",
        "Earrings",
        "Glasses_F",
        "Ring_Lt",
        "Bracelets",
        "Glove",
        "Socks",
        "ArmBand",
        "Beard",
        "Moles",
        "Tattoo",
        "Makeup",
        "Eyebrow",
        "Lipstick",
        "Blush",
        "Eyeshadow",
        "Eyeliner"
    ]:

        if poly_count >= 5000:
            return "VERY HEAVY"

        elif poly_count >= 3000:
            return "HEAVY"

        return "Normal"


    # Skin, confirmed Sims and sliders are not judged by mesh poly thresholds
    elif cc_type in [
        "Skin",
        "Custom Sim",
        "Morph Slider"
    ]:

        return "Normal"

    # Default fallback
    if poly_count >= 30000:
        return "VERY HEAVY"

    elif poly_count >= 15000:
        return "HEAVY"

    return "Normal"
    
def format_size(value):

    if value >= 1024 * 1024:
        return f"{value/(1024*1024):.2f} MB"

    elif value >= 1024:
        return f"{value/1024:.1f} KB"

    return f"{value} bytes"
    
def get_column_map(ws):

    columns = {}

    for cell in ws[1]:

        if cell.value:

            columns[cell.value] = cell.column

    return columns 


def run_thumbnail_backfill(scan_root):
    """
    Fast thumbnail-only refresh for packages already represented in the Master DB.

    Deliberately does NOT call scan_package(), health checks, classification,
    EA comparison, risk analysis, CASP extraction, routing, or report generation.
    SHA-256 is the identity gate.  Only Image cells are changed, and existing
    user-entered Image text is preserved by _set_thumbnail_link().
    """
    import time

    started = time.time()
    package_paths = []
    for root, dirs, files in os.walk(scan_root):
        for file in files:
            if file.lower().endswith(".package"):
                package_paths.append(os.path.join(root, file))

    package_paths.sort(key=lambda p: os.path.normcase(os.path.abspath(p)))

    print("")
    print("THUMBNAIL BACKFILL MODE")
    print("  Full health/classification/risk scan: SKIPPED")
    print("  File moving/routing: SKIPPED")
    print("  Master DB changes: Image cells only")
    print("  Identity match: exact SHA256 Package Hash")
    print("TOTAL PACKAGES TO BACKFILL:", len(package_paths))
    print("")

    if not package_paths:
        print("THUMBNAIL BACKFILL: no .package files found.")
        return

    if not os.path.exists(MASTER_DATABASE):
        print("THUMBNAIL BACKFILL ERROR: Master Database does not exist:")
        print(MASTER_DATABASE)
        return

    print("THUMBNAIL BACKFILL: loading Master Database...")
    wb = openpyxl.load_workbook(MASTER_DATABASE)

    target_sheets = []
    known_hashes = set()

    # Image is an everyday review field on CC Master Database only.
    # The Packages technical sheet intentionally has no Image column, so
    # thumbnail backfill must not warn about or require one there.
    for sheet_name in ("CC Master Database",):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        cmap = get_column_map(ws)
        hash_col = cmap.get("Package Hash")
        image_col = cmap.get("Image")
        if not hash_col or not image_col:
            missing = []
            if not hash_col:
                missing.append("Package Hash")
            if not image_col:
                missing.append("Image")
            print(
                f"THUMBNAIL BACKFILL ERROR: {sheet_name} missing required "
                f"column(s): {', '.join(missing)}"
            )
            continue

        hash_rows = {}
        for row_idx in range(2, ws.max_row + 1):
            value = str(ws.cell(row_idx, hash_col).value or "").strip().lower()
            if len(value) == 64:
                hash_rows.setdefault(value, []).append(row_idx)
                known_hashes.add(value)

        target_sheets.append((ws, image_col, hash_rows))
        print(
            f"THUMBNAIL BACKFILL: indexed {sheet_name}: "
            f"{len(hash_rows):,} SHA256 hash(es)"
        )

    if not target_sheets:
        wb.close()
        print("THUMBNAIL BACKFILL ERROR: no compatible MDB sheet was found.")
        return

    extracted = 0
    reused = 0
    no_preview = 0
    errors = 0
    not_in_mdb = 0
    foreign = 0
    empty_or_hash_error = 0
    matched_packages = 0
    image_cells_updated = 0
    duplicate_hash_inputs = 0

    # Cache by hash so two byte-identical files supplied to the backfill folder
    # do not reopen/decompress the same package twice.
    processed_hashes = {}

    for index, path in enumerate(package_paths, 1):
        filename = os.path.basename(path)
        if index == 1 or index % 100 == 0 or index == len(package_paths):
            print(
                f"THUMBNAIL BACKFILL PROGRESS: "
                f"{index:,}/{len(package_paths):,}"
            )

        try:
            if os.path.getsize(path) == 0:
                empty_or_hash_error += 1
                print("  SKIPPED: empty package file")
                continue

            foreign_game, dbpf_major, dbpf_minor = _known_foreign_sims_game(path)
            if foreign_game:
                foreign += 1
                print(
                    f"  SKIPPED: {foreign_game} package "
                    f"(DBPF {dbpf_major}.{dbpf_minor})"
                )
                continue

            package_hash = calculate_file_hash(path)
            hash_key = str(package_hash or "").strip().lower()
            if len(hash_key) != 64:
                empty_or_hash_error += 1
                print("  SKIPPED: SHA256 could not be calculated")
                continue

            if hash_key not in known_hashes:
                not_in_mdb += 1
                print("  NOT IN MDB: thumbnail-only mode will not create a record")
                continue

            matched_packages += 1

            if hash_key in processed_hashes:
                thumbnail_path, thumbnail_status = processed_hashes[hash_key]
                duplicate_hash_inputs += 1
                print("  REUSED BACKFILL RESULT: identical SHA256 already processed")
            else:
                thumbnail_path, thumbnail_status = extract_best_package_thumbnail(
                    path,
                    package_hash
                )
                processed_hashes[hash_key] = (thumbnail_path, thumbnail_status)

                if thumbnail_status == "extracted":
                    extracted += 1
                elif thumbnail_status == "existing":
                    reused += 1
                elif thumbnail_status == "none":
                    no_preview += 1
                elif thumbnail_status == "error":
                    errors += 1

            if not thumbnail_path:
                continue

            for ws, image_col, hash_rows in target_sheets:
                for row_idx in hash_rows.get(hash_key, []):
                    cell = ws.cell(row_idx, image_col)
                    before_value = cell.value
                    before_link = cell.hyperlink.target if cell.hyperlink else None
                    _set_thumbnail_link(cell, thumbnail_path)
                    after_link = cell.hyperlink.target if cell.hyperlink else None
                    if cell.value != before_value or after_link != before_link:
                        image_cells_updated += 1

        except Exception as exc:
            errors += 1
            print(
                "  THUMBNAIL BACKFILL ERROR:",
                type(exc).__name__ + ":",
                exc
            )

    if image_cells_updated:
        print("THUMBNAIL BACKFILL: saving Master Database...")
        wb.save(MASTER_DATABASE)
        print("THUMBNAIL BACKFILL: Master Database saved.")
    else:
        print("THUMBNAIL BACKFILL: no Image cells needed changing; MDB not rewritten.")

    wb.close()

    elapsed = int(time.time() - started)
    minutes, seconds = divmod(elapsed, 60)

    print("")
    print("THUMBNAIL BACKFILL SUMMARY:")
    print(f"  Packages supplied: {len(package_paths):,}")
    print(f"  Packages matched to existing MDB SHA256: {matched_packages:,}")
    print(f"  New thumbnails extracted: {extracted:,}")
    print(f"  Existing thumbnails reused: {reused:,}")
    print(f"  No dedicated preview available: {no_preview:,}")
    if duplicate_hash_inputs:
        print(f"  Duplicate input hashes reused in this run: {duplicate_hash_inputs:,}")
    if not_in_mdb:
        print(f"  Not already in MDB (left untouched): {not_in_mdb:,}")
    if foreign:
        print(f"  Sims 2 / Sims 4 packages skipped: {foreign:,}")
    if empty_or_hash_error:
        print(f"  Empty/hash failures skipped: {empty_or_hash_error:,}")
    print(f"  Thumbnail errors: {errors:,}")
    print(f"  MDB Image cells updated: {image_cells_updated:,}")
    print(f"  Thumbnail folder: {THUMBNAILS_FOLDER}")
    print(f"  Backfill time: {minutes} minutes {seconds} seconds")
    print("THUMBNAIL BACKFILL: complete.")
    print("")


# -------------------------
# COMPACT MASTER DB REVIEW LAYOUT + COLOUR CODING TEST - 2026-09-05
# -------------------------
# Keep the everyday review fields on the left, move Normalized Package and
# other internal/reference fields away from the review area, and keep every
# verbose Broken ... Count/IDs diagnostic column as the final block.  Existing
# scanner-generated workbooks are migrated once by header name, so package
# records stay attached to their values even when the physical column order
# changes.
MASTER_REVIEW_FRONT_COLUMNS = [
    "Package",
    "Image",
    "Status",
    "Health Status",
    "Health Issues",
    "CC Type",
    "CASP Instance IDs",
    "Risk",
    "Risk Score",
    "Library Location",
    "CC Magic Profiles",
    "Notes",
    "Content Type",
    "Package Format",
    "HQ Warning",
    "Poly Warning",
    "Highest Polys",
    "Biggest Texture",
    "Skin Type",
    "Last Scanned",
    "Duplicate Files",
    "BONE Resource Count",
    "STBL IDs",
    "BGEO IDs",
    "FACE IDs",
    "FBLN IDs",
    "KEY IDs",
    "GEOM Instance IDs",
    "IMG Instance IDs",
    "VPXY Instance IDs",
]

MASTER_REVIEW_TAIL_COLUMNS = [
    "Subtype",
    "Original Source Path",
    "Package Hash",
    "Normalized Package",
]

PACKAGES_REVIEW_FRONT_COLUMNS = [
    "Package",
    "Health Status",
    "Health Issues",
    "CC Type",
    "Subtype",
    "CASP IDs",
    "Risk Level",
    "Risk Score",
    "Scan Date",
    "Package Format",
    "Highest Poly",
    "Highest Texture",
    "GEOM IDs",
    "IMG IDs",
    "VPXY IDs",
    "BONE Resource Count",
    "STBL IDs",
    "BGEO IDs",
    "FACE IDs",
    "FBLN IDs",
    "KEY IDs",
]

PACKAGES_REVIEW_TAIL_COLUMNS = [
    "Original Source Path",
    "Package Hash",
    "Normalized Package",
]


def _reorder_scanner_sheet_columns(ws, front_columns, tail_columns, label):
    """Reorder scanner-owned columns by header without copying row data.

    This intentionally operates only on sheets with a complete, unique header
    row and no formulas.  Those are invariants of the scanner-generated Master
    DB sheets.  If a user has introduced formulas or malformed headers, the
    migration is skipped rather than risking their workbook.
    """

    max_column = ws.max_column
    if max_column < 1:
        return False

    current_headers = [
        ws.cell(1, col_idx).value
        for col_idx in range(1, max_column + 1)
    ]

    if any(value in (None, "") for value in current_headers):
        print(f"MASTER DB: {label} layout migration skipped - blank header detected.")
        return False

    if len(set(current_headers)) != len(current_headers):
        print(f"MASTER DB: {label} layout migration skipped - duplicate header detected.")
        return False

    if any(cell.data_type == "f" for cell in ws._cells.values()):
        print(f"MASTER DB: {label} layout migration skipped - formula cells detected.")
        return False

    current_set = set(current_headers)
    broken_headers = [
        header for header in current_headers
        if isinstance(header, str) and header.startswith("Broken ")
    ]

    front_present = [
        header for header in front_columns
        if header in current_set
    ]
    tail_present = [
        header for header in tail_columns
        if header in current_set
    ]

    reserved = set(front_present) | set(tail_present) | set(broken_headers)
    other_nonbroken = [
        header for header in current_headers
        if header not in reserved
    ]

    target_headers = (
        front_present
        + other_nonbroken
        + tail_present
        + broken_headers
    )

    if target_headers == current_headers:
        return False

    new_position = {
        header: index
        for index, header in enumerate(target_headers, start=1)
    }
    old_to_new = {
        old_index: new_position[header]
        for old_index, header in enumerate(current_headers, start=1)
    }

    # Re-key the existing Cell objects instead of copying ~1.5 million cells.
    # Styles, values, hyperlinks and comments remain on the same Cell objects.
    new_cells = {}
    for (row_index, old_column), cell in list(ws._cells.items()):
        new_column = old_to_new.get(old_column, old_column)
        cell.column = new_column
        new_cells[(row_index, new_column)] = cell
    ws._cells = new_cells

    # Old width/hidden metadata belongs to old coordinates.  The layout pass
    # below reapplies intentional widths, outline levels and hidden columns by
    # header name, so clear stale coordinate-based column dimensions now.
    ws.column_dimensions.clear()

    print(f"MASTER DB: {label} columns migrated to review-friendly order.")
    return True


def _normalise_master_dropdown_ranges(ws):
    """Keep Status/Library dropdowns attached after a column migration."""

    column_map = get_column_map(ws)
    if "Status" not in column_map or "Library Location" not in column_map:
        return

    status_column = ws.cell(1, column_map["Status"]).column_letter
    location_column = ws.cell(1, column_map["Library Location"]).column_letter
    status_range = f"{status_column}2:{status_column}1048576"
    location_range = f"{location_column}2:{location_column}1048576"

    for validation in ws.data_validations.dataValidation:
        formula = str(validation.formula1 or "")
        if "Settings!$A$2:$A$6" in formula:
            validation.sqref = status_range
        elif "Settings!$B$2:$B$7" in formula:
            validation.sqref = location_range


# -------------------------
# COMPACT DISPLAY + HEADER DIVIDERS - TEST - 2026-09-05
# -------------------------
HEADER_BOTTOM_SIDE = Side(style="medium", color="FF8C8C8C")


def _apply_table_header_divider(ws):
    """Give scanner table headers a clear but restrained lower divider."""
    if ws.max_row < 1:
        return
    for cell in ws[1]:
        if cell.value not in (None, ""):
            cell.font = Font(bold=True)
            cell.alignment = Alignment(
                horizontal="left",
                vertical="top",
                wrap_text=True
            )
            cell.border = Border(bottom=HEADER_BOTTOM_SIDE)


def _display_cc_type_with_subtype(cc_type, subtype):
    """Compact main-view label while preserving raw Subtype separately."""
    base = str(cc_type or "").strip()
    sub = str(subtype or "").strip()
    if not sub:
        return base

    suffix = f" — {sub}"
    if base.endswith(suffix):
        return base

    # Prevent repeated suffixes if an already-combined Master DB is reopened.
    if " — " in base:
        base = base.split(" — ", 1)[0].strip()

    return f"{base}{suffix}" if base else sub


def _display_content_type(value):
    """Short spreadsheet display label without changing scanner logic."""
    text = str(value or "").strip()
    return "Custom" if text == "Custom Content" else text


def _display_health_status(value):
    """User-facing health label; internal scanner logic remains BROKEN."""
    text = str(value or "").strip()
    return "Malformed" if text.upper() == "BROKEN" else text


def _sync_health_status_display(ws, last_row):
    """Migrate/display BROKEN health cells as Malformed without changing logic."""
    column_map = get_column_map(ws)
    health_col = column_map.get("Health Status")
    if not health_col:
        return 0
    changed = 0
    for row_num in range(2, last_row + 1):
        cell = ws.cell(row_num, health_col)
        new_value = _display_health_status(cell.value)
        if new_value != str(cell.value or "").strip():
            cell.value = new_value
            changed += 1
    return changed


def _sync_master_content_type_display(ws, last_real_row):
    """Shorten the common Custom Content label in the visible Master DB."""
    column_map = get_column_map(ws)
    content_col = column_map.get("Content Type")
    if not content_col:
        return
    for row_num in range(2, last_real_row + 1):
        cell = ws.cell(row_num, content_col)
        cell.value = _display_content_type(cell.value)


def _sync_master_type_subtype_display(ws, last_real_row):
    """Combine Type/Subtype visually in CC Master without dropping raw subtype."""
    column_map = get_column_map(ws)
    type_col = column_map.get("CC Type")
    subtype_col = column_map.get("Subtype")
    if not type_col or not subtype_col:
        return
    for row_num in range(2, last_real_row + 1):
        cc_cell = ws.cell(row_num, type_col)
        cc_cell.value = _display_cc_type_with_subtype(
            cc_cell.value,
            ws.cell(row_num, subtype_col).value
        )


def _apply_casp_master_compact_layout(casp_ws):
    """Keep high-value CASP fields visible without oversized text columns."""
    column_map = get_column_map(casp_ws)
    widths = {
        "CASP Instance": 19,
        "Package": 34,
        "Normalized Package": 28,
        "CC Type": 16,
        "Data Type": 9,
        "Age Gender": 9,
        "Category": 16,
        "Body Type": 13,
        "Overlay Priority": 13,
    }
    for header, width in widths.items():
        if header in column_map:
            casp_ws.column_dimensions[
                get_column_letter(column_map[header])
            ].width = width
    _apply_table_header_divider(casp_ws)


# -------------------------
# MASTER DB ATTENTION COLOURS - DIRECT FILL TEST - 2026-09-05
# -------------------------
# LibreOffice did not reliably display the XLSX conditional-formatting rules
# used by the earlier colour tests.  Apply restrained pastel fills directly to
# the saved cells instead.  The scanner refreshes only its own known palette on
# every run, so unrelated user formatting is left alone.  Safe/OK remain plain.
MASTER_DB_FAINT_RED = "FFF4CCCC"
MASTER_DB_FAINT_ORANGE = "FFFCE5CD"
MASTER_DB_FAINT_YELLOW = "FFFFF2CC"
MASTER_DB_FAINT_GREY = "FFD9D9D9"
MASTER_DB_SCANNER_FILL_RGBS = {
    MASTER_DB_FAINT_RED,
    MASTER_DB_FAINT_ORANGE,
    MASTER_DB_FAINT_YELLOW,
    MASTER_DB_FAINT_GREY,
}

SIMS2_MAJOR1_ERROR_NEEDLE_A = "expected major version(s) '2, 3'"
SIMS2_MAJOR1_ERROR_NEEDLE_B = "found '1'"
FOREIGN_GAME_MARKER = "S3CC_FOREIGN_GAME="


def _read_dbpf_version(path):
    """Return (major, minor) for a DBPF file, or (None, None) if unavailable."""
    try:
        with open(path, "rb") as handle:
            header = handle.read(12)
        if len(header) < 12 or header[:4] != b"DBPF":
            return None, None
        major = int.from_bytes(header[4:8], "little", signed=False)
        minor = int.from_bytes(header[8:12], "little", signed=False)
        return major, minor
    except Exception:
        return None, None


def _known_foreign_sims_game(path):
    """Identify known non-Sims-3 DBPF package versions before s3pi parses them."""
    major, minor = _read_dbpf_version(path)
    if major == 1:
        return "SIMS_2", major, minor
    if major == 2 and minor == 1:
        return "SIMS_4", major, minor
    return "", major, minor


def _foreign_game_from_error_text(value):
    text = str(value or "")
    upper = text.upper()
    if f"{FOREIGN_GAME_MARKER}SIMS_2" in upper:
        return "SIMS_2"
    if f"{FOREIGN_GAME_MARKER}SIMS_4" in upper:
        return "SIMS_4"
    return ""


def _is_sims2_major1_error_text(value):
    """Recognise old s3pi DBPF-major-1 rejection text from earlier scans."""
    text = str(value or "").strip().lower()
    return (
        SIMS2_MAJOR1_ERROR_NEEDLE_A in text
        and SIMS2_MAJOR1_ERROR_NEEDLE_B in text
    )


def _row_is_foreign_game_unsupported(row):
    """Keep known Sims 2/Sims 4 packages out of the Sims 3 Master DB."""
    if len(row) > 22 and str(row[22] or "").strip().upper() == "UNSUPPORTED GAME":
        return True
    return len(row) > 25 and _is_sims2_major1_error_text(row[25])


def _clear_scanner_attention_conditional_formatting(ws):
    """Remove only obsolete scanner-owned conditional-formatting rules."""
    try:
        rules_by_range = ws.conditional_formatting._cf_rules
        for cf_key in list(rules_by_range.keys()):
            kept_rules = []
            for rule in rules_by_range.get(cf_key, []):
                formulas = [str(item) for item in (getattr(rule, "formula", None) or [])]
                if any("S3CCATTN" in formula for formula in formulas):
                    continue
                kept_rules.append(rule)
            if kept_rules:
                rules_by_range[cf_key] = kept_rules
            else:
                del rules_by_range[cf_key]
    except Exception:
        pass


def _scanner_fill_rgb(cell):
    """Return ARGB for one solid scanner-style fill, else an empty string."""
    try:
        fill = cell.fill
        if getattr(fill, "fill_type", None) != "solid":
            return ""
        rgb = getattr(getattr(fill, "fgColor", None), "rgb", None)
        return str(rgb or "").upper()
    except Exception:
        return ""


def _clear_scanner_fill(cell):
    """Clear only a fill that matches this scanner's own attention palette."""
    if _scanner_fill_rgb(cell) in MASTER_DB_SCANNER_FILL_RGBS:
        cell.fill = PatternFill(fill_type=None)


def _remove_existing_sims2_major1_rows(ws, label):
    """Remove old accidental DBPF-major-1 rows from a Sims 3 database sheet."""
    column_map = get_column_map(ws)
    if "Health Issues" not in column_map:
        return 0

    issues_col = column_map["Health Issues"]
    removed = 0
    for row_num in range(ws.max_row, 1, -1):
        if _is_sims2_major1_error_text(ws.cell(row_num, issues_col).value):
            ws.delete_rows(row_num, 1)
            removed += 1

    if removed:
        print(f"MASTER DB: removed {removed:,} accidental Sims 2/DBPF-v1 row(s) from {label}.")
    return removed



def _split_verbose_health_issue_entries(value):
    """Split an old multiline Health Issues stack-trace cell into resource entries."""
    text = str(value or "").strip()
    if not text:
        return []

    entries = []
    current = []
    for line in text.splitlines():
        if re.match(r"^\s*0x[0-9A-Fa-f]{8,16}\s*:\s*", line):
            if current:
                entries.append("\n".join(current).strip())
            current = [line.strip()]
        elif current:
            current.append(line)

    if current:
        entries.append("\n".join(current).strip())

    return [entry for entry in entries if entry]


def _health_issue_cell_needs_condensing(value):
    text = str(value or "")
    if not text.strip():
        return False
    if " — " in text and not any(
        marker in text
        for marker in (
            "TargetInvocationException",
            "\n   at ",
            "\nat ",
            "--- End of inner exception stack trace ---",
        )
    ):
        return False
    return (
        len(text) > 240
        or "TargetInvocationException" in text
        or "\n   at " in text
        or "\nat " in text
        or "--- End of inner exception stack trace ---" in text
    )


def _condense_existing_health_issue_cells(ws, last_row=None):
    """Migrate old scanner stack traces in Health Issues to the compact format."""
    column_map = get_column_map(ws)
    if "Health Issues" not in column_map:
        return 0

    issues_col = column_map["Health Issues"]
    if last_row is None:
        last_row = ws.max_row
    last_row = min(int(last_row or 0), ws.max_row)

    broken_id_columns = []
    for header, col_idx in column_map.items():
        if header.startswith("Broken ") and header.endswith(" IDs"):
            label = header[len("Broken "):-len(" IDs")].strip().upper()
            broken_id_columns.append((label, col_idx))

    changed = 0
    for row_num in range(2, last_row + 1):
        cell = ws.cell(row_num, issues_col)
        raw_value = cell.value
        if not _health_issue_cell_needs_condensing(raw_value):
            continue

        raw_text = str(raw_value or "").strip()
        entries = _split_verbose_health_issue_entries(raw_text)

        # Package-level failures do not carry resource instance IDs.
        if not entries:
            condensed = _condense_package_health_issue(raw_text)
            if condensed and condensed != raw_text:
                cell.value = condensed
                changed += 1
            continue

        id_to_label = {}
        for label, col_idx in broken_id_columns:
            ids_text = str(ws.cell(row_num, col_idx).value or "")
            for match in re.findall(r"0x[0-9A-Fa-f]{8,16}", ids_text):
                id_to_label["0x" + match[2:].upper()] = label

        grouped = {}
        order = []
        for entry in entries:
            instance_id, _error_text = _health_issue_instance_and_text(entry)
            label = id_to_label.get(instance_id, "RESOURCE")
            if label not in grouped:
                grouped[label] = []
                order.append(label)
            grouped[label].append(entry)

        condensed_lines = _build_condensed_health_issues(
            [(label, grouped[label]) for label in order]
        )
        condensed = "\n".join(condensed_lines).strip()
        if condensed and condensed != raw_text:
            cell.value = condensed
            changed += 1

    return changed


def _apply_direct_attention_colours(master_ws, packages_ws, master_last_row):
    """Apply visible direct fills for attention states and exact duplicates."""

    _clear_scanner_attention_conditional_formatting(master_ws)
    _clear_scanner_attention_conditional_formatting(packages_ws)

    red_fill = PatternFill(fill_type="solid", fgColor=MASTER_DB_FAINT_RED)
    orange_fill = PatternFill(fill_type="solid", fgColor=MASTER_DB_FAINT_ORANGE)
    yellow_fill = PatternFill(fill_type="solid", fgColor=MASTER_DB_FAINT_YELLOW)
    grey_fill = PatternFill(fill_type="solid", fgColor=MASTER_DB_FAINT_GREY)

    packages_map = get_column_map(packages_ws)
    master_map = get_column_map(master_ws)

    # Exact duplicate SHA-256 groups come only from the physical Packages ledger.
    duplicate_hashes = set()
    if "Package Hash" in packages_map:
        hash_col = packages_map["Package Hash"]
        counts = {}
        for row_num in range(2, packages_ws.max_row + 1):
            value = str(packages_ws.cell(row_num, hash_col).value or "").strip().lower()
            if value:
                counts[value] = counts.get(value, 0) + 1
        duplicate_hashes = {value for value, count in counts.items() if count > 1}

    def apply_sheet(ws, column_map, last_row, has_manual_status=False):
        package_col = column_map.get("Package")
        health_col = column_map.get("Health Status")
        risk_col = column_map.get("Risk") or column_map.get("Risk Level")
        poly_col = column_map.get("Poly Warning")
        hash_col = column_map.get("Package Hash")
        status_col = column_map.get("Status") if has_manual_status else None

        # Remove the scanner's previous colour from controlled cells first.
        controlled_cols = [c for c in (health_col, risk_col, poly_col, hash_col) if c]
        for row_num in range(2, last_row + 1):
            for col_idx in controlled_cols:
                _clear_scanner_fill(ws.cell(row_num, col_idx))

            # If this row was previously greyed as Deleted but is no longer
            # Deleted, restore only scanner-grey cells and leave user fills alone.
            is_deleted = (
                status_col is not None
                and str(ws.cell(row_num, status_col).value or "").strip().lower() in ("delete", "deleted")
            )
            row_was_scanner_grey = (
                package_col is not None
                and _scanner_fill_rgb(ws.cell(row_num, package_col)) == MASTER_DB_FAINT_GREY
            )
            if row_was_scanner_grey and not is_deleted:
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row_num, col_idx)
                    if _scanner_fill_rgb(cell) == MASTER_DB_FAINT_GREY:
                        cell.fill = PatternFill(fill_type=None)

            if health_col is not None:
                health = str(ws.cell(row_num, health_col).value or "").strip().upper()
                if health in ("BROKEN", "MALFORMED"):
                    ws.cell(row_num, health_col).fill = red_fill

            if risk_col is not None:
                risk = str(ws.cell(row_num, risk_col).value or "").strip().upper()
                if risk == "EXTREME":
                    ws.cell(row_num, risk_col).fill = red_fill
                elif risk == "HIGH":
                    ws.cell(row_num, risk_col).fill = orange_fill
                elif risk == "MODERATE":
                    ws.cell(row_num, risk_col).fill = yellow_fill

            if poly_col is not None:
                poly = str(ws.cell(row_num, poly_col).value or "").strip().upper()
                if poly == "VERY HEAVY":
                    ws.cell(row_num, poly_col).fill = red_fill
                elif poly == "HEAVY":
                    ws.cell(row_num, poly_col).fill = orange_fill

            if hash_col is not None:
                hash_value = str(ws.cell(row_num, hash_col).value or "").strip().lower()
                if hash_value and hash_value in duplicate_hashes:
                    ws.cell(row_num, hash_col).fill = orange_fill

            # Deleted deliberately overrides all warning colours across the row.
            if is_deleted:
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row_num, col_idx).fill = grey_fill

    apply_sheet(master_ws, master_map, master_last_row, has_manual_status=True)
    apply_sheet(packages_ws, packages_map, packages_ws.max_row, has_manual_status=False)
    return len(duplicate_hashes)


def _apply_review_sheet_layout(ws, layout_kind):
    """Apply compact widths and collapsible diagnostic sections."""

    column_map = get_column_map(ws)

    # Clear stale outline/hidden state before applying the current layout.
    for col_idx in range(1, ws.max_column + 1):
        dim = ws.column_dimensions[get_column_letter(col_idx)]
        dim.outlineLevel = 0
        dim.hidden = False
        dim.collapsed = False

    if layout_kind == "master":
        widths = {
            "Package": 36,
            "Image": 15,
            "Status": 12,
            "Health Status": 12,
            "Health Issues": 15,
            "CC Type": 18,
            "Subtype": 11,
            "CASP Instance IDs": 20,
            "Risk": 10,
            "Risk Score": 10,
            "Library Location": 16,
            "Notes": 12,
            "Content Type": 12,
            "Package Format": 16,
            "HQ Warning": 14,
            "Poly Warning": 14,
            "Highest Polys": 13,
            "Biggest Texture": 16,
            "Skin Type": 11,
            "Last Scanned": 14,
            "Duplicate Files": 30,
            "Original Source Path": 24,
            "BONE Resource Count": 18,
            "STBL IDs": 26,
            "BGEO IDs": 26,
            "FACE IDs": 26,
            "FBLN IDs": 26,
            "KEY IDs": 26,
            "GEOM Instance IDs": 28,
            "IMG Instance IDs": 28,
            "VPXY Instance IDs": 28,
            "Package Hash": 20,
            "Normalized Package": 32,
        }
        performance_headers = {
            "Risk", "Risk Score", "Content Type", "Package Format",
            "HQ Warning", "Poly Warning", "Highest Polys",
            "Biggest Texture", "Skin Type", "Last Scanned",
        }
        technical_headers = {
            "Duplicate Files", "Original Source Path", "BONE Resource Count",
            "STBL IDs", "BGEO IDs", "FACE IDs", "FBLN IDs", "KEY IDs",
            "GEOM Instance IDs", "IMG Instance IDs", "VPXY Instance IDs",
            "Package Hash", "Normalized Package",
        }
    else:
        widths = {
            "Package": 36,
            "Health Status": 12,
            "Health Issues": 15,
            "CC Type": 14,
            "Subtype": 11,
            "CASP IDs": 18,
            "Risk Level": 10,
            "Risk Score": 10,
            "Scan Date": 14,
            "Package Format": 16,
            "Highest Poly": 13,
            "Highest Texture": 16,
            "Original Source Path": 24,
            "GEOM IDs": 28,
            "IMG IDs": 28,
            "VPXY IDs": 28,
            "BONE Resource Count": 18,
            "STBL IDs": 26,
            "BGEO IDs": 26,
            "FACE IDs": 26,
            "FBLN IDs": 26,
            "KEY IDs": 26,
            "Package Hash": 20,
            "Normalized Package": 32,
        }
        performance_headers = {
            "Risk Level", "Risk Score", "Scan Date", "Package Format",
            "Highest Poly", "Highest Texture",
        }
        technical_headers = {
            "Original Source Path", "GEOM IDs", "IMG IDs", "VPXY IDs",
            "BONE Resource Count", "STBL IDs", "BGEO IDs", "FACE IDs",
            "FBLN IDs", "KEY IDs", "Package Hash", "Normalized Package",
        }

    for header, col_idx in column_map.items():
        letter = get_column_letter(col_idx)
        dim = ws.column_dimensions[letter]

        if isinstance(header, str) and header.startswith("Broken "):
            # Count fields are tiny; ID diagnostics are kept readable but no
            # longer allowed to make the workbook dozens of screens wide.
            dim.width = 11 if header.endswith(" Count") else 24
            dim.outlineLevel = 3 if layout_kind == "master" else 2
            dim.hidden = True
            ws.cell(1, col_idx).alignment = Alignment(
                wrap_text=True,
                vertical="top"
            )
            continue

        dim.width = widths.get(header, 16)

        # CC Master displays Subtype inside CC Type; retain the raw Subtype
        # field for compatibility but keep it out of the everyday review view.
        if layout_kind == "master" and header == "Subtype":
            dim.outlineLevel = 2
            dim.hidden = True
            continue

        if header in performance_headers:
            dim.outlineLevel = 1
        elif header in technical_headers:
            dim.outlineLevel = 2 if layout_kind == "master" else 1

    # Package Hash stays visible so exact duplicates can be checked manually.
    # It remains in the far-right technical section.

    # LibreOffice/Excel use outline metadata to show +/- collapse controls.
    ws.sheet_properties.outlinePr.summaryRight = True

    for header in (
        "Package", "Health Status", "Health Issues", "CC Type", "Subtype",
        "CASP Instance IDs", "CASP IDs", "Library Location", "CC Magic Profiles", "Risk", "Risk Level",
        "Risk Score", "Notes", "Skin Type", "Duplicate Files", "Original Source Path",
        "Normalized Package", "GEOM IDs", "IMG IDs", "VPXY IDs",
        "STBL IDs", "BGEO IDs", "FACE IDs", "FBLN IDs", "KEY IDs",
        "GEOM Instance IDs", "IMG Instance IDs", "VPXY Instance IDs",
    ):
        if header in column_map:
            ws.cell(1, column_map[header]).alignment = Alignment(
                wrap_text=True,
                vertical="top"
            )
    
def update_master_database(rows, casp_records):

    from datetime import datetime
    from bisect import insort
    import time

    master_db_started = time.perf_counter()
    print(f"MASTER DB: starting update for {len(rows):,} packages / {len(casp_records):,} CASP records...")

    if DEBUG:
        print("ROWS DEBUG:")
        print(rows[0] if rows else "NO ROWS")

    if os.path.exists(MASTER_DATABASE):

        print("MASTER DB: loading existing workbook...")
        load_started = time.perf_counter()
        wb = openpyxl.load_workbook(MASTER_DATABASE)
        print(f"MASTER DB: workbook loaded in {time.perf_counter() - load_started:.1f}s")

        if "CC Master Database" in wb.sheetnames:

            ws = wb["CC Master Database"]

        else:

            ws = wb.active
            ws.title = "CC Master Database"

        column_map = get_column_map(ws)

    else:

        print("MASTER DB: creating new workbook...")
        wb = Workbook()
        ws = wb.active
        ws.title = "CC Master Database"
        
    # -------------------------
    # CREATE SETTINGS SHEET
    # -------------------------

        if "Settings" not in wb.sheetnames:

            settings_ws = wb.create_sheet("Settings")

            settings_ws["A1"] = "Status Options"
            settings_ws["A2"] = "Unreviewed"
            settings_ws["A3"] = "Keep"
            settings_ws["A4"] = "Review"
            settings_ws["A5"] = "Favourite"
            settings_ws["A6"] = "Delete"


            settings_ws["B1"] = "Library Options"
            settings_ws["B2"] = "Unsorted"
            settings_ws["B3"] = "In Game Use"
            settings_ws["B4"] = "In Library Collection"
            settings_ws["B5"] = "CC Magic Merge"
            settings_ws["B6"] = "Archive"
            settings_ws["B7"] = "Deleted"     
        
        ws.append([
            "Package",
            "Image",
            "Status",
            "Library Location",
            "CC Magic Profiles",
            "Notes",
            "Normalized Package",
            "CC Type",
            "Subtype",
            "Package Format",
            "Original Source Path",
            "Content Type",
            "Risk",
            "Risk Score",
            "Health Status",
            "Broken GEOM Count",
            "Broken GEOM IDs",
            "Broken CASP Count",
            "Broken CASP IDs",
            "Broken IMG Count",
            "Broken IMG IDs",
            "Broken VPXY Count",
            "Broken VPXY IDs",
            "Broken BONE Count",
            "Broken BONE IDs",
            "Broken BOND Count",
            "Broken BOND IDs",
            "Broken Morph Count",
            "Broken Morph IDs",
            "Broken Build/Buy Count",
            "Broken Build/Buy IDs",
            "Broken PRESET Count",
            "Broken PRESET IDs",
            "Broken TONE Count",
            "Broken TONE IDs",
            "Broken IMAG Count",
            "Broken IMAG IDs",
            "Broken TXTC Count",
            "Broken TXTC IDs",
            "Broken PTRN Count",
            "Broken PTRN IDs",
            "Broken SIMO Count",
            "Broken SIMO IDs",
            "Broken TXTF Count",
            "Broken TXTF IDs",
            "Broken FAMD Count",
            "Broken FAMD IDs",
            "Broken OBJD Count",
            "Broken OBJD IDs",
            "Broken OBJK Count",
            "Broken OBJK IDs",
            "Broken XML Count",
            "Broken XML IDs",
            "Broken STBL Count",
            "Broken STBL IDs",
            "Broken NMAP Count",
            "Broken NMAP IDs",
            "Broken RSLT Count",
            "Broken RSLT IDs",
            "Broken FTPT Count",
            "Broken FTPT IDs",
            "Broken RIG Count",
            "Broken RIG IDs",
            "Broken JAZZ Count",
            "Broken JAZZ IDs",
            "Broken S3SA Count",
            "Broken S3SA IDs",
            "Broken LITE Count",
            "Broken LITE IDs",
            "Broken CLIP Count",
            "Broken CLIP IDs",
            "Broken MATD Count",
            "Broken MATD IDs",
            "Broken MTST Count",
            "Broken MTST IDs",
            "Broken MDLR Count",
            "Broken MDLR IDs",
            "Broken CWAL Count",
            "Broken CWAL IDs",
            "Broken CFEN Count",
            "Broken CFEN IDs",
            "Broken CTPT Count",
            "Broken CTPT IDs",
            "Broken NGMP Count",
            "Broken NGMP IDs",
            "Broken CRMT Count",
            "Broken CRMT IDs",
            "Broken SPT2 Count",
            "Broken SPT2 IDs",
            "Broken _SPT Count",
            "Broken _SPT IDs",
            "Broken CRAL Count",
            "Broken CRAL IDs",
            "Broken LAYO Count",
            "Broken LAYO IDs",
            "Broken HINF Count",
            "Broken HINF IDs",
            "Broken 2ARY Count",
            "Broken 2ARY IDs",
            "Broken Complate Count",
            "Broken Complate IDs",
            "Broken 63A33EA7 Count",
            "Broken 63A33EA7 IDs",
            "Broken Package Storage Count",
            "Broken Package Storage IDs",
            "Broken Fireplace Count",
            "Broken Fireplace IDs",
            "Broken THUM Count",
            "Broken THUM IDs",
            "Broken 626F60CE Count",
            "Broken 626F60CE IDs",
            "Broken ICON Count",
            "Broken ICON IDs",
            "Broken THUM 0580A2B5 Count",
            "Broken THUM 0580A2B5 IDs",
            "Broken THUM 0580A2B4 Count",
            "Broken THUM 0580A2B4 IDs",
            "Broken THUM 0580A2B6 Count",
            "Broken THUM 0580A2B6 IDs",
            "Broken SNAP 0580A2CD Count",
            "Broken SNAP 0580A2CD IDs",
            "Broken SNAP 0580A2CF Count",
            "Broken SNAP 0580A2CF IDs",
            "Broken TWNI Count",
            "Broken TWNI IDs",
            "Broken MANIFEST Count",
            "Broken MANIFEST IDs",
            "Broken 8F78F74B Count",
            "Broken 8F78F74B IDs",
            "Health Issues",
            "HQ Warning",
            "Poly Warning",
            "Highest Polys",
            "Biggest Texture",
            "Skin Type",
            "Last Scanned",
            "CASP Instance IDs",
            "Duplicate Files",
            "BONE Resource Count",
            "STBL IDs",
            "BGEO IDs",
            "FACE IDs",
            "FBLN IDs",
            "KEY IDs",
            "GEOM Instance IDs",
            "IMG Instance IDs",
            "VPXY Instance IDs",
            "Package Hash"
        ])
        # -------------------------
        # ADD DROPDOWN MENUS
        # -------------------------

        status_validation = DataValidation(
            type="list",
            formula1="=Settings!$A$2:$A$6",
            allow_blank=True
        )

        location_validation = DataValidation(
            type="list",
            formula1="=Settings!$B$2:$B$7",
            allow_blank=True
        )


        ws.add_data_validation(status_validation)
        ws.add_data_validation(location_validation)


        status_validation.add("C2:C1048576")
        location_validation.add("D2:D1048576")

    # -------------------------
    # ENSURE SETTINGS SHEET EXISTS
    # -------------------------

    if "Settings" not in wb.sheetnames:

        settings_ws = wb.create_sheet("Settings")

        settings_ws["A1"] = "Status Options"
        settings_ws["A2"] = "Unreviewed"
        settings_ws["A3"] = "Keep"
        settings_ws["A4"] = "Review"
        settings_ws["A5"] = "Favourite"
        settings_ws["A6"] = "Delete"

        settings_ws["B1"] = "Library Options"
        settings_ws["B2"] = "Unsorted"
        settings_ws["B3"] = "In Game Use"
        settings_ws["B4"] = "In Library Collection"
        settings_ws["B5"] = "CC Magic Merge"
        settings_ws["B6"] = "Archive"
        settings_ws["B7"] = "Deleted"

    # -------------------------
    # CREATE PACKAGES SHEET
    # -------------------------
    if "Packages" not in wb.sheetnames:

        packages_ws = wb.create_sheet("Packages")

        packages_ws.append([
            "Package",
            "Normalized Package",
            "Scan Date",
            "Risk Level",
            "Risk Score",
            "Health Status",
            "Broken GEOM Count",
            "Broken GEOM IDs",
            "Broken CASP Count",
            "Broken CASP IDs",
            "Broken IMG Count",
            "Broken IMG IDs",
            "Broken VPXY Count",
            "Broken VPXY IDs",
            "Broken BONE Count",
            "Broken BONE IDs",
            "Broken BOND Count",
            "Broken BOND IDs",
            "Broken Morph Count",
            "Broken Morph IDs",
            "Broken Build/Buy Count",
            "Broken Build/Buy IDs",
            "Broken PRESET Count",
            "Broken PRESET IDs",
            "Broken TONE Count",
            "Broken TONE IDs",
            "Broken IMAG Count",
            "Broken IMAG IDs",
            "Broken TXTC Count",
            "Broken TXTC IDs",
            "Broken PTRN Count",
            "Broken PTRN IDs",
            "Broken SIMO Count",
            "Broken SIMO IDs",
            "Broken TXTF Count",
            "Broken TXTF IDs",
            "Broken FAMD Count",
            "Broken FAMD IDs",
            "Broken OBJD Count",
            "Broken OBJD IDs",
            "Broken OBJK Count",
            "Broken OBJK IDs",
            "Broken XML Count",
            "Broken XML IDs",
            "Broken STBL Count",
            "Broken STBL IDs",
            "Broken NMAP Count",
            "Broken NMAP IDs",
            "Broken RSLT Count",
            "Broken RSLT IDs",
            "Broken FTPT Count",
            "Broken FTPT IDs",
            "Broken RIG Count",
            "Broken RIG IDs",
            "Broken JAZZ Count",
            "Broken JAZZ IDs",
            "Broken S3SA Count",
            "Broken S3SA IDs",
            "Broken LITE Count",
            "Broken LITE IDs",
            "Broken CLIP Count",
            "Broken CLIP IDs",
            "Broken MATD Count",
            "Broken MATD IDs",
            "Broken MTST Count",
            "Broken MTST IDs",
            "Broken MDLR Count",
            "Broken MDLR IDs",
            "Broken CWAL Count",
            "Broken CWAL IDs",
            "Broken CFEN Count",
            "Broken CFEN IDs",
            "Broken CTPT Count",
            "Broken CTPT IDs",
            "Broken NGMP Count",
            "Broken NGMP IDs",
            "Broken CRMT Count",
            "Broken CRMT IDs",
            "Broken SPT2 Count",
            "Broken SPT2 IDs",
            "Broken _SPT Count",
            "Broken _SPT IDs",
            "Broken CRAL Count",
            "Broken CRAL IDs",
            "Broken LAYO Count",
            "Broken LAYO IDs",
            "Broken HINF Count",
            "Broken HINF IDs",
            "Broken 2ARY Count",
            "Broken 2ARY IDs",
            "Broken Complate Count",
            "Broken Complate IDs",
            "Broken 63A33EA7 Count",
            "Broken 63A33EA7 IDs",
            "Broken Package Storage Count",
            "Broken Package Storage IDs",
            "Broken Fireplace Count",
            "Broken Fireplace IDs",
            "Broken THUM Count",
            "Broken THUM IDs",
            "Broken 626F60CE Count",
            "Broken 626F60CE IDs",
            "Broken ICON Count",
            "Broken ICON IDs",
            "Broken THUM 0580A2B5 Count",
            "Broken THUM 0580A2B5 IDs",
            "Broken THUM 0580A2B4 Count",
            "Broken THUM 0580A2B4 IDs",
            "Broken THUM 0580A2B6 Count",
            "Broken THUM 0580A2B6 IDs",
            "Broken SNAP 0580A2CD Count",
            "Broken SNAP 0580A2CD IDs",
            "Broken SNAP 0580A2CF Count",
            "Broken SNAP 0580A2CF IDs",
            "Broken TWNI Count",
            "Broken TWNI IDs",
            "Broken MANIFEST Count",
            "Broken MANIFEST IDs",
            "Broken 8F78F74B Count",
            "Broken 8F78F74B IDs",
            "Health Issues",
            "CC Type",
            "Subtype",
            "Package Format",
            "Original Source Path",
            "CASP IDs",
            "GEOM IDs",
            "IMG IDs",
            "VPXY IDs",
            "BONE Resource Count",
            "STBL IDs",
            "BGEO IDs",
            "FACE IDs",
            "FBLN IDs",
            "KEY IDs",
            "Highest Poly",
            "Highest Texture",
            "Package Hash"
        ])

    else:

        packages_ws = wb["Packages"]

    packages_column_map = get_column_map(packages_ws)

    if "Bone IDs" in packages_column_map and "BONE Resource Count" not in packages_column_map:

        packages_ws.cell(
            1,
            packages_column_map["Bone IDs"]
        ).value = "BONE Resource Count"

        packages_column_map = get_column_map(packages_ws)

    # -------------------------
    # ADD MISSING PACKAGES COLUMNS
    # -------------------------

    required_packages_columns = [
        "Package",
        "Normalized Package",
        "Scan Date",
        "Risk Level",
        "Risk Score",
        "Health Status",
        "Broken GEOM Count",
        "Broken GEOM IDs",
        "Broken CASP Count",
        "Broken CASP IDs",
        "Broken IMG Count",
        "Broken IMG IDs",
        "Broken VPXY Count",
        "Broken VPXY IDs",
        "Broken BONE Count",
        "Broken BONE IDs",
        "Broken BOND Count",
        "Broken BOND IDs",
        "Broken Morph Count",
        "Broken Morph IDs",
        "Broken Build/Buy Count",
        "Broken Build/Buy IDs",
        "Broken PRESET Count",
        "Broken PRESET IDs",
        "Broken TONE Count",
        "Broken TONE IDs",
        "Broken IMAG Count",
        "Broken IMAG IDs",
        "Broken TXTC Count",
        "Broken TXTC IDs",
        "Broken PTRN Count",
        "Broken PTRN IDs",
        "Broken SIMO Count",
        "Broken SIMO IDs",
        "Broken TXTF Count",
        "Broken TXTF IDs",
        "Broken FAMD Count",
        "Broken FAMD IDs",
        "Broken OBJD Count",
        "Broken OBJD IDs",
        "Broken OBJK Count",
        "Broken OBJK IDs",
        "Broken XML Count",
        "Broken XML IDs",
        "Broken STBL Count",
        "Broken STBL IDs",
        "Broken NMAP Count",
        "Broken NMAP IDs",
        "Broken RSLT Count",
        "Broken RSLT IDs",
        "Broken FTPT Count",
        "Broken FTPT IDs",
        "Broken RIG Count",
        "Broken RIG IDs",
        "Broken JAZZ Count",
        "Broken JAZZ IDs",
        "Broken S3SA Count",
        "Broken S3SA IDs",
        "Broken LITE Count",
        "Broken LITE IDs",
        "Broken CLIP Count",
        "Broken CLIP IDs",
        "Broken MATD Count",
        "Broken MATD IDs",
        "Broken MTST Count",
        "Broken MTST IDs",
        "Broken MDLR Count",
        "Broken MDLR IDs",
        "Broken CWAL Count",
        "Broken CWAL IDs",
        "Broken CFEN Count",
        "Broken CFEN IDs",
        "Broken CTPT Count",
        "Broken CTPT IDs",
        "Broken NGMP Count",
        "Broken NGMP IDs",
        "Broken CRMT Count",
        "Broken CRMT IDs",
        "Broken SPT2 Count",
        "Broken SPT2 IDs",
        "Broken _SPT Count",
        "Broken _SPT IDs",
        "Broken CRAL Count",
        "Broken CRAL IDs",
        "Broken LAYO Count",
        "Broken LAYO IDs",
        "Broken HINF Count",
        "Broken HINF IDs",
        "Broken 2ARY Count",
        "Broken 2ARY IDs",
        "Broken Complate Count",
        "Broken Complate IDs",
        "Broken 63A33EA7 Count",
        "Broken 63A33EA7 IDs",
        "Broken Package Storage Count",
        "Broken Package Storage IDs",
        "Broken Fireplace Count",
        "Broken Fireplace IDs",
        "Broken THUM Count",
        "Broken THUM IDs",
        "Broken 626F60CE Count",
        "Broken 626F60CE IDs",
        "Broken ICON Count",
        "Broken ICON IDs",
        "Broken THUM 0580A2B5 Count",
        "Broken THUM 0580A2B5 IDs",
        "Broken THUM 0580A2B4 Count",
        "Broken THUM 0580A2B4 IDs",
        "Broken THUM 0580A2B6 Count",
        "Broken THUM 0580A2B6 IDs",
        "Broken SNAP 0580A2CD Count",
        "Broken SNAP 0580A2CD IDs",
        "Broken SNAP 0580A2CF Count",
        "Broken SNAP 0580A2CF IDs",
        "Broken TWNI Count",
        "Broken TWNI IDs",
        "Broken MANIFEST Count",
        "Broken MANIFEST IDs",
        "Broken 8F78F74B Count",
        "Broken 8F78F74B IDs",
        "Health Issues",
        "CC Type",
        "Subtype",
        "Package Format",
        "Original Source Path",
        "CASP IDs",
        "GEOM IDs",
        "IMG IDs",
        "VPXY IDs",
        "BONE Resource Count",
        "STBL IDs",
        "BGEO IDs",
        "FACE IDs",
        "FBLN IDs",
        "KEY IDs",
        "Highest Poly",
        "Highest Texture",
        "Package Hash"
    ]

    for header in required_packages_columns:

        if header not in packages_column_map:

            packages_ws.cell(
                1,
                packages_ws.max_column + 1
            ).value = header

    packages_column_map = get_column_map(packages_ws)

    _reorder_scanner_sheet_columns(
        packages_ws,
        PACKAGES_REVIEW_FRONT_COLUMNS,
        PACKAGES_REVIEW_TAIL_COLUMNS,
        "Packages"
    )
    packages_column_map = get_column_map(packages_ws)

    column_map = get_column_map(ws)

    # -------------------------
    # MIGRATE OLD BONE HEADER
    # -------------------------

    if "Bone IDs" in column_map and "BONE Resource Count" not in column_map:

        ws.cell(
            1,
            column_map["Bone IDs"]
        ).value = "BONE Resource Count"

        column_map = get_column_map(ws)

    # -------------------------
    # ADD MISSING SCANNER COLUMNS
    # -------------------------

    required_master_columns = [
        "Package",
        "Image",
        "Status",
        "Library Location",
        "CC Magic Profiles",
        "Notes",
        "Normalized Package",
        "CC Type",
        "Subtype",
        "Package Format",
        "Original Source Path",
        "Content Type",
        "Risk",
        "Risk Score",
        "Health Status",
        "Broken GEOM Count",
        "Broken GEOM IDs",
        "Broken CASP Count",
        "Broken CASP IDs",
        "Broken IMG Count",
        "Broken IMG IDs",
        "Broken VPXY Count",
        "Broken VPXY IDs",
        "Broken BONE Count",
        "Broken BONE IDs",
        "Broken BOND Count",
        "Broken BOND IDs",
        "Broken Morph Count",
        "Broken Morph IDs",
        "Broken Build/Buy Count",
        "Broken Build/Buy IDs",
        "Broken PRESET Count",
        "Broken PRESET IDs",
        "Broken TONE Count",
        "Broken TONE IDs",
        "Broken IMAG Count",
        "Broken IMAG IDs",
        "Broken TXTC Count",
        "Broken TXTC IDs",
        "Broken PTRN Count",
        "Broken PTRN IDs",
        "Broken SIMO Count",
        "Broken SIMO IDs",
        "Broken TXTF Count",
        "Broken TXTF IDs",
        "Broken FAMD Count",
        "Broken FAMD IDs",
        "Broken OBJD Count",
        "Broken OBJD IDs",
        "Broken OBJK Count",
        "Broken OBJK IDs",
        "Broken XML Count",
        "Broken XML IDs",
        "Broken STBL Count",
        "Broken STBL IDs",
        "Broken NMAP Count",
        "Broken NMAP IDs",
        "Broken RSLT Count",
        "Broken RSLT IDs",
        "Broken FTPT Count",
        "Broken FTPT IDs",
        "Broken RIG Count",
        "Broken RIG IDs",
        "Broken JAZZ Count",
        "Broken JAZZ IDs",
        "Broken S3SA Count",
        "Broken S3SA IDs",
        "Broken LITE Count",
        "Broken LITE IDs",
        "Broken CLIP Count",
        "Broken CLIP IDs",
        "Broken MATD Count",
        "Broken MATD IDs",
        "Broken MTST Count",
        "Broken MTST IDs",
        "Broken MDLR Count",
        "Broken MDLR IDs",
        "Broken CWAL Count",
        "Broken CWAL IDs",
        "Broken CFEN Count",
        "Broken CFEN IDs",
        "Broken CTPT Count",
        "Broken CTPT IDs",
        "Broken NGMP Count",
        "Broken NGMP IDs",
        "Broken CRMT Count",
        "Broken CRMT IDs",
        "Broken SPT2 Count",
        "Broken SPT2 IDs",
        "Broken _SPT Count",
        "Broken _SPT IDs",
        "Broken CRAL Count",
        "Broken CRAL IDs",
        "Broken LAYO Count",
        "Broken LAYO IDs",
        "Broken HINF Count",
        "Broken HINF IDs",
        "Broken 2ARY Count",
        "Broken 2ARY IDs",
        "Broken Complate Count",
        "Broken Complate IDs",
        "Broken 63A33EA7 Count",
        "Broken 63A33EA7 IDs",
        "Broken Package Storage Count",
        "Broken Package Storage IDs",
        "Broken Fireplace Count",
        "Broken Fireplace IDs",
        "Broken THUM Count",
        "Broken THUM IDs",
        "Broken 626F60CE Count",
        "Broken 626F60CE IDs",
        "Broken ICON Count",
        "Broken ICON IDs",
        "Broken THUM 0580A2B5 Count",
        "Broken THUM 0580A2B5 IDs",
        "Broken THUM 0580A2B4 Count",
        "Broken THUM 0580A2B4 IDs",
        "Broken THUM 0580A2B6 Count",
        "Broken THUM 0580A2B6 IDs",
        "Broken SNAP 0580A2CD Count",
        "Broken SNAP 0580A2CD IDs",
        "Broken SNAP 0580A2CF Count",
        "Broken SNAP 0580A2CF IDs",
        "Broken TWNI Count",
        "Broken TWNI IDs",
        "Broken MANIFEST Count",
        "Broken MANIFEST IDs",
        "Broken 8F78F74B Count",
        "Broken 8F78F74B IDs",
        "Health Issues",
        "HQ Warning",
        "Poly Warning",
        "Highest Polys",
        "Biggest Texture",
        "Skin Type",
        "Last Scanned",
        "CASP Instance IDs",
        "Duplicate Files",
        "BONE Resource Count",
        "STBL IDs",
        "BGEO IDs",
        "FACE IDs",
        "FBLN IDs",
        "KEY IDs",
        "GEOM Instance IDs",
        "IMG Instance IDs",
        "VPXY Instance IDs",
        "Package Hash"
    ]

    for header in required_master_columns:

        if header not in column_map:

            ws.cell(
                1,
                ws.max_column + 1
            ).value = header

    column_map = get_column_map(ws)

    _reorder_scanner_sheet_columns(
        ws,
        MASTER_REVIEW_FRONT_COLUMNS,
        MASTER_REVIEW_TAIL_COLUMNS,
        "CC Master Database"
    )
    _normalise_master_dropdown_ranges(ws)
    column_map = get_column_map(ws)

    # Remove accidental Sims 2 / DBPF-major-1 scan-error rows from earlier
    # versions.  They are unsupported input, not Sims 3 Master DB records.
    _remove_existing_sims2_major1_rows(ws, "CC Master Database")
    _remove_existing_sims2_major1_rows(packages_ws, "Packages")
    column_map = get_column_map(ws)
    packages_column_map = get_column_map(packages_ws)

    # -------------------------
    # ENSURE DROPDOWN MENUS EXIST
    # -------------------------

    status_column = ws.cell(
        1,
        column_map["Status"]
    ).column_letter

    location_column = ws.cell(
        1,
        column_map["Library Location"]
    ).column_letter

    status_range = (
        f"{status_column}2:"
        f"{status_column}1048576"
    )

    location_range = (
        f"{location_column}2:"
        f"{location_column}1048576"
    )

    existing_validations = list(
        ws.data_validations.dataValidation
    )

    status_validation_exists = False
    location_validation_exists = False

    for validation in existing_validations:
        formula = str(validation.formula1 or "")
        if "Settings!$A$2:$A$6" in formula:
            validation.sqref = status_range
            status_validation_exists = True
        elif "Settings!$B$2:$B$7" in formula:
            validation.sqref = location_range
            location_validation_exists = True

    if not status_validation_exists:

        status_validation = DataValidation(
            type="list",
            formula1="=Settings!$A$2:$A$6",
            allow_blank=True
        )

        ws.add_data_validation(status_validation)
        status_validation.add(status_range)

    if not location_validation_exists:

        location_validation = DataValidation(
            type="list",
            formula1="=Settings!$B$2:$B$7",
            allow_blank=True
        )

        ws.add_data_validation(location_validation)
        location_validation.add(location_range)

    # -------------------------
    # INDEX CC MASTER DATABASE
    # -------------------------
    # One worksheet pass builds all identity lookups used below.  The old
    # updater re-scanned the entire ~150-column Master sheet once or twice
    # for every incoming package, which becomes extremely slow at 9,000+.

    print("MASTER DB: indexing CC Master Database...")
    index_started = time.perf_counter()

    existing_casp = {}
    master_package_rows = {}
    master_row_package = {}
    master_row_casp_sets = {}
    master_row_hash = {}
    master_hash_rows = {}

    master_package_col = column_map["Package"]
    master_casp_col = column_map["CASP Instance IDs"]
    master_hash_col = column_map["Package Hash"]

    # MASTER DB REAL-PACKAGE APPEND TEST - 2026-09-02
    # openpyxl's ws.max_row can include visually empty rows that only carry
    # formatting/data-validation cells.  Do not let those phantom rows decide
    # where new package records are appended.  The Package column is the
    # authoritative indicator of a real CC Master record.
    last_real_master_row = 1
    for row_num in range(ws.max_row, 1, -1):
        if ws.cell(row_num, master_package_col).value not in (None, ""):
            last_real_master_row = row_num
            break

    phantom_trailing_rows = max(ws.max_row - last_real_master_row, 0)
    if phantom_trailing_rows:
        print(
            f"MASTER DB: ignoring {phantom_trailing_rows:,} formatted/empty "
            f"trailing CC Master rows when choosing append position"
        )

    for row_num in range(2, last_real_master_row + 1):

        package_name = ws.cell(row_num, master_package_col).value
        casp_cell = ws.cell(row_num, master_casp_col).value or ""
        package_hash = ws.cell(row_num, master_hash_col).value or ""

        if package_name:
            master_package_rows.setdefault(package_name, []).append(row_num)
            master_row_package[row_num] = package_name

        casp_set = {
            cid
            for cid in str(casp_cell).split("\n")
            if cid
        }
        master_row_casp_sets[row_num] = casp_set

        # Match the old dictionary semantics: later worksheet rows win for
        # a CASP ID during cross-filename CASP duplicate lookup.
        for cid in casp_set:
            existing_casp[cid] = row_num

        master_row_hash[row_num] = package_hash
        if package_hash:
            master_hash_rows.setdefault(package_hash, []).append(row_num)

    print(
        f"MASTER DB: indexed through real row {last_real_master_row:,} "
        f"in {time.perf_counter() - index_started:.1f}s"
    )

    # -------------------------
    # Add / Update entries
    # -------------------------

    total_master_updates = len(rows)
    next_master_row = last_real_master_row + 1

    for r_index, r in enumerate(rows, start=1):

        if r_index == 1 or r_index % 500 == 0 or r_index == total_master_updates:
            print(f"MASTER DB: CC Master {r_index:,}/{total_master_updates:,}")

        package = r[0]

        detected_package_format = detect_package_format(package)
        subtype = detect_subtype(
            package,
            r[6],
            detected_package_format,
            r[142] if len(r) > 142 else ""
        )
        package_format = detected_package_format
        if package_format == "Single Item":
            package_format = ""

        casp_ids = r[1]
        incoming_casp_set = {
            cid
            for cid in str(casp_ids).split("\n")
            if cid
        } if casp_ids else set()
        incoming_hash = r[21] if len(r) > 21 else ""
        original_source_path = r[138] if len(r) > 138 else ""

        matched_row = None

        # Try matching package name first. Candidate rows are kept in worksheet
        # order, preserving the old first-qualifying-row behaviour.
        for row_num in master_package_rows.get(package, []):

            existing_casp_set = master_row_casp_sets.get(row_num, set())

            # If either package has CASP data, only treat the filename match as
            # the same package when they share a CASP ID.
            if incoming_casp_set or existing_casp_set:
                if incoming_casp_set.intersection(existing_casp_set):
                    matched_row = row_num
                    break
                continue

            # NON-CASP PACKAGE IDENTITY: same filename still requires hash.
            existing_hash = master_row_hash.get(row_num, "")
            if incoming_hash and existing_hash:
                if incoming_hash == existing_hash:
                    matched_row = row_num
                    break
                continue

            # Do not merge a legacy unhashed row merely because the filename
            # is the same.
            continue

        # Cross-filename CASP duplicate lookup.
        if matched_row is None and casp_ids:
            for cid in str(casp_ids).split("\n"):
                if cid in existing_casp:
                    matched_row = existing_casp[cid]
                    break

        # Cross-filename NON-CASP duplicate lookup.  Hash lists are kept in
        # worksheet order, so the first row matches the old full-sheet scan.
        if matched_row is None and not casp_ids and incoming_hash:
            hash_rows = master_hash_rows.get(incoming_hash, [])
            if hash_rows:
                matched_row = hash_rows[0]

        if matched_row:

                existing_package_name = ws.cell(
                    matched_row,
                    column_map["Package"]
                ).value or package

                # Different filename matched through package identity:
                # record it as a duplicate without overwriting
                # the main package's scan information.
                if package != existing_package_name:

                    old_files = ws.cell(
                        matched_row,
                        column_map["Duplicate Files"]
                    ).value or ""

                    if package not in old_files.split("\n"):

                        if old_files:
                            old_files += "\n"

                        old_files += package

                        ws.cell(
                            matched_row,
                            column_map["Duplicate Files"]
                        ).value = old_files

                    if original_source_path:
                        _old_source_paths = ws.cell(
                            matched_row,
                            column_map["Original Source Path"]
                        ).value or ""
                        ws.cell(
                            matched_row,
                            column_map["Original Source Path"]
                        ).value = append_unique_source_path(
                            _old_source_paths,
                            original_source_path
                        )

                    duplicate_thumbnail = thumbnail_paths_by_hash.get(
                        str(incoming_hash or "").lower(),
                        ""
                    )
                    if duplicate_thumbnail and "Image" in column_map:
                        _set_thumbnail_link(
                            ws.cell(matched_row, column_map["Image"]),
                            duplicate_thumbnail
                        )

                    continue

                # Refresh scanned information. Keep old identity metadata so
                # the live indexes can be updated without re-scanning the sheet.
                old_casp_set = master_row_casp_sets.get(matched_row, set()).copy()
                old_hash = master_row_hash.get(matched_row, "")

                ws.cell(
                    matched_row,
                    column_map["Normalized Package"]
                ).value = normalize_package_name(
                    existing_package_name
                )

                ws.cell(matched_row, column_map["CASP Instance IDs"]).value = r[1]
                ws.cell(matched_row, column_map["GEOM Instance IDs"]).value = r[2]
                ws.cell(matched_row, column_map["IMG Instance IDs"]).value = r[3]
                ws.cell(matched_row, column_map["VPXY Instance IDs"]).value = r[4]

                ws.cell(matched_row, column_map["CC Type"]).value = _display_cc_type_with_subtype(r[6], subtype)
                ws.cell(matched_row, column_map["Subtype"]).value = subtype
                ws.cell(matched_row, column_map["Package Format"]).value = package_format
                if original_source_path:
                    _old_source_paths = ws.cell(
                        matched_row,
                        column_map["Original Source Path"]
                    ).value or ""
                    ws.cell(
                        matched_row,
                        column_map["Original Source Path"]
                    ).value = append_unique_source_path(
                        _old_source_paths,
                        original_source_path
                    )
                ws.cell(matched_row, column_map["Content Type"]).value = _display_content_type(r[7])
                ws.cell(matched_row, column_map["Risk"]).value = r[8]
                ws.cell(matched_row, column_map["Risk Score"]).value = r[9]
                ws.cell(matched_row, column_map["Health Status"]).value = r[22]
                ws.cell(matched_row, column_map["Broken GEOM Count"]).value = r[23]
                ws.cell(matched_row, column_map["Broken GEOM IDs"]).value = r[24]
                ws.cell(matched_row, column_map["Broken CASP Count"]).value = r[26]
                ws.cell(matched_row, column_map["Broken CASP IDs"]).value = r[27]
                ws.cell(matched_row, column_map["Broken IMG Count"]).value = r[28]
                ws.cell(matched_row, column_map["Broken IMG IDs"]).value = r[29]
                ws.cell(matched_row, column_map["Broken VPXY Count"]).value = r[30]
                ws.cell(matched_row, column_map["Broken VPXY IDs"]).value = r[31]
                ws.cell(matched_row, column_map["Broken BONE Count"]).value = r[32]
                ws.cell(matched_row, column_map["Broken BONE IDs"]).value = r[33]
                ws.cell(matched_row, column_map["Broken BOND Count"]).value = r[34]
                ws.cell(matched_row, column_map["Broken BOND IDs"]).value = r[35]
                ws.cell(matched_row, column_map["Broken Morph Count"]).value = r[36]
                ws.cell(matched_row, column_map["Broken Morph IDs"]).value = r[37]
                ws.cell(matched_row, column_map["Broken Build/Buy Count"]).value = r[38]
                ws.cell(matched_row, column_map["Broken Build/Buy IDs"]).value = r[39]
                ws.cell(matched_row, column_map["Broken PRESET Count"]).value = r[40]
                ws.cell(matched_row, column_map["Broken PRESET IDs"]).value = r[41]
                ws.cell(matched_row, column_map["Broken TONE Count"]).value = r[42]
                ws.cell(matched_row, column_map["Broken TONE IDs"]).value = r[43]
                ws.cell(matched_row, column_map["Broken IMAG Count"]).value = r[44]
                ws.cell(matched_row, column_map["Broken IMAG IDs"]).value = r[45]
                ws.cell(matched_row, column_map["Broken TXTC Count"]).value = r[46]
                ws.cell(matched_row, column_map["Broken TXTC IDs"]).value = r[47]
                ws.cell(matched_row, column_map["Broken PTRN Count"]).value = r[48]
                ws.cell(matched_row, column_map["Broken PTRN IDs"]).value = r[49]
                ws.cell(matched_row, column_map["Broken SIMO Count"]).value = r[50]
                ws.cell(matched_row, column_map["Broken SIMO IDs"]).value = r[51]
                ws.cell(matched_row, column_map["Broken TXTF Count"]).value = r[52]
                ws.cell(matched_row, column_map["Broken TXTF IDs"]).value = r[53]
                ws.cell(matched_row, column_map["Broken FAMD Count"]).value = r[54]
                ws.cell(matched_row, column_map["Broken FAMD IDs"]).value = r[55]
                ws.cell(matched_row, column_map["Broken OBJD Count"]).value = r[56]
                ws.cell(matched_row, column_map["Broken OBJD IDs"]).value = r[57]
                ws.cell(matched_row, column_map["Broken OBJK Count"]).value = r[58]
                ws.cell(matched_row, column_map["Broken OBJK IDs"]).value = r[59]
                ws.cell(matched_row, column_map["Broken XML Count"]).value = r[60]
                ws.cell(matched_row, column_map["Broken XML IDs"]).value = r[61]
                ws.cell(matched_row, column_map["Broken STBL Count"]).value = r[62]
                ws.cell(matched_row, column_map["Broken STBL IDs"]).value = r[63]
                ws.cell(matched_row, column_map["Broken NMAP Count"]).value = r[64]
                ws.cell(matched_row, column_map["Broken NMAP IDs"]).value = r[65]
                ws.cell(matched_row, column_map["Broken RSLT Count"]).value = r[66]
                ws.cell(matched_row, column_map["Broken RSLT IDs"]).value = r[67]
                ws.cell(matched_row, column_map["Broken FTPT Count"]).value = r[68]
                ws.cell(matched_row, column_map["Broken FTPT IDs"]).value = r[69]
                ws.cell(matched_row, column_map["Broken RIG Count"]).value = r[70]
                ws.cell(matched_row, column_map["Broken RIG IDs"]).value = r[71]
                ws.cell(matched_row, column_map["Broken JAZZ Count"]).value = r[72]
                ws.cell(matched_row, column_map["Broken JAZZ IDs"]).value = r[73]
                ws.cell(matched_row, column_map["Broken S3SA Count"]).value = r[74]
                ws.cell(matched_row, column_map["Broken S3SA IDs"]).value = r[75]
                ws.cell(matched_row, column_map["Broken LITE Count"]).value = r[76]
                ws.cell(matched_row, column_map["Broken LITE IDs"]).value = r[77]
                ws.cell(matched_row, column_map["Broken CLIP Count"]).value = r[78]
                ws.cell(matched_row, column_map["Broken CLIP IDs"]).value = r[79]
                ws.cell(matched_row, column_map["Broken MATD Count"]).value = r[80]
                ws.cell(matched_row, column_map["Broken MATD IDs"]).value = r[81]
                ws.cell(matched_row, column_map["Broken MTST Count"]).value = r[82]
                ws.cell(matched_row, column_map["Broken MTST IDs"]).value = r[83]
                ws.cell(matched_row, column_map["Broken MDLR Count"]).value = r[84]
                ws.cell(matched_row, column_map["Broken MDLR IDs"]).value = r[85]
                ws.cell(matched_row, column_map["Broken CWAL Count"]).value = r[86]
                ws.cell(matched_row, column_map["Broken CWAL IDs"]).value = r[87]
                ws.cell(matched_row, column_map["Broken CFEN Count"]).value = r[88]
                ws.cell(matched_row, column_map["Broken CFEN IDs"]).value = r[89]
                ws.cell(matched_row, column_map["Broken CTPT Count"]).value = r[90]
                ws.cell(matched_row, column_map["Broken CTPT IDs"]).value = r[91]
                ws.cell(matched_row, column_map["Broken NGMP Count"]).value = r[92]
                ws.cell(matched_row, column_map["Broken NGMP IDs"]).value = r[93]
                ws.cell(matched_row, column_map["Broken CRMT Count"]).value = r[94]
                ws.cell(matched_row, column_map["Broken CRMT IDs"]).value = r[95]
                ws.cell(matched_row, column_map["Broken SPT2 Count"]).value = r[96]
                ws.cell(matched_row, column_map["Broken SPT2 IDs"]).value = r[97]
                ws.cell(matched_row, column_map["Broken _SPT Count"]).value = r[98]
                ws.cell(matched_row, column_map["Broken _SPT IDs"]).value = r[99]
                ws.cell(matched_row, column_map["Broken CRAL Count"]).value = r[100]
                ws.cell(matched_row, column_map["Broken CRAL IDs"]).value = r[101]
                ws.cell(matched_row, column_map["Broken LAYO Count"]).value = r[102]
                ws.cell(matched_row, column_map["Broken LAYO IDs"]).value = r[103]
                ws.cell(matched_row, column_map["Broken HINF Count"]).value = r[104]
                ws.cell(matched_row, column_map["Broken HINF IDs"]).value = r[105]
                ws.cell(matched_row, column_map["Broken 2ARY Count"]).value = r[106]
                ws.cell(matched_row, column_map["Broken 2ARY IDs"]).value = r[107]
                ws.cell(matched_row, column_map["Broken Complate Count"]).value = r[108]
                ws.cell(matched_row, column_map["Broken Complate IDs"]).value = r[109]
                ws.cell(matched_row, column_map["Broken 63A33EA7 Count"]).value = r[110]
                ws.cell(matched_row, column_map["Broken 63A33EA7 IDs"]).value = r[111]
                ws.cell(matched_row, column_map["Broken Package Storage Count"]).value = r[112]
                ws.cell(matched_row, column_map["Broken Package Storage IDs"]).value = r[113]
                ws.cell(matched_row, column_map["Broken Fireplace Count"]).value = r[114]
                ws.cell(matched_row, column_map["Broken Fireplace IDs"]).value = r[115]
                ws.cell(matched_row, column_map["Broken THUM Count"]).value = r[116]
                ws.cell(matched_row, column_map["Broken THUM IDs"]).value = r[117]
                ws.cell(matched_row, column_map["Broken 626F60CE Count"]).value = r[118]
                ws.cell(matched_row, column_map["Broken 626F60CE IDs"]).value = r[119]
                ws.cell(matched_row, column_map["Broken ICON Count"]).value = r[120]
                ws.cell(matched_row, column_map["Broken ICON IDs"]).value = r[121]
                ws.cell(matched_row, column_map["Broken THUM 0580A2B5 Count"]).value = r[122]
                ws.cell(matched_row, column_map["Broken THUM 0580A2B5 IDs"]).value = r[123]
                ws.cell(matched_row, column_map["Broken THUM 0580A2B4 Count"]).value = r[124]
                ws.cell(matched_row, column_map["Broken THUM 0580A2B4 IDs"]).value = r[125]
                ws.cell(matched_row, column_map["Broken THUM 0580A2B6 Count"]).value = r[126]
                ws.cell(matched_row, column_map["Broken THUM 0580A2B6 IDs"]).value = r[127]
                ws.cell(matched_row, column_map["Broken SNAP 0580A2CD Count"]).value = r[128]
                ws.cell(matched_row, column_map["Broken SNAP 0580A2CD IDs"]).value = r[129]
                ws.cell(matched_row, column_map["Broken SNAP 0580A2CF Count"]).value = r[130]
                ws.cell(matched_row, column_map["Broken SNAP 0580A2CF IDs"]).value = r[131]
                ws.cell(matched_row, column_map["Broken TWNI Count"]).value = r[132]
                ws.cell(matched_row, column_map["Broken TWNI IDs"]).value = r[133]
                ws.cell(matched_row, column_map["Broken MANIFEST Count"]).value = r[134]
                ws.cell(matched_row, column_map["Broken MANIFEST IDs"]).value = r[135]
                ws.cell(matched_row, column_map["Broken 8F78F74B Count"]).value = r[136]
                ws.cell(matched_row, column_map["Broken 8F78F74B IDs"]).value = r[137]
                ws.cell(matched_row, column_map["Health Issues"]).value = r[25]
                ws.cell(matched_row, column_map["HQ Warning"]).value = r[10]
                ws.cell(matched_row, column_map["Poly Warning"]).value = r[11]
                ws.cell(matched_row, column_map["Highest Polys"]).value = r[12]
                ws.cell(matched_row, column_map["Biggest Texture"]).value = r[13]
                ws.cell(matched_row, column_map["Skin Type"]).value = r[14]
                ws.cell(matched_row, column_map["Last Scanned"]).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ws.cell(matched_row, column_map["STBL IDs"]).value = r[15]
                ws.cell(matched_row, column_map["BGEO IDs"]).value = r[16]
                ws.cell(matched_row, column_map["FACE IDs"]).value = r[17]
                ws.cell(matched_row, column_map["FBLN IDs"]).value = r[18]
                ws.cell(matched_row, column_map["KEY IDs"]).value = r[19]

                ws.cell(
                    matched_row,
                    column_map["BONE Resource Count"]
                ).value = r[20]

                ws.cell(
                    matched_row,
                    column_map["Package Hash"]
                ).value = r[21] if len(r) > 21 else ""

                current_thumbnail = thumbnail_paths_by_hash.get(
                    str(incoming_hash or "").lower(),
                    ""
                )
                if current_thumbnail and "Image" in column_map:
                    _set_thumbnail_link(
                        ws.cell(matched_row, column_map["Image"]),
                        current_thumbnail
                    )

                # Refresh live CASP identity indexes in O(number of this row's
                # CASPs), rather than scanning every CASP mapping.
                for cid in old_casp_set:
                    if existing_casp.get(cid) == matched_row:
                        del existing_casp[cid]

                master_row_casp_sets[matched_row] = incoming_casp_set.copy()
                for cid in incoming_casp_set:
                    existing_casp[cid] = matched_row

                # Refresh the hash lookup while preserving worksheet-row order.
                new_hash = incoming_hash
                if old_hash != new_hash:
                    if old_hash:
                        old_hash_rows = master_hash_rows.get(old_hash, [])
                        try:
                            old_hash_rows.remove(matched_row)
                        except ValueError:
                            pass
                        if not old_hash_rows:
                            master_hash_rows.pop(old_hash, None)

                    if new_hash:
                        new_hash_rows = master_hash_rows.setdefault(new_hash, [])
                        if matched_row not in new_hash_rows:
                            insort(new_hash_rows, matched_row)

                master_row_hash[matched_row] = new_hash

        # -------------------------
        # Formatting
        # -------------------------


        else:

            # Append after the last REAL package row, not ws.max_row.
            # This safely reuses any formatted-but-empty trailing rows.
            new_row = next_master_row
            next_master_row += 1
            last_real_master_row = new_row
            ws.cell(new_row, column_map["Package"]).value = r[0]

            ws.cell(
                new_row,
                column_map["Normalized Package"]
            ).value = normalize_package_name(r[0])

            ws.cell(new_row, column_map["CASP Instance IDs"]).value = r[1]
            ws.cell(new_row, column_map["GEOM Instance IDs"]).value = r[2]
            ws.cell(new_row, column_map["IMG Instance IDs"]).value = r[3]
            ws.cell(new_row, column_map["VPXY Instance IDs"]).value = r[4]

            ws.cell(new_row, column_map["CC Type"]).value = _display_cc_type_with_subtype(r[6], subtype)
            ws.cell(new_row, column_map["Subtype"]).value = subtype
            ws.cell(new_row, column_map["Package Format"]).value = package_format
            ws.cell(new_row, column_map["Original Source Path"]).value = original_source_path
            ws.cell(new_row, column_map["Content Type"]).value = _display_content_type(r[7])
            ws.cell(new_row, column_map["Risk"]).value = r[8]
            ws.cell(new_row, column_map["Risk Score"]).value = r[9]
            ws.cell(new_row, column_map["Health Status"]).value = r[22]
            ws.cell(new_row, column_map["Broken GEOM Count"]).value = r[23]
            ws.cell(new_row, column_map["Broken GEOM IDs"]).value = r[24]
            ws.cell(new_row, column_map["Broken CASP Count"]).value = r[26]
            ws.cell(new_row, column_map["Broken CASP IDs"]).value = r[27]
            ws.cell(new_row, column_map["Broken IMG Count"]).value = r[28]
            ws.cell(new_row, column_map["Broken IMG IDs"]).value = r[29]
            ws.cell(new_row, column_map["Broken VPXY Count"]).value = r[30]
            ws.cell(new_row, column_map["Broken VPXY IDs"]).value = r[31]
            ws.cell(new_row, column_map["Broken BONE Count"]).value = r[32]
            ws.cell(new_row, column_map["Broken BONE IDs"]).value = r[33]
            ws.cell(new_row, column_map["Broken BOND Count"]).value = r[34]
            ws.cell(new_row, column_map["Broken BOND IDs"]).value = r[35]
            ws.cell(new_row, column_map["Broken Morph Count"]).value = r[36]
            ws.cell(new_row, column_map["Broken Morph IDs"]).value = r[37]
            ws.cell(new_row, column_map["Broken Build/Buy Count"]).value = r[38]
            ws.cell(new_row, column_map["Broken Build/Buy IDs"]).value = r[39]
            ws.cell(new_row, column_map["Broken PRESET Count"]).value = r[40]
            ws.cell(new_row, column_map["Broken PRESET IDs"]).value = r[41]
            ws.cell(new_row, column_map["Broken TONE Count"]).value = r[42]
            ws.cell(new_row, column_map["Broken TONE IDs"]).value = r[43]
            ws.cell(new_row, column_map["Broken IMAG Count"]).value = r[44]
            ws.cell(new_row, column_map["Broken IMAG IDs"]).value = r[45]
            ws.cell(new_row, column_map["Broken TXTC Count"]).value = r[46]
            ws.cell(new_row, column_map["Broken TXTC IDs"]).value = r[47]
            ws.cell(new_row, column_map["Broken PTRN Count"]).value = r[48]
            ws.cell(new_row, column_map["Broken PTRN IDs"]).value = r[49]
            ws.cell(new_row, column_map["Broken SIMO Count"]).value = r[50]
            ws.cell(new_row, column_map["Broken SIMO IDs"]).value = r[51]
            ws.cell(new_row, column_map["Broken TXTF Count"]).value = r[52]
            ws.cell(new_row, column_map["Broken TXTF IDs"]).value = r[53]
            ws.cell(new_row, column_map["Broken FAMD Count"]).value = r[54]
            ws.cell(new_row, column_map["Broken FAMD IDs"]).value = r[55]
            ws.cell(new_row, column_map["Broken OBJD Count"]).value = r[56]
            ws.cell(new_row, column_map["Broken OBJD IDs"]).value = r[57]
            ws.cell(new_row, column_map["Broken OBJK Count"]).value = r[58]
            ws.cell(new_row, column_map["Broken OBJK IDs"]).value = r[59]
            ws.cell(new_row, column_map["Broken XML Count"]).value = r[60]
            ws.cell(new_row, column_map["Broken XML IDs"]).value = r[61]
            ws.cell(new_row, column_map["Broken STBL Count"]).value = r[62]
            ws.cell(new_row, column_map["Broken STBL IDs"]).value = r[63]
            ws.cell(new_row, column_map["Broken NMAP Count"]).value = r[64]
            ws.cell(new_row, column_map["Broken NMAP IDs"]).value = r[65]
            ws.cell(new_row, column_map["Broken RSLT Count"]).value = r[66]
            ws.cell(new_row, column_map["Broken RSLT IDs"]).value = r[67]
            ws.cell(new_row, column_map["Broken FTPT Count"]).value = r[68]
            ws.cell(new_row, column_map["Broken FTPT IDs"]).value = r[69]
            ws.cell(new_row, column_map["Broken RIG Count"]).value = r[70]
            ws.cell(new_row, column_map["Broken RIG IDs"]).value = r[71]
            ws.cell(new_row, column_map["Broken JAZZ Count"]).value = r[72]
            ws.cell(new_row, column_map["Broken JAZZ IDs"]).value = r[73]
            ws.cell(new_row, column_map["Broken S3SA Count"]).value = r[74]
            ws.cell(new_row, column_map["Broken S3SA IDs"]).value = r[75]
            ws.cell(new_row, column_map["Broken LITE Count"]).value = r[76]
            ws.cell(new_row, column_map["Broken LITE IDs"]).value = r[77]
            ws.cell(new_row, column_map["Broken CLIP Count"]).value = r[78]
            ws.cell(new_row, column_map["Broken CLIP IDs"]).value = r[79]
            ws.cell(new_row, column_map["Broken MATD Count"]).value = r[80]
            ws.cell(new_row, column_map["Broken MATD IDs"]).value = r[81]
            ws.cell(new_row, column_map["Broken MTST Count"]).value = r[82]
            ws.cell(new_row, column_map["Broken MTST IDs"]).value = r[83]
            ws.cell(new_row, column_map["Broken MDLR Count"]).value = r[84]
            ws.cell(new_row, column_map["Broken MDLR IDs"]).value = r[85]
            ws.cell(new_row, column_map["Broken CWAL Count"]).value = r[86]
            ws.cell(new_row, column_map["Broken CWAL IDs"]).value = r[87]
            ws.cell(new_row, column_map["Broken CFEN Count"]).value = r[88]
            ws.cell(new_row, column_map["Broken CFEN IDs"]).value = r[89]
            ws.cell(new_row, column_map["Broken CTPT Count"]).value = r[90]
            ws.cell(new_row, column_map["Broken CTPT IDs"]).value = r[91]
            ws.cell(new_row, column_map["Broken NGMP Count"]).value = r[92]
            ws.cell(new_row, column_map["Broken NGMP IDs"]).value = r[93]
            ws.cell(new_row, column_map["Broken CRMT Count"]).value = r[94]
            ws.cell(new_row, column_map["Broken CRMT IDs"]).value = r[95]
            ws.cell(new_row, column_map["Broken SPT2 Count"]).value = r[96]
            ws.cell(new_row, column_map["Broken SPT2 IDs"]).value = r[97]
            ws.cell(new_row, column_map["Broken _SPT Count"]).value = r[98]
            ws.cell(new_row, column_map["Broken _SPT IDs"]).value = r[99]
            ws.cell(new_row, column_map["Broken CRAL Count"]).value = r[100]
            ws.cell(new_row, column_map["Broken CRAL IDs"]).value = r[101]
            ws.cell(new_row, column_map["Broken LAYO Count"]).value = r[102]
            ws.cell(new_row, column_map["Broken LAYO IDs"]).value = r[103]
            ws.cell(new_row, column_map["Broken HINF Count"]).value = r[104]
            ws.cell(new_row, column_map["Broken HINF IDs"]).value = r[105]
            ws.cell(new_row, column_map["Broken 2ARY Count"]).value = r[106]
            ws.cell(new_row, column_map["Broken 2ARY IDs"]).value = r[107]
            ws.cell(new_row, column_map["Broken Complate Count"]).value = r[108]
            ws.cell(new_row, column_map["Broken Complate IDs"]).value = r[109]
            ws.cell(new_row, column_map["Broken 63A33EA7 Count"]).value = r[110]
            ws.cell(new_row, column_map["Broken 63A33EA7 IDs"]).value = r[111]
            ws.cell(new_row, column_map["Broken Package Storage Count"]).value = r[112]
            ws.cell(new_row, column_map["Broken Package Storage IDs"]).value = r[113]
            ws.cell(new_row, column_map["Broken Fireplace Count"]).value = r[114]
            ws.cell(new_row, column_map["Broken Fireplace IDs"]).value = r[115]
            ws.cell(new_row, column_map["Broken THUM Count"]).value = r[116]
            ws.cell(new_row, column_map["Broken THUM IDs"]).value = r[117]
            ws.cell(new_row, column_map["Broken 626F60CE Count"]).value = r[118]
            ws.cell(new_row, column_map["Broken 626F60CE IDs"]).value = r[119]
            ws.cell(new_row, column_map["Broken ICON Count"]).value = r[120]
            ws.cell(new_row, column_map["Broken ICON IDs"]).value = r[121]
            ws.cell(new_row, column_map["Broken THUM 0580A2B5 Count"]).value = r[122]
            ws.cell(new_row, column_map["Broken THUM 0580A2B5 IDs"]).value = r[123]
            ws.cell(new_row, column_map["Broken THUM 0580A2B4 Count"]).value = r[124]
            ws.cell(new_row, column_map["Broken THUM 0580A2B4 IDs"]).value = r[125]
            ws.cell(new_row, column_map["Broken THUM 0580A2B6 Count"]).value = r[126]
            ws.cell(new_row, column_map["Broken THUM 0580A2B6 IDs"]).value = r[127]
            ws.cell(new_row, column_map["Broken SNAP 0580A2CD Count"]).value = r[128]
            ws.cell(new_row, column_map["Broken SNAP 0580A2CD IDs"]).value = r[129]
            ws.cell(new_row, column_map["Broken SNAP 0580A2CF Count"]).value = r[130]
            ws.cell(new_row, column_map["Broken SNAP 0580A2CF IDs"]).value = r[131]
            ws.cell(new_row, column_map["Broken TWNI Count"]).value = r[132]
            ws.cell(new_row, column_map["Broken TWNI IDs"]).value = r[133]
            ws.cell(new_row, column_map["Broken MANIFEST Count"]).value = r[134]
            ws.cell(new_row, column_map["Broken MANIFEST IDs"]).value = r[135]
            ws.cell(new_row, column_map["Broken 8F78F74B Count"]).value = r[136]
            ws.cell(new_row, column_map["Broken 8F78F74B IDs"]).value = r[137]
            ws.cell(new_row, column_map["Health Issues"]).value = r[25]
            ws.cell(new_row, column_map["HQ Warning"]).value = r[10]
            ws.cell(new_row, column_map["Poly Warning"]).value = r[11]
            ws.cell(new_row, column_map["Highest Polys"]).value = r[12]
            ws.cell(new_row, column_map["Biggest Texture"]).value = r[13]
            ws.cell(new_row, column_map["Skin Type"]).value = r[14]
            ws.cell(new_row, column_map["Last Scanned"]).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws.cell(new_row, column_map["STBL IDs"]).value = r[15]
            ws.cell(new_row, column_map["BGEO IDs"]).value = r[16]
            ws.cell(new_row, column_map["FACE IDs"]).value = r[17]
            ws.cell(new_row, column_map["FBLN IDs"]).value = r[18]
            ws.cell(new_row, column_map["KEY IDs"]).value = r[19]

            ws.cell(
                new_row,
                column_map["BONE Resource Count"]
            ).value = r[20]

            ws.cell(
                new_row,
                column_map["Package Hash"]
            ).value = r[21] if len(r) > 21 else ""

            current_thumbnail = thumbnail_paths_by_hash.get(
                str(incoming_hash or "").lower(),
                ""
            )
            if current_thumbnail and "Image" in column_map:
                _set_thumbnail_link(
                    ws.cell(new_row, column_map["Image"]),
                    current_thumbnail
                )

            # Add the new row to every live identity index so later packages
            # in this same scan see exactly the same data the worksheet contains.
            master_package_rows.setdefault(package, []).append(new_row)
            master_row_package[new_row] = package
            master_row_casp_sets[new_row] = incoming_casp_set.copy()
            for cid in incoming_casp_set:
                existing_casp[cid] = new_row

            master_row_hash[new_row] = incoming_hash
            if incoming_hash:
                master_hash_rows.setdefault(incoming_hash, []).append(new_row)

    # -------------------------
    # Formatting
    # -------------------------

    ws.freeze_panes = "A2"

    # Keep Excel AutoFilter attached to the complete current Master DB range.
    # The indexed updater can add rows/columns without touching structural scan logic.
    if last_real_master_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = "A1:{}{}".format(
            get_column_letter(ws.max_column),
            last_real_master_row
        )


    wrap_headers = [
        "Package",
        "Notes",
        "Normalized Package",
        "Content Type",
        "Broken GEOM IDs",
        "Broken CASP Count",
        "Broken CASP IDs",
        "Broken IMG Count",
        "Broken IMG IDs",
        "Broken VPXY Count",
        "Broken VPXY IDs",
        "Broken BONE Count",
        "Broken BONE IDs",
        "Broken BOND Count",
        "Broken BOND IDs",
        "Broken Morph Count",
        "Broken Morph IDs",
        "Broken Build/Buy Count",
        "Broken Build/Buy IDs",
        "Broken PRESET Count",
        "Broken PRESET IDs",
        "Broken TONE Count",
        "Broken TONE IDs",
        "Broken IMAG Count",
        "Broken IMAG IDs",
        "Broken TXTC Count",
        "Broken TXTC IDs",
        "Broken PTRN Count",
        "Broken PTRN IDs",
        "Broken SIMO Count",
        "Broken SIMO IDs",
        "Broken TXTF Count",
        "Broken TXTF IDs",
        "Broken FAMD Count",
        "Broken FAMD IDs",
        "Broken OBJD Count",
        "Broken OBJD IDs",
        "Broken OBJK Count",
        "Broken OBJK IDs",
        "Broken XML Count",
        "Broken XML IDs",
        "Broken STBL Count",
        "Broken STBL IDs",
        "Broken NMAP Count",
        "Broken NMAP IDs",
        "Broken RSLT Count",
        "Broken RSLT IDs",
        "Broken FTPT Count",
        "Broken FTPT IDs",
        "Broken RIG Count",
        "Broken RIG IDs",
        "Broken JAZZ Count",
        "Broken JAZZ IDs",
        "Broken S3SA Count",
        "Broken S3SA IDs",
        "Broken LITE Count",
        "Broken LITE IDs",
        "Broken CLIP Count",
        "Broken CLIP IDs",
        "Broken MATD Count",
        "Broken MATD IDs",
        "Broken MTST Count",
        "Broken MTST IDs",
        "Broken MDLR Count",
        "Broken MDLR IDs",
        "Broken CWAL Count",
        "Broken CWAL IDs",
        "Broken CFEN Count",
        "Broken CFEN IDs",
        "Broken CTPT Count",
        "Broken CTPT IDs",
        "Broken NGMP Count",
        "Broken NGMP IDs",
        "Broken CRMT Count",
        "Broken CRMT IDs",
        "Broken SPT2 Count",
        "Broken SPT2 IDs",
        "Broken _SPT Count",
        "Broken _SPT IDs",
        "Broken CRAL Count",
        "Broken CRAL IDs",
        "Broken LAYO Count",
        "Broken LAYO IDs",
        "Broken HINF Count",
        "Broken HINF IDs",
        "Broken 2ARY Count",
        "Broken 2ARY IDs",
        "Broken Complate Count",
        "Broken Complate IDs",
        "Broken 63A33EA7 Count",
        "Broken 63A33EA7 IDs",
        "Broken Package Storage Count",
        "Broken Package Storage IDs",
        "Broken Fireplace Count",
        "Broken Fireplace IDs",
        "Broken THUM Count",
        "Broken THUM IDs",
        "Broken 626F60CE Count",
        "Broken 626F60CE IDs",
        "Broken ICON Count",
        "Broken ICON IDs",
        "Broken THUM 0580A2B5 Count",
        "Broken THUM 0580A2B5 IDs",
        "Broken THUM 0580A2B4 Count",
        "Broken THUM 0580A2B4 IDs",
        "Broken THUM 0580A2B6 Count",
        "Broken THUM 0580A2B6 IDs",
        "Broken SNAP 0580A2CD Count",
        "Broken SNAP 0580A2CD IDs",
        "Broken SNAP 0580A2CF Count",
        "Broken SNAP 0580A2CF IDs",
        "Broken TWNI Count",
        "Broken TWNI IDs",
        "Broken MANIFEST Count",
        "Broken MANIFEST IDs",
        "Broken 8F78F74B Count",
        "Broken 8F78F74B IDs",
        "Health Issues",
        "CASP Instance IDs",
        "Duplicate Files",
        "STBL IDs",
        "BGEO IDs",
        "FACE IDs",
        "FBLN IDs",
        "KEY IDs",
        "GEOM Instance IDs",
        "IMG Instance IDs",
        "VPXY Instance IDs"
    ]

    # Existing cell formatting is preserved by openpyxl. Reapplying wrap
    # alignment to ~100 full columns on every large scan was one of the major
    # post-scan bottlenecks. Keep headers wrapped; do not rewrite every data
    # cell's style during routine updates.
    for header in wrap_headers:
        if header in column_map:
            ws.cell(1, column_map[header]).alignment = Alignment(
                wrap_text=True,
                vertical="top"
            )


    main_column_widths = {
        "Package": 55,
        "Image": 18,
        "Status": 16,
        "Library Location": 20,
        "CC Magic Profiles": 28,
        "Notes": 35,
        "Normalized Package": 45,
        "CC Type": 18,
        "Content Type": 13,
        "Risk": 14,
        "Risk Score": 12,
        "Health Status": 14,
        "Broken GEOM Count": 18,
        "Broken GEOM IDs": 45,
        "Broken CASP Count": 18,
        "Broken CASP IDs": 45,
        "Broken IMG Count": 18,
        "Broken IMG IDs": 45,
        "Broken VPXY Count": 18,
        "Broken VPXY IDs": 45,
        "Broken BONE Count": 18,
        "Broken BONE IDs": 45,
        "Broken BOND Count": 18,
        "Broken BOND IDs": 45,
        "Broken Morph Count": 18,
        "Broken Morph IDs": 50,
        "Broken Build/Buy Count": 20,
        "Broken Build/Buy IDs": 50,
        "Broken PRESET Count": 20,
        "Broken PRESET IDs": 50,
        "Broken TONE Count": 20,
        "Broken TONE IDs": 50,
        "Broken IMAG Count": 20,
        "Broken IMAG IDs": 50,
        "Broken TXTC Count": 20,
        "Broken TXTC IDs": 50,
        "Broken PTRN Count": 20,
        "Broken PTRN IDs": 50,
        "Broken SIMO Count": 20,
        "Broken SIMO IDs": 50,
        "Broken TXTF Count": 20,
        "Broken TXTF IDs": 50,
        "Broken FAMD Count": 20,
        "Broken FAMD IDs": 50,
        "Broken OBJD Count": 20,
        "Broken OBJD IDs": 50,
        "Broken OBJK Count": 20,
        "Broken OBJK IDs": 50,
        "Broken XML Count": 20,
        "Broken XML IDs": 50,
        "Broken STBL Count": 20,
        "Broken STBL IDs": 50,
        "Broken NMAP Count": 20,
        "Broken NMAP IDs": 50,
        "Broken RSLT Count": 20,
        "Broken RSLT IDs": 50,
        "Broken FTPT Count": 20,
        "Broken FTPT IDs": 50,
        "Broken RIG Count": 20,
        "Broken RIG IDs": 50,
        "Broken JAZZ Count": 20,
        "Broken JAZZ IDs": 50,
        "Broken S3SA Count": 20,
        "Broken S3SA IDs": 50,
        "Broken LITE Count": 20,
        "Broken LITE IDs": 50,
        "Broken CLIP Count": 20,
        "Broken CLIP IDs": 50,
        "Broken MATD Count": 20,
        "Broken MATD IDs": 50,
        "Broken MTST Count": 20,
        "Broken MTST IDs": 50,
        "Broken MDLR Count": 22,
        "Broken MDLR IDs": 50,
        "Broken CWAL Count": 20,
        "Broken CWAL IDs": 50,
        "Broken CFEN Count": 20,
        "Broken CFEN IDs": 50,
        "Broken CTPT Count": 20,
        "Broken CTPT IDs": 50,
        "Broken NGMP Count": 20,
        "Broken NGMP IDs": 50,
        "Broken CRMT Count": 20,
        "Broken CRMT IDs": 50,
        "Broken SPT2 Count": 20,
        "Broken SPT2 IDs": 50,
        "Broken _SPT Count": 20,
        "Broken _SPT IDs": 50,
        "Broken CRAL Count": 20,
        "Broken CRAL IDs": 50,
        "Broken LAYO Count": 20,
        "Broken LAYO IDs": 50,
        "Broken HINF Count": 20,
        "Broken HINF IDs": 50,
        "Broken 2ARY Count": 20,
        "Broken 2ARY IDs": 50,
        "Broken Complate Count": 20,
        "Broken Complate IDs": 50,
        "Broken 63A33EA7 Count": 20,
        "Broken 63A33EA7 IDs": 50,
        "Broken Package Storage Count": 24,
        "Broken Package Storage IDs": 60,
        "Broken Fireplace Count": 20,
        "Broken Fireplace IDs": 50,
        "Broken THUM Count": 20,
        "Broken THUM IDs": 50,
        "Broken 626F60CE Count": 20,
        "Broken 626F60CE IDs": 50,
        "Broken ICON Count": 20,
        "Broken ICON IDs": 50,
        "Broken THUM 0580A2B5 Count": 24,
        "Broken THUM 0580A2B5 IDs": 50,
        "Broken THUM 0580A2B4 Count": 24,
        "Broken THUM 0580A2B4 IDs": 50,
        "Broken THUM 0580A2B6 Count": 24,
        "Broken THUM 0580A2B6 IDs": 50,
        "Broken SNAP 0580A2CD Count": 24,
        "Broken SNAP 0580A2CD IDs": 50,
        "Broken SNAP 0580A2CF Count": 24,
        "Broken SNAP 0580A2CF IDs": 62,
        "Broken TWNI Count": 20,
        "Broken TWNI IDs": 62,
        "Broken MANIFEST Count": 22,
        "Broken MANIFEST IDs": 62,
        "Broken 8F78F74B Count": 22,
        "Broken 8F78F74B IDs": 62,
        "Health Issues": 55,
        "HQ Warning": 16,
        "Poly Warning": 16,
        "Highest Polys": 14,
        "Biggest Texture": 18,
        "Skin Type": 18,
        "Last Scanned": 14,
        "CASP Instance IDs": 18,
        "Duplicate Files": 45,
        "BONE Resource Count": 20,
        "STBL IDs": 45,
        "BGEO IDs": 45,
        "FACE IDs": 45,
        "FBLN IDs": 45,
        "KEY IDs": 45,
        "GEOM Instance IDs": 45,
        "IMG Instance IDs": 45,
        "VPXY Instance IDs": 45
    }
    for header, width in main_column_widths.items():

        if header in column_map:

            column_letter = ws.cell(
                1,
                column_map[header]
            ).column_letter

            ws.column_dimensions[column_letter].width = width

    # -------------------------
    # CASP MASTER SHEET
    # -------------------------

    if "CASP Master" in wb.sheetnames:

        casp_ws = wb["CASP Master"]

    else:

        casp_ws = wb.create_sheet(
            "CASP Master"
        )

        casp_ws.append([
            "CASP Instance",
            "Package",
            "Normalized Package",
            "CC Type",
            "Data Type",
            "Age Gender",
            "Category",
            "Body Type",
            "Overlay Priority"
        ])


    # -------------------------
    # ENSURE CASP MASTER COLUMNS EXIST
    # -------------------------

    casp_column_map = get_column_map(casp_ws)

    required_casp_columns = [
        "CASP Instance",
        "Package",
        "Normalized Package",
        "CC Type",
        "Data Type",
        "Age Gender",
        "Category",
        "Body Type",
        "Overlay Priority"
    ]

    for header in required_casp_columns:

        if header not in casp_column_map:

            casp_ws.cell(
                1,
                casp_ws.max_column + 1
            ).value = header

    casp_column_map = get_column_map(casp_ws)


    existing_casp_records = {}

    for row_num, row in enumerate(
        casp_ws.iter_rows(
            min_row=2,
            values_only=True
        ),
        start=2
    ):

        casp_id = row[
            casp_column_map["CASP Instance"] - 1
        ]

        if casp_id:

            existing_casp_records[casp_id] = row_num


    # -------------------------
    # MASTER CASP OWNERSHIP
    # -------------------------

    master_casp_owner = {}

    # Reuse the already-built Master identity indexes.
    for master_row_num, master_casp_set in master_row_casp_sets.items():
        master_package = master_row_package.get(master_row_num)
        if master_package:
            for casp_id in master_casp_set:
                master_casp_owner[casp_id] = master_package

    print("MASTER DB: updating CASP Master...")
    total_casp_updates = len(casp_records)
    casp_rows_touched = set()

    for casp_index, record in enumerate(casp_records, start=1):

        if casp_index == 1 or casp_index % 1000 == 0 or casp_index == total_casp_updates:
            print(f"MASTER DB: CASP Master {casp_index:,}/{total_casp_updates:,}")

        casp_id = record["CASP Instance"]
        record_package = record["Package"]

        owner_package = master_casp_owner.get(casp_id)

        # A different package sharing this CASP is a duplicate.
        # Do not let it replace the main package in CASP Master.
        if (
            owner_package
            and record_package != owner_package
        ):
            continue

        if casp_id in existing_casp_records:

            casp_row = existing_casp_records[casp_id]

            casp_ws.cell(
                casp_row,
                casp_column_map["Package"]
            ).value = record["Package"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Normalized Package"]
            ).value = normalize_package_name(
                record["Package"]
            )

            casp_ws.cell(
                casp_row,
                casp_column_map["CC Type"]
            ).value = record["CC Type"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Data Type"]
            ).value = record["Data Type"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Age Gender"]
            ).value = record["Age Gender"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Category"]
            ).value = record["Category"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Body Type"]
            ).value = record["Body Type"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Overlay Priority"]
            ).value = record["Overlay Priority"]

        else:

            casp_row = casp_ws.max_row + 1

            casp_ws.cell(
                casp_row,
                casp_column_map["CASP Instance"]
            ).value = record["CASP Instance"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Package"]
            ).value = record["Package"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Normalized Package"]
            ).value = normalize_package_name(
                record["Package"]
            )

            casp_ws.cell(
                casp_row,
                casp_column_map["CC Type"]
            ).value = record["CC Type"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Data Type"]
            ).value = record["Data Type"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Age Gender"]
            ).value = record["Age Gender"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Category"]
            ).value = record["Category"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Body Type"]
            ).value = record["Body Type"]

            casp_ws.cell(
                casp_row,
                casp_column_map["Overlay Priority"]
            ).value = record["Overlay Priority"]

            existing_casp_records[casp_id] = casp_row

        # Count distinct CASP Master worksheet rows actually written during
        # this run. Duplicate-owner records skipped above are intentionally
        # not counted.
        casp_rows_touched.add(casp_row)


    casp_ws.freeze_panes = "A2"

    # Expand/restore the filter range after CASP Master updates.
    if casp_ws.max_row >= 1 and casp_ws.max_column >= 1:
        casp_ws.auto_filter.ref = "A1:{}{}".format(
            get_column_letter(casp_ws.max_column),
            casp_ws.max_row
        )

    _apply_casp_master_compact_layout(casp_ws)

    # -------------------------
    # UPDATE PACKAGES SHEET
    # -------------------------

    print("MASTER DB: indexing Packages sheet...")
    packages_index_started = time.perf_counter()

    existing_packages = {}
    packages_row_casp_sets = {}
    packages_row_hash = {}

    packages_package_col = packages_column_map["Package"]
    packages_casp_col = packages_column_map["CASP IDs"]
    packages_hash_col = packages_column_map["Package Hash"]

    for row_num in range(2, packages_ws.max_row + 1):
        package_name = packages_ws.cell(row_num, packages_package_col).value
        existing_casp_ids = packages_ws.cell(row_num, packages_casp_col).value or ""
        existing_hash = packages_ws.cell(row_num, packages_hash_col).value or ""

        if package_name:
            existing_packages.setdefault(package_name, []).append(row_num)

        packages_row_casp_sets[row_num] = {
            cid
            for cid in str(existing_casp_ids).split("\n")
            if cid
        }
        packages_row_hash[row_num] = existing_hash

    print(
        f"MASTER DB: indexed {max(packages_ws.max_row - 1, 0):,} Packages rows "
        f"in {time.perf_counter() - packages_index_started:.1f}s"
    )

    if DEBUG:
        print("EXISTING PACKAGES:")
        print(existing_packages)

    print("MASTER DB: updating Packages sheet...")
    total_package_updates = len(rows)
    package_rows_touched = set()

    for packages_index, row in enumerate(rows, start=1):

        if packages_index == 1 or packages_index % 500 == 0 or packages_index == total_package_updates:
            print(f"MASTER DB: Packages {packages_index:,}/{total_package_updates:,}")

        detected_package_format = detect_package_format(row[0])

        subtype = detect_subtype(
            row[0],
            row[6],
            detected_package_format,
            row[142] if len(row) > 142 else ""
        )

        package_format = detected_package_format

        if package_format == "Single Item":
            package_format = ""

        original_source_path = row[138] if len(row) > 138 else ""

        package_row = None

        candidate_rows = existing_packages.get(
            row[0],
            []
        )

        incoming_casp_ids = {
            cid
            for cid in str(row[1]).split("\n")
            if cid
        }

        for candidate_row in candidate_rows:

            existing_casp_set = packages_row_casp_sets.get(
                candidate_row,
                set()
            )

            # CASP packages with the same filename are only
            # the same package when their CASP IDs overlap.
            if incoming_casp_ids or existing_casp_set:

                if incoming_casp_ids.intersection(
                    existing_casp_set
                ):
                    package_row = candidate_row
                    break

                continue

            # -------------------------
            # NON-CASP PACKAGE IDENTITY
            # -------------------------

            incoming_hash = (
                row[21]
                if len(row) > 21
                else ""
            )

            existing_hash = packages_row_hash.get(candidate_row, "")

            if incoming_hash and existing_hash:

                if incoming_hash == existing_hash:
                    package_row = candidate_row
                    break

                continue

            # Do not merge a legacy unhashed row merely because
            # another package has the same filename.
            continue

        if package_row is not None:

            packages_ws.cell(
                package_row,
                packages_column_map["Normalized Package"]
            ).value = normalize_package_name(row[0])

            packages_ws.cell(
                package_row,
                packages_column_map["Scan Date"]
            ).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            packages_ws.cell(
                package_row,
                packages_column_map["Risk Level"]
            ).value = row[8]

            packages_ws.cell(
                package_row,
                packages_column_map["Risk Score"]
            ).value = row[9]

            packages_ws.cell(
                package_row,
                packages_column_map["Health Status"]
            ).value = row[22]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken GEOM Count"]
            ).value = row[23]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken GEOM IDs"]
            ).value = row[24]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CASP Count"]
            ).value = row[26]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CASP IDs"]
            ).value = row[27]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken IMG Count"]
            ).value = row[28]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken IMG IDs"]
            ).value = row[29]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken VPXY Count"]
            ).value = row[30]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken VPXY IDs"]
            ).value = row[31]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken BONE Count"]
            ).value = row[32]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken BONE IDs"]
            ).value = row[33]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken BOND Count"]
            ).value = row[34]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken BOND IDs"]
            ).value = row[35]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Morph Count"]
            ).value = row[36]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Morph IDs"]
            ).value = row[37]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Build/Buy Count"]
            ).value = row[38]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Build/Buy IDs"]
            ).value = row[39]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken PRESET Count"]
            ).value = row[40]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken PRESET IDs"]
            ).value = row[41]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TONE Count"]
            ).value = row[42]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TONE IDs"]
            ).value = row[43]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken IMAG Count"]
            ).value = row[44]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken IMAG IDs"]
            ).value = row[45]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TXTC Count"]
            ).value = row[46]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TXTC IDs"]
            ).value = row[47]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken PTRN Count"]
            ).value = row[48]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken PTRN IDs"]
            ).value = row[49]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SIMO Count"]
            ).value = row[50]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SIMO IDs"]
            ).value = row[51]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TXTF Count"]
            ).value = row[52]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TXTF IDs"]
            ).value = row[53]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken FAMD Count"]
            ).value = row[54]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken FAMD IDs"]
            ).value = row[55]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken OBJD Count"]
            ).value = row[56]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken OBJD IDs"]
            ).value = row[57]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken OBJK Count"]
            ).value = row[58]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken OBJK IDs"]
            ).value = row[59]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken XML Count"]
            ).value = row[60]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken XML IDs"]
            ).value = row[61]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken STBL Count"]
            ).value = row[62]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken STBL IDs"]
            ).value = row[63]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken NMAP Count"]
            ).value = row[64]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken NMAP IDs"]
            ).value = row[65]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken RSLT Count"]
            ).value = row[66]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken RSLT IDs"]
            ).value = row[67]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken FTPT Count"]
            ).value = row[68]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken FTPT IDs"]
            ).value = row[69]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken RIG Count"]
            ).value = row[70]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken RIG IDs"]
            ).value = row[71]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken JAZZ Count"]
            ).value = row[72]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken JAZZ IDs"]
            ).value = row[73]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken S3SA Count"]
            ).value = row[74]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken S3SA IDs"]
            ).value = row[75]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken LITE Count"]
            ).value = row[76]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken LITE IDs"]
            ).value = row[77]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CLIP Count"]
            ).value = row[78]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CLIP IDs"]
            ).value = row[79]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MATD Count"]
            ).value = row[80]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MATD IDs"]
            ).value = row[81]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MTST Count"]
            ).value = row[82]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MTST IDs"]
            ).value = row[83]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MDLR Count"]
            ).value = row[84]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MDLR IDs"]
            ).value = row[85]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CWAL Count"]
            ).value = row[86]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CWAL IDs"]
            ).value = row[87]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CFEN Count"]
            ).value = row[88]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CFEN IDs"]
            ).value = row[89]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CTPT Count"]
            ).value = row[90]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CTPT IDs"]
            ).value = row[91]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken NGMP Count"]
            ).value = row[92]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken NGMP IDs"]
            ).value = row[93]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CRMT Count"]
            ).value = row[94]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CRMT IDs"]
            ).value = row[95]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SPT2 Count"]
            ).value = row[96]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SPT2 IDs"]
            ).value = row[97]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken _SPT Count"]
            ).value = row[98]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken _SPT IDs"]
            ).value = row[99]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CRAL Count"]
            ).value = row[100]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CRAL IDs"]
            ).value = row[101]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken LAYO Count"]
            ).value = row[102]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken LAYO IDs"]
            ).value = row[103]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken HINF Count"]
            ).value = row[104]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken HINF IDs"]
            ).value = row[105]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 2ARY Count"]
            ).value = row[106]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 2ARY IDs"]
            ).value = row[107]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Complate Count"]
            ).value = row[108]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Complate IDs"]
            ).value = row[109]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 63A33EA7 Count"]
            ).value = row[110]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 63A33EA7 IDs"]
            ).value = row[111]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Package Storage Count"]
            ).value = row[112]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Package Storage IDs"]
            ).value = row[113]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Fireplace Count"]
            ).value = row[114]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Fireplace IDs"]
            ).value = row[115]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM Count"]
            ).value = row[116]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM IDs"]
            ).value = row[117]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 626F60CE Count"]
            ).value = row[118]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 626F60CE IDs"]
            ).value = row[119]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken ICON Count"]
            ).value = row[120]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken ICON IDs"]
            ).value = row[121]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B5 Count"]
            ).value = row[122]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B5 IDs"]
            ).value = row[123]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B4 Count"]
            ).value = row[124]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B4 IDs"]
            ).value = row[125]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B6 Count"]
            ).value = row[126]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B6 IDs"]
            ).value = row[127]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SNAP 0580A2CD Count"]
            ).value = row[128]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SNAP 0580A2CD IDs"]
            ).value = row[129]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SNAP 0580A2CF Count"]
            ).value = row[130]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SNAP 0580A2CF IDs"]
            ).value = row[131]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TWNI Count"]
            ).value = row[132]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TWNI IDs"]
            ).value = row[133]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MANIFEST Count"]
            ).value = row[134]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MANIFEST IDs"]
            ).value = row[135]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 8F78F74B Count"]
            ).value = row[136]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 8F78F74B IDs"]
            ).value = row[137]

            packages_ws.cell(
                package_row,
                packages_column_map["Health Issues"]
            ).value = row[25]

            packages_ws.cell(
                package_row,
                packages_column_map["CC Type"]
            ).value = row[6]

            packages_ws.cell(
                package_row,
                packages_column_map["Subtype"]
            ).value = subtype

            packages_ws.cell(
                package_row,
                packages_column_map["Package Format"]
            ).value = package_format

            if original_source_path:
                _old_source_paths = packages_ws.cell(
                    package_row,
                    packages_column_map["Original Source Path"]
                ).value or ""
                packages_ws.cell(
                    package_row,
                    packages_column_map["Original Source Path"]
                ).value = append_unique_source_path(
                    _old_source_paths,
                    original_source_path
                )

            packages_ws.cell(
                package_row,
                packages_column_map["CASP IDs"]
            ).value = row[1]

            packages_ws.cell(
                package_row,
                packages_column_map["GEOM IDs"]
            ).value = row[2]

            packages_ws.cell(
                package_row,
                packages_column_map["IMG IDs"]
            ).value = row[3]

            packages_ws.cell(
                package_row,
                packages_column_map["VPXY IDs"]
            ).value = row[4]

            packages_ws.cell(
                package_row,
                packages_column_map["BONE Resource Count"]
            ).value = row[20]

            packages_ws.cell(
                package_row,
                packages_column_map["STBL IDs"]
            ).value = row[15]

            packages_ws.cell(
                package_row,
                packages_column_map["BGEO IDs"]
            ).value = row[16]

            packages_ws.cell(
                package_row,
                packages_column_map["FACE IDs"]
            ).value = row[17]

            packages_ws.cell(
                package_row,
                packages_column_map["FBLN IDs"]
            ).value = row[18]

            packages_ws.cell(
                package_row,
                packages_column_map["KEY IDs"]
            ).value = row[19]

            packages_ws.cell(
                package_row,
                packages_column_map["Highest Poly"]
            ).value = row[12]

            packages_ws.cell(
                package_row,
                packages_column_map["Highest Texture"]
            ).value = row[13]
            
            packages_ws.cell(
                package_row,
                packages_column_map["Package Hash"]
            ).value = row[21] if len(row) > 21 else ""

            package_thumbnail = thumbnail_paths_by_hash.get(
                str((row[21] if len(row) > 21 else "") or "").lower(),
                ""
            )
            if package_thumbnail and "Image" in packages_column_map:
                _set_thumbnail_link(
                    packages_ws.cell(package_row, packages_column_map["Image"]),
                    package_thumbnail
                )

            packages_row_casp_sets[package_row] = incoming_casp_ids.copy()
            packages_row_hash[package_row] = row[21] if len(row) > 21 else ""

        else:

            package_row = packages_ws.max_row + 1

            packages_ws.cell(
                package_row,
                packages_column_map["Package"]
            ).value = row[0]

            packages_ws.cell(
                package_row,
                packages_column_map["Normalized Package"]
            ).value = normalize_package_name(row[0])

            packages_ws.cell(
                package_row,
                packages_column_map["Scan Date"]
            ).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            packages_ws.cell(
                package_row,
                packages_column_map["Risk Level"]
            ).value = row[8]

            packages_ws.cell(
                package_row,
                packages_column_map["Risk Score"]
            ).value = row[9]

            packages_ws.cell(
                package_row,
                packages_column_map["Health Status"]
            ).value = row[22]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken GEOM Count"]
            ).value = row[23]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken GEOM IDs"]
            ).value = row[24]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CASP Count"]
            ).value = row[26]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CASP IDs"]
            ).value = row[27]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken IMG Count"]
            ).value = row[28]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken IMG IDs"]
            ).value = row[29]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken VPXY Count"]
            ).value = row[30]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken VPXY IDs"]
            ).value = row[31]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken BONE Count"]
            ).value = row[32]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken BONE IDs"]
            ).value = row[33]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken BOND Count"]
            ).value = row[34]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken BOND IDs"]
            ).value = row[35]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Morph Count"]
            ).value = row[36]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Morph IDs"]
            ).value = row[37]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Build/Buy Count"]
            ).value = row[38]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Build/Buy IDs"]
            ).value = row[39]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken PRESET Count"]
            ).value = row[40]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken PRESET IDs"]
            ).value = row[41]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TONE Count"]
            ).value = row[42]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TONE IDs"]
            ).value = row[43]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken IMAG Count"]
            ).value = row[44]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken IMAG IDs"]
            ).value = row[45]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TXTC Count"]
            ).value = row[46]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TXTC IDs"]
            ).value = row[47]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken PTRN Count"]
            ).value = row[48]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken PTRN IDs"]
            ).value = row[49]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SIMO Count"]
            ).value = row[50]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SIMO IDs"]
            ).value = row[51]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TXTF Count"]
            ).value = row[52]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TXTF IDs"]
            ).value = row[53]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken FAMD Count"]
            ).value = row[54]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken FAMD IDs"]
            ).value = row[55]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken OBJD Count"]
            ).value = row[56]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken OBJD IDs"]
            ).value = row[57]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken OBJK Count"]
            ).value = row[58]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken OBJK IDs"]
            ).value = row[59]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken XML Count"]
            ).value = row[60]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken XML IDs"]
            ).value = row[61]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken STBL Count"]
            ).value = row[62]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken STBL IDs"]
            ).value = row[63]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken NMAP Count"]
            ).value = row[64]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken NMAP IDs"]
            ).value = row[65]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken RSLT Count"]
            ).value = row[66]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken RSLT IDs"]
            ).value = row[67]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken FTPT Count"]
            ).value = row[68]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken FTPT IDs"]
            ).value = row[69]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken RIG Count"]
            ).value = row[70]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken RIG IDs"]
            ).value = row[71]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken JAZZ Count"]
            ).value = row[72]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken JAZZ IDs"]
            ).value = row[73]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken S3SA Count"]
            ).value = row[74]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken S3SA IDs"]
            ).value = row[75]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken LITE Count"]
            ).value = row[76]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken LITE IDs"]
            ).value = row[77]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CLIP Count"]
            ).value = row[78]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CLIP IDs"]
            ).value = row[79]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MATD Count"]
            ).value = row[80]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MATD IDs"]
            ).value = row[81]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MTST Count"]
            ).value = row[82]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MTST IDs"]
            ).value = row[83]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MDLR Count"]
            ).value = row[84]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MDLR IDs"]
            ).value = row[85]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CWAL Count"]
            ).value = row[86]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CWAL IDs"]
            ).value = row[87]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CFEN Count"]
            ).value = row[88]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CFEN IDs"]
            ).value = row[89]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CTPT Count"]
            ).value = row[90]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CTPT IDs"]
            ).value = row[91]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken NGMP Count"]
            ).value = row[92]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken NGMP IDs"]
            ).value = row[93]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CRMT Count"]
            ).value = row[94]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CRMT IDs"]
            ).value = row[95]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SPT2 Count"]
            ).value = row[96]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SPT2 IDs"]
            ).value = row[97]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken _SPT Count"]
            ).value = row[98]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken _SPT IDs"]
            ).value = row[99]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CRAL Count"]
            ).value = row[100]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken CRAL IDs"]
            ).value = row[101]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken LAYO Count"]
            ).value = row[102]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken LAYO IDs"]
            ).value = row[103]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken HINF Count"]
            ).value = row[104]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken HINF IDs"]
            ).value = row[105]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 2ARY Count"]
            ).value = row[106]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 2ARY IDs"]
            ).value = row[107]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Complate Count"]
            ).value = row[108]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Complate IDs"]
            ).value = row[109]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 63A33EA7 Count"]
            ).value = row[110]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 63A33EA7 IDs"]
            ).value = row[111]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Package Storage Count"]
            ).value = row[112]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Package Storage IDs"]
            ).value = row[113]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Fireplace Count"]
            ).value = row[114]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken Fireplace IDs"]
            ).value = row[115]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM Count"]
            ).value = row[116]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM IDs"]
            ).value = row[117]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 626F60CE Count"]
            ).value = row[118]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 626F60CE IDs"]
            ).value = row[119]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken ICON Count"]
            ).value = row[120]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken ICON IDs"]
            ).value = row[121]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B5 Count"]
            ).value = row[122]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B5 IDs"]
            ).value = row[123]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B4 Count"]
            ).value = row[124]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B4 IDs"]
            ).value = row[125]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B6 Count"]
            ).value = row[126]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken THUM 0580A2B6 IDs"]
            ).value = row[127]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SNAP 0580A2CD Count"]
            ).value = row[128]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SNAP 0580A2CD IDs"]
            ).value = row[129]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SNAP 0580A2CF Count"]
            ).value = row[130]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken SNAP 0580A2CF IDs"]
            ).value = row[131]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TWNI Count"]
            ).value = row[132]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken TWNI IDs"]
            ).value = row[133]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MANIFEST Count"]
            ).value = row[134]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken MANIFEST IDs"]
            ).value = row[135]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 8F78F74B Count"]
            ).value = row[136]

            packages_ws.cell(
                package_row,
                packages_column_map["Broken 8F78F74B IDs"]
            ).value = row[137]

            packages_ws.cell(
                package_row,
                packages_column_map["Health Issues"]
            ).value = row[25]

            packages_ws.cell(
                package_row,
                packages_column_map["CC Type"]
            ).value = row[6]

            packages_ws.cell(
                package_row,
                packages_column_map["Subtype"]
            ).value = subtype

            packages_ws.cell(
                package_row,
                packages_column_map["Package Format"]
            ).value = package_format

            packages_ws.cell(
                package_row,
                packages_column_map["Original Source Path"]
            ).value = original_source_path

            packages_ws.cell(
                package_row,
                packages_column_map["CASP IDs"]
            ).value = row[1]

            packages_ws.cell(
                package_row,
                packages_column_map["GEOM IDs"]
            ).value = row[2]

            packages_ws.cell(
                package_row,
                packages_column_map["IMG IDs"]
            ).value = row[3]

            packages_ws.cell(
                package_row,
                packages_column_map["VPXY IDs"]
            ).value = row[4]

            packages_ws.cell(
                package_row,
                packages_column_map["BONE Resource Count"]
            ).value = row[20]

            packages_ws.cell(
                package_row,
                packages_column_map["STBL IDs"]
            ).value = row[15]

            packages_ws.cell(
                package_row,
                packages_column_map["BGEO IDs"]
            ).value = row[16]

            packages_ws.cell(
                package_row,
                packages_column_map["FACE IDs"]
            ).value = row[17]

            packages_ws.cell(
                package_row,
                packages_column_map["FBLN IDs"]
            ).value = row[18]

            packages_ws.cell(
                package_row,
                packages_column_map["KEY IDs"]
            ).value = row[19]

            packages_ws.cell(
                package_row,
                packages_column_map["Highest Poly"]
            ).value = row[12]

            packages_ws.cell(
                package_row,
                packages_column_map["Highest Texture"]
            ).value = row[13]
            
            packages_ws.cell(
                package_row,
                packages_column_map["Package Hash"]
            ).value = row[21] if len(row) > 21 else ""

            package_thumbnail = thumbnail_paths_by_hash.get(
                str((row[21] if len(row) > 21 else "") or "").lower(),
                ""
            )
            if package_thumbnail and "Image" in packages_column_map:
                _set_thumbnail_link(
                    packages_ws.cell(package_row, packages_column_map["Image"]),
                    package_thumbnail
                )

            # Add this new Packages row to the live lookup
            # so later files in the same scan can find it.
            existing_packages.setdefault(
                row[0],
                []
            ).append(package_row)
            packages_row_casp_sets[package_row] = incoming_casp_ids.copy()
            packages_row_hash[package_row] = row[21] if len(row) > 21 else ""

        # The same physical package identity can be encountered more than
        # once in one scan (for example, exact copies in different folders).
        # Count the workbook row once, not once per processed input file.
        package_rows_touched.add(package_row)

    packages_ws.freeze_panes = "A2"

    # Expand/restore the filter range after Packages updates.
    if packages_ws.max_row >= 1 and packages_ws.max_column >= 1:
        packages_ws.auto_filter.ref = "A1:{}{}".format(
            get_column_letter(packages_ws.max_column),
            packages_ws.max_row
        )

    # -------------------------
    # LIGHTWEIGHT COSMETIC MAINTENANCE
    # -------------------------
    # Do not walk/re-style every cell in a large workbook. Existing styles and
    # widths are preserved. For auxiliary sheets, inspect only a small sample
    # so new workbooks still get sensible widths without a million-cell pass.

    print("MASTER DB: lightweight formatting...")

    for sheet in wb:
        if sheet in (ws, packages_ws):
            continue

        sample_max_row = min(sheet.max_row, 200)
        for col_idx in range(1, sheet.max_column + 1):
            max_length = 0
            for row_num in range(1, sample_max_row + 1):
                value = sheet.cell(row_num, col_idx).value
                if value not in (None, ""):
                    max_length = max(max_length, len(str(value)))

            column_letter = get_column_letter(col_idx)
            desired_width = min(max_length + 3, 40) if max_length else 10
            current_width = sheet.column_dimensions[column_letter].width or 0
            if desired_width > current_width:
                sheet.column_dimensions[column_letter].width = desired_width

        # Headers only: cheap, predictable left alignment. Existing data-cell
        # formatting is left untouched.
        for cell in sheet[1]:
            cell.alignment = Alignment(horizontal="left")

    # Migrate old scanner-generated .NET/S3PI stack traces already stored in
    # Health Issues. The detailed raw failures remain in their historical logs;
    # exact resource IDs remain in the dedicated Broken <TYPE> IDs columns.
    master_condensed = _condense_existing_health_issue_cells(ws, last_real_master_row)
    packages_condensed = _condense_existing_health_issue_cells(packages_ws, packages_ws.max_row)
    if master_condensed or packages_condensed:
        print(
            f"MASTER DB: condensed verbose Health Issues cells: "
            f"CC Master={master_condensed:,} | Packages={packages_condensed:,}"
        )

    # Final compact/collapsible review layout after all package writes.
    # Existing Master rows are also migrated to the compact Type — Subtype view.
    _sync_master_type_subtype_display(ws, last_real_master_row)
    _sync_master_content_type_display(ws, last_real_master_row)
    master_health_labels = _sync_health_status_display(ws, last_real_master_row)
    packages_health_labels = _sync_health_status_display(packages_ws, packages_ws.max_row)
    if master_health_labels or packages_health_labels:
        print(
            f"MASTER DB: health label migration BROKEN -> Malformed: "
            f"CC Master={master_health_labels:,} | Packages={packages_health_labels:,}"
        )
    _apply_review_sheet_layout(ws, "master")
    _apply_review_sheet_layout(packages_ws, "packages")
    _apply_table_header_divider(ws)
    _apply_table_header_divider(packages_ws)
    _apply_casp_master_compact_layout(casp_ws)

    # Attention-only direct fills for reliable LibreOffice display.
    # Safe/OK/normal deliberately remain uncoloured.
    duplicate_hash_count = _apply_direct_attention_colours(
        ws, packages_ws, last_real_master_row
    )
    print(
        f"MASTER DB: direct attention colours applied; "
        f"exact duplicate SHA256 groups highlighted={duplicate_hash_count:,}"
    )

    if DEBUG:
        print("MASTER DATABASE FINAL REAL ROW COUNT:", last_real_master_row)

        print("LAST DATABASE REAL ROW:")
        print([cell.value for cell in ws[last_real_master_row]])

        print("SAVING MASTER DATABASE:", MASTER_DATABASE)

    print(
        f"MASTER DB: distinct rows updated: Packages={len(package_rows_touched):,} | "
        f"CASP Master={len(casp_rows_touched):,}"
    )
    print(
        f"MASTER DB: saving workbook... "
        f"({time.perf_counter() - master_db_started:.1f}s elapsed before save)"
    )
    save_started = time.perf_counter()
    wb.save(MASTER_DATABASE)
    print(
        f"MASTER DB: save complete in {time.perf_counter() - save_started:.1f}s; "
        f"total update {time.perf_counter() - master_db_started:.1f}s"
    )
        
    

def check_vpxy_health(package, entry, resource_objects):
    """Parse one VPXY resource. A parse exception is a structural health failure."""
    vpxy_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        try:
            byte_count = len(resource.AsBytes)
        except Exception:
            try:
                byte_count = int(entry.Filesize)
            except Exception:
                byte_count = 0

        print(
            f"VPXY CHECK: {vpxy_id} | Group={group_id} | "
            f"Bytes={byte_count:,} | Wrapper={type(resource).__name__} | Status=OK"
        )

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"VPXY HEALTH FAILURE: {vpxy_id} | {reason}")
        return False, reason


def _validate_face_legacy_tgi_size_40_34(raw):
    """
    Narrow compatibility validator for the observed legacy FACE variant.

    The affected resources are version 8 FACE files whose header reports a
    TGI size of 0x40, while the actual trailing TGI block is 0x34 bytes:
    one UInt32 count plus exactly three 16-byte TGI entries.  The difference
    is exactly 12 bytes, matching the three UInt32 FACE header fields.

    This fallback is intentionally restricted to that exact observed class.
    """
    try:
        if raw is None or len(raw) < 16:
            return False, "FACE resource too short for legacy TGI-size validation", None

        version, tgi_offset, declared_tgi_size = struct.unpack_from("<III", raw, 0)

        if version != 8:
            return False, f"FACE legacy variant version is {version}, expected 8", None

        if declared_tgi_size != 0x40:
            return False, (
                f"FACE legacy declared TGI size is 0x{declared_tgi_size:08X}, "
                "expected 0x00000040"
            ), None

        # In this FACE layout the offset is relative to byte 8.
        tgi_pos = 8 + tgi_offset

        if tgi_pos < 12 or tgi_pos + 4 > len(raw):
            return False, "FACE legacy TGI start lies outside the resource", None

        tgi_count = struct.unpack_from("<I", raw, tgi_pos)[0]

        if tgi_count != 3:
            return False, f"FACE legacy TGI count is {tgi_count}, expected 3", None

        actual_tgi_size = 4 + (tgi_count * 16)

        if actual_tgi_size != 0x34:
            return False, (
                f"FACE legacy actual TGI size is 0x{actual_tgi_size:08X}, "
                "expected 0x00000034"
            ), None

        if declared_tgi_size != actual_tgi_size + 12:
            return False, "FACE legacy TGI-size convention is internally inconsistent", None

        required_end = tgi_pos + actual_tgi_size

        if required_end != len(raw):
            return False, (
                f"FACE legacy TGI block does not end at EOF: "
                f"required={required_end}, resource={len(raw)}"
            ), None

        # Parse every TGI entry to ensure all 16-byte entries are physically present.
        entries = []
        pos = tgi_pos + 4

        for _ in range(tgi_count):
            resource_type, resource_group = struct.unpack_from("<II", raw, pos)
            instance = struct.unpack_from("<Q", raw, pos + 8)[0]
            entries.append((resource_type, resource_group, instance))
            pos += 16

        if pos != len(raw):
            return False, "FACE legacy TGI parser did not finish exactly at EOF", None

        details = {
            "version": version,
            "tgi_count": tgi_count,
            "declared_tgi_size": declared_tgi_size,
            "actual_tgi_size": actual_tgi_size,
            "byte_count": len(raw),
        }

        return True, "", details

    except Exception as e:
        return False, f"{type(e).__name__}: {e}", None


def _validate_bond_extended_key_variant(raw):
    """
    Narrow compatibility validator for the observed complete BOND variant
    that contains private object keys (objectKeyCount > publicKeyCount).

    It validates the complete key table, ObjectData table, version/count,
    every 44-byte BoneAdjust record, all ObjectData ranges, and exact EOF.
    """
    try:
        if raw is None or len(raw) < 28:
            return False, "BOND resource too short for extended-key validation", None

        (
            context_version,
            public_key_count,
            external_key_count,
            delay_load_key_count,
            object_key_count,
        ) = struct.unpack_from("<IIIII", raw, 0)

        if context_version != 3:
            return False, (
                f"BOND context version is {context_version}, expected 3"
            ), None

        counts = (
            public_key_count,
            external_key_count,
            delay_load_key_count,
            object_key_count,
        )

        if any(c > 4096 for c in counts):
            return False, "BOND key count is implausibly large", None

        if object_key_count <= public_key_count:
            return False, (
                "BOND is not the observed extended/private-key variant"
            ), None

        private_key_count = object_key_count - public_key_count

        # objectKeyCount includes public + private object keys.
        total_tgi_records = (
            object_key_count +
            external_key_count +
            delay_load_key_count
        )

        pos = 20
        tgi_bytes = total_tgi_records * 16

        if pos + tgi_bytes > len(raw):
            return False, "BOND TGI key table is truncated", None

        # Parse each full 16-byte TGI.
        for _ in range(total_tgi_records):
            struct.unpack_from("<IIQ", raw, pos)
            pos += 16

        object_data_bytes = object_key_count * 8

        if pos + object_data_bytes + 8 > len(raw):
            return False, "BOND ObjectData/header region is truncated", None

        object_ranges = []

        for _ in range(object_key_count):
            object_pos, object_len = struct.unpack_from("<II", raw, pos)
            pos += 8

            object_end = object_pos + object_len

            if object_pos > len(raw) or object_end > len(raw):
                return False, (
                    f"BOND ObjectData range exceeds EOF: "
                    f"position={object_pos}, length={object_len}, "
                    f"resource={len(raw)}"
                ), None

            object_ranges.append((object_pos, object_len, object_end))

        version, bone_adjust_count = struct.unpack_from("<II", raw, pos)
        pos += 8

        if version != 1:
            return False, f"BOND data version is {version}, expected 1", None

        if bone_adjust_count > 4096:
            return False, "BOND BoneAdjust count is implausibly large", None

        required_end = pos + (bone_adjust_count * 44)

        if required_end != len(raw):
            return False, (
                f"BOND BoneAdjust data does not end at EOF: "
                f"required={required_end}, resource={len(raw)}"
            ), None

        # Physically parse every complete 44-byte BoneAdjust:
        # UInt32 slot hash + ten Single values.
        for _ in range(bone_adjust_count):
            struct.unpack_from("<I10f", raw, pos)
            pos += 44

        if pos != len(raw):
            return False, "BOND extended-key parser did not finish exactly at EOF", None

        # The real sample has at least one ObjectData range ending exactly at EOF.
        # Require this as another narrow safeguard for this compatibility class.
        if not any(end == len(raw) for _, _, end in object_ranges):
            return False, "BOND extended-key variant has no ObjectData range ending at EOF", None

        details = {
            "context_version": context_version,
            "public_key_count": public_key_count,
            "private_key_count": private_key_count,
            "external_key_count": external_key_count,
            "delay_load_key_count": delay_load_key_count,
            "object_key_count": object_key_count,
            "bone_adjust_count": bone_adjust_count,
            "byte_count": len(raw),
        }

        return True, "", details

    except Exception as e:
        return False, f"{type(e).__name__}: {e}", None


def check_bone_bond_health(package, entry, resource_objects, resource_label):
    """Parse one BONE/BOND resource, with one narrow BOND compatibility fallback."""
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        try:
            byte_count = len(resource.AsBytes)
        except Exception:
            try:
                byte_count = int(entry.Filesize)
            except Exception:
                byte_count = 0

        print(
            f"{resource_label} CHECK: {resource_id} | "
            f"Group={group_id} | Bytes={byte_count:,} | "
            f"Wrapper={type(resource).__name__} | Status=OK"
        )
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"

        # TEST COMPATIBILITY PATH:
        # Only BOND EndOfStream failures are eligible, and only if the raw
        # bytes satisfy the complete extended/private-key structure above.
        if resource_label == "BOND" and "EndOfStreamException" in reason:
            try:
                raw = _get_raw_resource_bytes(package, entry)
                compat_ok, compat_error, details = _validate_bond_extended_key_variant(raw)

                if compat_ok:
                    print(
                        f"BOND EXTENDED KEY CHECK: {resource_id} | Group={group_id} | "
                        f"Bytes={details['byte_count']:,} | "
                        f"Context={details['context_version']} | "
                        f"Public={details['public_key_count']} | "
                        f"Private={details['private_key_count']} | "
                        f"Objects={details['object_key_count']} | "
                        f"Adjustments={details['bone_adjust_count']} | Status=OK"
                    )
                    return True, ""

                reason = f"{reason} | BOND compatibility fallback rejected: {compat_error}"

            except Exception as fallback_e:
                reason = (
                    f"{reason} | BOND compatibility fallback error: "
                    f"{type(fallback_e).__name__}: {fallback_e}"
                )

        print(f"{resource_label} HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_morph_health(package, entry, resource_objects, resource_label):
    """Parse one morph resource, with one narrow legacy FACE compatibility fallback."""
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        try:
            byte_count = len(resource.AsBytes)
        except Exception:
            try:
                byte_count = int(entry.Filesize)
            except Exception:
                byte_count = 0

        if resource_label in ("BGEO", "FACE", "FBLN", "BBLN"):
            print(
                f"{resource_label} CHECK: {resource_id} | Group={group_id} | "
                f"Bytes={byte_count:,} | Wrapper={type(resource).__name__} | Status=OK"
            )

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"

        # TEST COMPATIBILITY PATH:
        # Only the exact observed 0x40-vs-0x34 FACE TGI-size exception is
        # eligible, and raw bytes must satisfy the complete narrow variant.
        if (
            resource_label == "FACE"
            and "Size of TGIBlock read: 0x00000040, actual: 0x00000034" in reason
        ):
            try:
                raw = _get_raw_resource_bytes(package, entry)
                compat_ok, compat_error, details = _validate_face_legacy_tgi_size_40_34(raw)

                if compat_ok:
                    print(
                        f"FACE LEGACY TGI-SIZE CHECK: {resource_id} | Group={group_id} | "
                        f"Bytes={details['byte_count']:,} | "
                        f"Version={details['version']} | TGIs={details['tgi_count']} | "
                        f"DeclaredTGI=0x{details['declared_tgi_size']:X} | "
                        f"ActualTGI=0x{details['actual_tgi_size']:X} | Status=OK"
                    )
                    return True, ""

                reason = f"{reason} | FACE compatibility fallback rejected: {compat_error}"

            except Exception as fallback_e:
                reason = (
                    f"{reason} | FACE compatibility fallback error: "
                    f"{type(fallback_e).__name__}: {fallback_e}"
                )

        print(f"{resource_label} HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def diagnose_bbln_resource(package, entry, resource_objects):
    """BBLN remains diagnostic-only until we naturally test a real BBLN package."""
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        try:
            byte_count = len(resource.AsBytes)
        except Exception:
            try:
                byte_count = int(entry.Filesize)
            except Exception:
                byte_count = 0

        print(
            f"BBLN CHECK: {resource_id} | Bytes={byte_count:,} | "
            f"Wrapper={type(resource).__name__} | Status=OK"
        )
        return True

    except Exception as e:
        print(
            f"BBLN CHECK: {resource_id} | Status=PARSE FAILED | "
            f"Error={type(e).__name__}: {e}"
        )
        return False


def check_preset_health(package, entry, resource_objects):
    """Conservative raw-structure health check for PRESET/CBLN resources."""
    resource_id = f"0x{int(entry.Instance):016X}"
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)
        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        if byte_count < 16:
            reason = f"PRESET/CBLN truncated: {byte_count} bytes (too short for minimum structure)"
            print(f"PRESET HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if raw[:8] == b"\\x00" * 8:
            reason = "PRESET/CBLN invalid zero header"
            print(f"PRESET HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        version = int.from_bytes(raw[0:4], "little", signed=False)
        name_length = int.from_bytes(raw[4:8], "little", signed=False)

        if name_length == 0 or name_length > 1024:
            reason = f"PRESET/CBLN invalid name length: {name_length} (version={version})"
            print(f"PRESET HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        name_end = 8 + (name_length * 2)
        if name_end > byte_count:
            reason = f"PRESET/CBLN truncated name: declares {name_length} UTF-16 characters but resource is only {byte_count} bytes"
            print(f"PRESET HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        try:
            preset_name = raw[8:name_end].decode("utf-16le")
        except UnicodeDecodeError as e:
            reason = f"PRESET/CBLN invalid UTF-16 name: {e}"
            print(f"PRESET HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if not preset_name.strip("\\x00").strip():
            reason = "PRESET/CBLN has empty preset name"
            print(f"PRESET HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        try:
            group_id = f"0x{int(entry.ResourceGroup):08X}"
        except Exception:
            group_id = "UNKNOWN"

        print(
            f"PRESET CHECK: {resource_id} | Group={group_id} | "
            f"Bytes={byte_count:,} | Wrapper={type(resource).__name__} | "
            f"Status=OK"
        )

        return True, ""
    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"PRESET HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason

def _validate_jpeg_structure(raw):
    """
    Conservative dependency-free JPEG structural validation used only for
    compatibility when a resource is stored under IMAG type 0x2F7D0004.

    Requires:
      - SOI marker
      - valid marker/segment boundaries
      - a valid SOF with positive dimensions
      - at least one SOS scan
      - a final EOI marker with no trailing bytes

    Handles stuffed 0xFF00 bytes, restart markers, and multiple/progressive scans.
    """
    byte_count = len(raw)

    if byte_count < 4:
        return False, 0, 0, f"JPEG too short: {byte_count} bytes"

    if not raw.startswith(b"\xFF\xD8"):
        return False, 0, 0, "JPEG missing SOI marker"

    sof_markers = set(
        list(range(0xC0, 0xC4))
        + list(range(0xC5, 0xC8))
        + list(range(0xC9, 0xCC))
        + list(range(0xCD, 0xD0))
    )

    pos = 2
    saw_sof = False
    saw_sos = False
    width = 0
    height = 0

    while pos < byte_count:
        if raw[pos] != 0xFF:
            return (
                False,
                width,
                height,
                f"JPEG expected marker at offset {pos}",
            )

        while pos < byte_count and raw[pos] == 0xFF:
            pos += 1

        if pos >= byte_count:
            return False, width, height, "JPEG truncated marker"

        marker = raw[pos]
        pos += 1

        if marker == 0x00:
            return False, width, height, "JPEG stuffed byte outside scan data"

        # End Of Image.
        if marker == 0xD9:
            if not saw_sof:
                return False, width, height, "JPEG has no SOF marker"
            if not saw_sos:
                return False, width, height, "JPEG has no SOS marker"
            if pos != byte_count:
                return (
                    False,
                    width,
                    height,
                    f"JPEG has {byte_count - pos} trailing bytes after EOI",
                )
            return True, width, height, ""

        if marker == 0xD8:
            return False, width, height, "JPEG has unexpected nested SOI marker"

        # TEM and restart markers do not carry a segment length.
        if marker == 0x01 or 0xD0 <= marker <= 0xD7:
            continue

        if pos + 2 > byte_count:
            return False, width, height, "JPEG truncated segment length"

        segment_length = int.from_bytes(raw[pos:pos + 2], "big", signed=False)

        if segment_length < 2:
            return (
                False,
                width,
                height,
                f"JPEG invalid segment length {segment_length}",
            )

        segment_end = pos + segment_length

        if segment_end > byte_count:
            return False, width, height, "JPEG segment overruns resource"

        if marker in sof_markers:
            if segment_length < 7:
                return False, width, height, "JPEG SOF segment too short"

            height = int.from_bytes(raw[pos + 3:pos + 5], "big", signed=False)
            width = int.from_bytes(raw[pos + 5:pos + 7], "big", signed=False)

            if width <= 0 or height <= 0:
                return (
                    False,
                    width,
                    height,
                    f"JPEG invalid dimensions: {width}x{height}",
                )

            saw_sof = True

        if marker == 0xDA:
            saw_sos = True
            pos = segment_end

            # Entropy-coded scan data continues until an unstuffed marker.
            while pos < byte_count:
                if raw[pos] != 0xFF:
                    pos += 1
                    continue

                marker_pos = pos
                q = pos + 1

                while q < byte_count and raw[q] == 0xFF:
                    q += 1

                if q >= byte_count:
                    return False, width, height, "JPEG truncated entropy marker"

                scan_marker = raw[q]

                # 0xFF00 is a stuffed literal 0xFF byte.
                if scan_marker == 0x00:
                    pos = q + 1
                    continue

                # Restart markers are part of entropy-coded scan data.
                if 0xD0 <= scan_marker <= 0xD7:
                    pos = q + 1
                    continue

                # Re-enter the normal marker parser at this marker prefix.
                pos = marker_pos
                break

            continue

        pos = segment_end

    return False, width, height, "JPEG missing EOI marker"


def check_imag_resource_health(package, entry, resource_objects):
    """
    IMAG health handling.

    PNG IMAG (0x2F7D0004) is permanently validated because it has been tested
    with a real known-good package and a deliberately truncated copy.

    JPEG/TGA IMAG types remain diagnostic-only until real examples are tested.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    resource_type = int(entry.ResourceType)
    expected_format = IMAG_TYPES.get(resource_type, "UNKNOWN")

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        # PNG IMAG - PERMANENT HEALTH CHECK
        if expected_format == "PNG":

            png_sig = b"\x89PNG\r\n\x1a\n"

            if byte_count < 24:
                reason = (
                    f"PNG IMAG truncated: {byte_count} bytes "
                    f"(too short for PNG structure)"
                )
                print(f"IMAG HEALTH FAILURE: {resource_id} | {reason}")
                return False, reason

            if not raw.startswith(png_sig):
                # Compatibility: some known-good Sims 3 UI mods store a complete
                # JPEG payload under the nominal PNG IMAG type 0x2F7D0004.
                # Accept only when the JPEG itself passes structural validation.
                if raw.startswith(b"\xFF\xD8"):
                    jpeg_ok, jpeg_width, jpeg_height, jpeg_reason = (
                        _validate_jpeg_structure(raw)
                    )

                    if jpeg_ok:
                        print(
                            f"IMAG JPEG-IN-PNG COMPATIBILITY: {resource_id} | "
                            f"Type=0x{resource_type:08X} | "
                            f"Bytes={byte_count:,} | "
                            f"Size={jpeg_width}x{jpeg_height} | Status=OK"
                        )
                        return True, ""

                    reason = (
                        "PNG IMAG contains malformed JPEG compatibility payload: "
                        f"{jpeg_reason}"
                    )
                    print(f"IMAG HEALTH FAILURE: {resource_id} | {reason}")
                    return False, reason

                reason = (
                    "PNG IMAG signature mismatch: "
                    f"{raw[:16].hex().upper()}"
                )
                print(f"IMAG HEALTH FAILURE: {resource_id} | {reason}")
                return False, reason

            width = int.from_bytes(raw[16:20], "big", signed=False)
            height = int.from_bytes(raw[20:24], "big", signed=False)

            if width <= 0 or height <= 0:
                reason = f"PNG IMAG invalid dimensions: {width}x{height}"
                print(f"IMAG HEALTH FAILURE: {resource_id} | {reason}")
                return False, reason

            iend_ok = (
                byte_count >= 12
                and raw[-12:-8] == b"\x00\x00\x00\x00"
                and raw[-8:-4] == b"IEND"
            )

            if not iend_ok:
                reason = "PNG IMAG missing/invalid IEND terminator"
                print(f"IMAG HEALTH FAILURE: {resource_id} | {reason}")
                return False, reason

            return True, ""

        # JPEG IMAG - DIAGNOSTIC ONLY
        if expected_format == "JPEG":

            if byte_count < 4:
                status = "TRUNCATED CANDIDATE"
                details = f"Header={raw[:16].hex().upper()}"
            elif not raw.startswith(b"\xFF\xD8"):
                status = "SIGNATURE MISMATCH"
                details = f"Header={raw[:16].hex().upper()}"
            else:
                end_ok = raw.endswith(b"\xFF\xD9")
                status = "OK" if end_ok else "TRUNCATED CANDIDATE"
                details = f"SOI=OK | EOI={'OK' if end_ok else 'MISSING'}"

            print(
                f"IMAG CHECK: {resource_id} | "
                f"Type=0x{resource_type:08X} | Format=JPEG | "
                f"Bytes={byte_count:,} | {details} | "
                f"Wrapper={type(resource).__name__} | Status={status}"
            )
            return True, ""

        # TGA IMAG - DIAGNOSTIC ONLY
        if expected_format == "TGA":

            if byte_count < 18:
                status = "TRUNCATED CANDIDATE"
                details = f"Header={raw[:16].hex().upper()}"
            else:
                image_type = raw[2]
                width = int.from_bytes(raw[12:14], "little", signed=False)
                height = int.from_bytes(raw[14:16], "little", signed=False)
                pixel_depth = raw[16]

                plausible_type = image_type in (0, 1, 2, 3, 9, 10, 11)
                plausible_depth = pixel_depth in (0, 8, 15, 16, 24, 32)

                status = (
                    "OK"
                    if (
                        width > 0
                        and height > 0
                        and plausible_type
                        and plausible_depth
                    )
                    else "STRUCTURE SUSPICIOUS"
                )

                details = (
                    f"Size={width}x{height} | "
                    f"ImageType={image_type} | PixelDepth={pixel_depth}"
                )

            print(
                f"IMAG CHECK: {resource_id} | "
                f"Type=0x{resource_type:08X} | Format=TGA | "
                f"Bytes={byte_count:,} | {details} | "
                f"Wrapper={type(resource).__name__} | Status={status}"
            )
            return True, ""

        print(
            f"IMAG CHECK: {resource_id} | "
            f"Type=0x{resource_type:08X} | "
            f"Format={expected_format} | Status=UNKNOWN FORMAT"
        )
        return True, ""

    except Exception as e:

        if expected_format == "PNG":
            reason = f"{type(e).__name__}: {e}"
            print(f"IMAG HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        print(
            f"IMAG CHECK: {resource_id} | "
            f"Type=0x{resource_type:08X} | "
            f"Format={expected_format} | Status=READ/PARSE FAILED | "
            f"Error={type(e).__name__}: {e}"
        )
        return True, ""


def check_ptrn_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 PTRN pattern-list resources.

    PTRN is validated as text/XML. A resource is considered broken when it
    cannot be read, cannot be decoded as plausible text, or is malformed XML.

    We do not require an exact byte size or a fixed child count.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        if byte_count == 0:
            reason = "PTRN empty resource"
            print(f"PTRN HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append(("utf-8-sig", "UTF-8 BOM"))
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append(("utf-16", "UTF-16 BOM"))

        decode_attempts.extend([
            ("utf-8", "UTF-8"),
            ("utf-16-le", "UTF-16LE"),
        ])

        decoded = None
        last_decode_error = None

        for codec, _label in decode_attempts:
            try:
                candidate = raw.decode(codec)
                visible = candidate.replace("\x00", "")

                if not visible.strip():
                    continue

                if candidate.count("\x00") > max(4, len(candidate) // 4):
                    if codec != "utf-16-le":
                        continue

                decoded = candidate
                break

            except Exception as e:
                last_decode_error = e

        if decoded is None:
            reason = (
                "PTRN text decode failed"
                + (
                    f": {type(last_decode_error).__name__}: {last_decode_error}"
                    if last_decode_error else ""
                )
            )
            print(f"PTRN HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        cleaned = decoded.strip("\x00 \t\r\n")

        try:
            ET.fromstring(cleaned)
        except ET.ParseError as e:
            reason = f"PTRN XML parse failed: {e}"
            print(f"PTRN HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"PTRN HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_famd_health(package, entry, resource_objects):
    """
    Permanent conservative structural health check for Sims 3 FAMD /
    Family Data resources.

    The installed S3PI build exposes FAMD as generic IResource, so this
    check validates only unmistakable raw structural damage and avoids
    assuming one exact valid version or byte size.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        if byte_count < 16:
            reason = f"FAMD truncated: {byte_count} bytes"
            print(f"FAMD HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if raw[:12] == b"\x00" * 12:
            reason = "FAMD invalid zero header"
            print(f"FAMD HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"FAMD HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_objd_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 OBJD / Object Definition.

    OBJD has a real S3PI catalog parser. Any exception while deserializing the
    resource is treated as structural corruption.

    No fixed OBJD version or byte-size requirement is imposed.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"OBJD HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_objk_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 OBJK / Object Key.

    OBJK has a real S3PI parser. Any exception while deserializing the
    resource is treated as structural corruption.

    No fixed OBJK version or byte-size requirement is imposed.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"OBJK HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason



def _decode_resource_text_views(raw):
    """Classification-only text views; failures never affect health status."""
    decoded_views = []
    for codec in ("utf-8", "utf-16-le", "utf-16-be"):
        try:
            text = raw.decode(codec, errors="ignore")
            if text:
                decoded_views.append(text.replace("\x00", ""))
        except Exception:
            pass
    return decoded_views


def detect_objk_script_class(package, entry):
    """Return the exact scriptClass value stored in a healthy OBJK, if readable."""
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        # OBJK stores string properties as:
        #   property-name NUL + uint32 little-endian byte length + UTF-8 bytes
        # Reading that structure directly avoids regex artefacts where the
        # length byte can be mistaken for part of the class name.
        marker = b"scriptClass\x00"
        marker_pos = raw.find(marker)
        if marker_pos >= 0:
            length_pos = marker_pos + len(marker)
            if length_pos + 4 <= len(raw):
                value_length = struct.unpack_from("<I", raw, length_pos)[0]
                value_start = length_pos + 4
                value_end = value_start + value_length
                if 0 < value_length <= 1024 and value_end <= len(raw):
                    class_name = raw[value_start:value_end].decode(
                        "utf-8", errors="ignore"
                    ).strip("\x00 \t\r\n")
                    if "." in class_name:
                        return True, class_name

        # Conservative fallback for uncommon OBJK encodings.
        class_pattern = re.compile(
            r"(?:[A-Za-z_][A-Za-z0-9_+`]*\.)+[A-Za-z_][A-Za-z0-9_+`]*"
        )
        for text in _decode_resource_text_views(raw):
            marker_pos = text.lower().find("scriptclass")
            if marker_pos < 0:
                continue
            window = text[marker_pos:marker_pos + 768]
            # Prefer a Sims3-rooted class because it cannot absorb an OBJK
            # string-length byte into a fake identifier prefix.
            sims_match = re.search(
                r"Sims3(?:\.[A-Za-z_][A-Za-z0-9_+`]*)+",
                window
            )
            if sims_match:
                return True, sims_match.group(0)
            match = class_pattern.search(window)
            if match and not match.group(0).lower().startswith("scriptclass"):
                return True, match.group(0)
        return False, ""
    except Exception:
        return False, ""


def detect_custom_objk_script_class(package, entry):
    """Conservative evidence for an object whose custom script lives elsewhere."""
    found, class_name = detect_objk_script_class(package, entry)
    if not found or not class_name.startswith("Sims3.Gameplay.Objects."):
        return False, ""

    # External-script packages have no local S3SA to prove ownership of the
    # class.  Require a creator-like lowercase namespace segment BEFORE the
    # final class name.  The leaf itself is deliberately ignored: EA classes
    # such as ...Lighting.Mimics.lightingStreetlightWooden use lowercase leaf
    # names and must remain ordinary Build/Buy.
    parts = class_name.split(".")
    namespace_parts = parts[3:-1]  # after Sims3.Gameplay.Objects, before leaf
    if any(part and part[0].islower() for part in namespace_parts):
        return True, class_name
    return False, ""


def _rva_to_pe_offset(pe_bytes, rva, pe_header_offset, section_count, optional_size):
    section_table = pe_header_offset + 4 + 20 + optional_size
    for index in range(section_count):
        offset = section_table + index * 40
        virtual_size = struct.unpack_from("<I", pe_bytes, offset + 8)[0]
        virtual_address = struct.unpack_from("<I", pe_bytes, offset + 12)[0]
        raw_size = struct.unpack_from("<I", pe_bytes, offset + 16)[0]
        raw_pointer = struct.unpack_from("<I", pe_bytes, offset + 20)[0]
        span = max(virtual_size, raw_size)
        if virtual_address <= rva < virtual_address + span:
            return raw_pointer + (rva - virtual_address)
    if 0 <= rva < len(pe_bytes):
        return rva
    raise ValueError("RVA is outside PE sections")


def _extract_dotnet_typedef_names(raw):
    """Return full TypeDef names from the .NET assembly embedded in S3SA bytes."""
    try:
        mz_offset = raw.find(b"MZ")
        if mz_offset < 0:
            return set()
        pe_bytes = raw[mz_offset:]
        if len(pe_bytes) < 0x40:
            return set()

        pe_header_offset = struct.unpack_from("<I", pe_bytes, 0x3C)[0]
        if pe_bytes[pe_header_offset:pe_header_offset + 4] != b"PE\x00\x00":
            return set()

        coff = pe_header_offset + 4
        section_count = struct.unpack_from("<H", pe_bytes, coff + 2)[0]
        optional_size = struct.unpack_from("<H", pe_bytes, coff + 16)[0]
        optional = coff + 20
        optional_magic = struct.unpack_from("<H", pe_bytes, optional)[0]
        if optional_magic == 0x10B:
            data_directories = optional + 96
        elif optional_magic == 0x20B:
            data_directories = optional + 112
        else:
            return set()

        cli_rva = struct.unpack_from(
            "<I", pe_bytes, data_directories + (14 * 8)
        )[0]
        if not cli_rva:
            return set()
        cli_offset = _rva_to_pe_offset(
            pe_bytes, cli_rva, pe_header_offset, section_count, optional_size
        )
        metadata_rva = struct.unpack_from("<I", pe_bytes, cli_offset + 8)[0]
        metadata = _rva_to_pe_offset(
            pe_bytes, metadata_rva, pe_header_offset, section_count, optional_size
        )
        if pe_bytes[metadata:metadata + 4] != b"BSJB":
            return set()

        version_length = struct.unpack_from("<I", pe_bytes, metadata + 12)[0]
        cursor = metadata + 16 + version_length
        cursor = (cursor + 3) & ~3
        stream_count = struct.unpack_from("<H", pe_bytes, cursor + 2)[0]
        cursor += 4
        streams = {}
        for _ in range(stream_count):
            stream_offset = struct.unpack_from("<I", pe_bytes, cursor)[0]
            stream_size = struct.unpack_from("<I", pe_bytes, cursor + 4)[0]
            name_start = cursor + 8
            name_end = pe_bytes.index(b"\x00", name_start)
            stream_name = pe_bytes[name_start:name_end].decode("ascii", errors="ignore")
            cursor = (name_end + 1 + 3) & ~3
            streams[stream_name] = (metadata + stream_offset, stream_size)

        tables_stream = streams.get("#~") or streams.get("#-")
        strings_stream = streams.get("#Strings")
        if not tables_stream or not strings_stream:
            return set()

        table_offset = tables_stream[0]
        strings_offset, strings_size = strings_stream
        cursor = table_offset + 4
        heap_sizes = pe_bytes[cursor + 2]
        cursor += 4
        valid_mask = struct.unpack_from("<Q", pe_bytes, cursor)[0]
        cursor += 16  # valid + sorted masks

        row_counts = [0] * 64
        for table_index in range(64):
            if (valid_mask >> table_index) & 1:
                row_counts[table_index] = struct.unpack_from(
                    "<I", pe_bytes, cursor
                )[0]
                cursor += 4

        string_index_size = 4 if (heap_sizes & 0x01) else 2
        guid_index_size = 4 if (heap_sizes & 0x02) else 2

        def table_index_size(table_index):
            return 4 if row_counts[table_index] >= 65536 else 2

        def coded_index_size(target_tables, tag_bits):
            max_rows = max(row_counts[index] for index in target_tables)
            return 4 if max_rows >= (1 << (16 - tag_bits)) else 2

        resolution_scope_size = coded_index_size([0, 26, 35, 1], 2)
        typedef_or_ref_size = coded_index_size([2, 1, 27], 2)

        module_row_size = 2 + string_index_size + (guid_index_size * 3)
        typeref_row_size = resolution_scope_size + (string_index_size * 2)
        typedef_row_size = (
            4
            + (string_index_size * 2)
            + typedef_or_ref_size
            + table_index_size(4)
            + table_index_size(6)
        )

        typedef_start = cursor
        if row_counts[0]:
            typedef_start += row_counts[0] * module_row_size
        if row_counts[1]:
            typedef_start += row_counts[1] * typeref_row_size

        def read_index(offset, size):
            if size == 2:
                return struct.unpack_from("<H", pe_bytes, offset)[0]
            return struct.unpack_from("<I", pe_bytes, offset)[0]

        def read_string(index):
            if index <= 0 or index >= strings_size:
                return ""
            start = strings_offset + index
            end = pe_bytes.find(b"\x00", start, strings_offset + strings_size)
            if end < 0:
                end = strings_offset + strings_size
            return pe_bytes[start:end].decode("utf-8", errors="ignore")

        defined_types = set()
        for row in range(row_counts[2]):
            offset = typedef_start + row * typedef_row_size + 4
            name_index = read_index(offset, string_index_size)
            offset += string_index_size
            namespace_index = read_index(offset, string_index_size)
            type_name = read_string(name_index)
            namespace = read_string(namespace_index)
            if not type_name or type_name == "<Module>":
                continue
            full_name = (namespace + "." + type_name) if namespace else type_name
            defined_types.add(full_name)
        return defined_types
    except Exception:
        # Classification metadata is best-effort and must never affect health.
        return set()


def extract_s3sa_defined_types(package, entry):
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        return _extract_dotnet_typedef_names(bytes(resource.AsBytes))
    except Exception:
        return set()


def check_xml_health(package, entry, resource_objects):
    """
    Structural health check for Sims 3 _XML resources (0x0333406C).

    Important compatibility rule:
    a byte sequence is not accepted merely because one codec can decode it.
    Each plausible decoding must also parse as a complete XML document. This
    avoids false BROKEN results when the first decodable codec is not the real
    encoding.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        if byte_count == 0:
            reason = "XML empty resource"
            print(f"XML HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        parse_errors = []
        decode_errors = []

        # ElementTree can natively honour BOMs and XML encoding declarations.
        # Try that first because it is the least assumption-heavy route.
        try:
            ET.fromstring(raw)
            return True, ""
        except ET.ParseError as e:
            parse_errors.append(("auto", e))
        except Exception as e:
            decode_errors.append(e)

        decode_attempts = []
        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append(("utf-8-sig", "UTF-8 BOM"))
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append(("utf-16", "UTF-16 BOM"))

        decode_attempts.extend([
            ("utf-8", "UTF-8"),
            ("utf-16-le", "UTF-16LE"),
            ("utf-16-be", "UTF-16BE"),
        ])

        seen_codecs = set()
        successful_decode = False

        for codec, label in decode_attempts:
            if codec in seen_codecs:
                continue
            seen_codecs.add(codec)

            try:
                candidate = raw.decode(codec)
            except Exception as e:
                decode_errors.append(e)
                continue

            # Reject obviously wrong single-byte decoding of UTF-16 data.
            if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                nul_count = candidate.count("\x00")
                if nul_count > max(4, len(candidate) // 4):
                    continue

            cleaned = candidate.strip("\ufeff\x00 \t\r\n")
            if not cleaned or not cleaned.startswith("<"):
                continue

            successful_decode = True
            try:
                ET.fromstring(cleaned)
                return True, ""
            except ET.ParseError as e:
                parse_errors.append((label, e))

        if parse_errors:
            # Prefer the parser's native byte-level interpretation first; it
            # normally produces the most faithful line/column for real damage.
            _label, best_error = parse_errors[0]
            reason = f"XML parse failed: {best_error}"
            print(f"XML HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        reason = "XML text decode failed"
        if decode_errors:
            last_decode_error = decode_errors[-1]
            reason += f": {type(last_decode_error).__name__}: {last_decode_error}"
        elif successful_decode:
            reason = "XML parse failed: no valid XML document found"

        print(f"XML HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"XML HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def detect_career_xml_marker(package, entry):
    """
    Classification-only detector for NRaas-style custom career loader XML.

    This deliberately does NOT participate in XML structural health. The
    permanent XML health checker remains unchanged. A package is marked as a
    career only when an XML resource has a CareerFiles root and contains a
    non-empty Careers entry inside a CareerFile block.
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "CareerFiles":
            return False

        for career_file in root.iter():
            if local_name(career_file.tag) != "CareerFile":
                continue

            for child in list(career_file):
                if (
                    local_name(child.tag) == "Careers"
                    and (child.text or "").strip()
                ):
                    return True

        return False

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False
















def detect_cooking_tuning_xml_marker(package, entry, resource_type):
    """
    Classification-only detector for the multi-table Sims 3 cooking /
    ingredients tuning family observed in the CookingAndIngredients overhaul.

    Returns one marker name for recognized cooking tables, otherwise None.

    Permanent XML health remains authoritative for normal 0x0333406C XML.
    The Store Ingredients resource (0xE5105066) has no separate permanent
    health checker, so this helper requires it to decode and parse completely
    and to match the Ingredients_store schema before it can contribute to
    classification.
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return None

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        root_name = local_name(root.tag)
        direct_children = list(root)

        def field_map(row):
            return {
                local_name(child.tag): (child.text or "").strip()
                for child in list(row)
            }

        # Store Ingredients is its own resource type in Sims 3.
        if resource_type == STORE_INGREDIENTS_TYPE:
            if root_name != "Ingredients_store":
                return None

            rows = [
                child for child in direct_children
                if local_name(child.tag) == "Data"
            ]

            if len(rows) < 10 or len(rows) != len(direct_children):
                return None

            good_rows = 0

            for row in rows:
                fields = field_map(row)

                if {
                    "Ingredient_Key",
                    "Ingredient_Name",
                    "Rarity",
                    "UiTab",
                    "Model_Name",
                    "Price",
                }.issubset(fields):
                    good_rows += 1

            if good_rows >= 10:
                return "store_ingredients"

            return None

        if resource_type != XML_TYPE:
            return None

        # Recipe master list.
        if root_name == "RecipeMasterList":
            rows = [
                child for child in direct_children
                if local_name(child.tag) == "Data"
            ]

            if len(rows) < 100 or len(rows) != len(direct_children):
                return None

            good_rows = 0

            for row in rows:
                fields = field_map(row)

                if {
                    "Recipe_Key",
                    "Recipe_Name",
                    "Recipe_Class",
                    "Level",
                    "Skill_Chunk",
                }.issubset(fields):
                    good_rows += 1

            if good_rows >= 100:
                return "recipe_master"

            return None

        # Main ingredients table.
        if root_name == "Ingredients":
            rows = [
                child for child in direct_children
                if local_name(child.tag) == "Data"
            ]

            if len(rows) < 80 or len(rows) != len(direct_children):
                return None

            good_rows = 0

            for row in rows:
                fields = field_map(row)

                if {
                    "Ingredient_Key",
                    "Ingredient_Name",
                    "Model_Name",
                }.issubset(fields):
                    good_rows += 1

            if good_rows >= 80:
                return "ingredients"

            return None

        # Plants table.
        if root_name == "Plants":
            table_counts = {}

            for child in direct_children:
                name = local_name(child.tag)
                table_counts[name] = table_counts.get(name, 0) + 1

            if (
                table_counts.get("PlantList", 0) >= 50
                and table_counts.get("MedatorInstanceList", 0) >= 10
            ):
                return "plants"

            return None

        # Bars / drinks / venue table.
        if root_name == "Bars":
            table_counts = {}

            for child in direct_children:
                name = local_name(child.tag)
                table_counts[name] = table_counts.get(name, 0) + 1

            if (
                table_counts.get("Drinks", 0) >= 40
                and table_counts.get("Ingredients", 0) >= 20
                and table_counts.get("Venues", 0) >= 100
            ):
                return "bars"

            return None

        # Shopping tabs used by the ingredient/recipe overhaul.
        if root_name == "ShopTabs":
            if len(direct_children) < 20:
                return None

            good_tabs = 0

            for child in direct_children:
                if (
                    child.attrib.get("localizedName")
                    and child.attrib.get("image")
                ):
                    good_tabs += 1

            if good_tabs >= 20:
                return "shop_tabs"

            return None

        # Individual recipe process tables. Their root is the recipe name,
        # so identify them by the repeated internal cooking schema rather
        # than by a filename or hardcoded recipe-name list.
        table_counts = {}

        for child in direct_children:
            name = local_name(child.tag)
            table_counts[name] = table_counts.get(name, 0) + 1

        if (
            table_counts.get("CookingProcess", 0) >= 10
            and table_counts.get("CuttingBoard", 0) >= 1
            and table_counts.get("FoodTray", 0) >= 1
        ):
            good_process_rows = 0

            for child in direct_children:
                if local_name(child.tag) != "CookingProcess":
                    continue

                fields = field_map(child)

                if (
                    "Cooking_Processes" in fields
                    and (
                        "Step_1" in fields
                        or "Supported_Processes" in fields
                    )
                ):
                    good_process_rows += 1

            if good_process_rows >= 10:
                return "recipe_process"

        return None

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return None


def detect_fishing_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the Sims 3 full Fishing tuning table.

    This does NOT replace permanent XML health. The marker targets the large
    GameplayData-style Fishing table and requires all four observed table
    families: Fish, Spawner, BoxItems, and GlobalSpawner.

    STBL resources are handled separately as optional localization auxiliaries.
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "Fishing":
            return False

        direct_children = list(root)

        if not direct_children:
            return False

        allowed_tables = {
            "Fish",
            "Spawner",
            "BoxItems",
            "GlobalSpawner",
        }

        if any(local_name(child.tag) not in allowed_tables for child in direct_children):
            return False

        fish_rows = [
            child for child in direct_children
            if local_name(child.tag) == "Fish"
        ]
        spawner_rows = [
            child for child in direct_children
            if local_name(child.tag) == "Spawner"
        ]
        box_rows = [
            child for child in direct_children
            if local_name(child.tag) == "BoxItems"
        ]
        global_rows = [
            child for child in direct_children
            if local_name(child.tag) == "GlobalSpawner"
        ]

        # Conservative full-table minimums grounded in the inspected package.
        if len(fish_rows) < 40:
            return False

        if len(spawner_rows) < 20:
            return False

        if len(box_rows) < 10:
            return False

        if len(global_rows) < 5:
            return False

        fish_required = {
            "Name",
            "Description",
            "Hex",
            "WaterTypes",
            "Level",
            "MinWeight",
            "MaxWeight",
            "SkillPoints",
            "Ingredient_Key",
            "Model_Name",
            "Rarity",
        }

        spawner_required = {
            "SpawnerClassName",
            "InactiveFish",
            "InactiveProbability",
            "AvgLength",
            "MaxRange",
        }

        box_required = {
            "ItemDbName",
            "Happy",
            "RelativeProbability",
            "ProductVersion",
        }

        global_required = {
            "WaterAffected",
            "Fish",
            "Probability",
            "AvgLength",
            "WorldName",
        }

        def field_map(row):
            return {
                local_name(child.tag): (child.text or "").strip()
                for child in list(row)
            }

        # Every row must retain the core schema for its table family.
        for row in fish_rows:
            if not fish_required.issubset(field_map(row)):
                return False

        for row in spawner_rows:
            if not spawner_required.issubset(field_map(row)):
                return False

        for row in box_rows:
            if not box_required.issubset(field_map(row)):
                return False

        for row in global_rows:
            if not global_required.issubset(field_map(row)):
                return False

        # Require substantial populated evidence, not only template/default rows.
        populated_fish = 0
        populated_spawners = 0
        populated_boxes = 0
        populated_globals = 0
        fish_price_rows = 0

        has_fish_template = False

        for row in fish_rows:
            fields = field_map(row)

            if (
                not fields.get("Name")
                and fields.get("Hex") == "0"
                and fields.get("WaterTypes")
                and fields.get("Level")
            ):
                has_fish_template = True

            if fields.get("Name") and fields.get("Hex"):
                populated_fish += 1

            # Normal fish rows carry MinPrice/MaxPrice, while the legitimate
            # special "Box" fishing-chest row does not. Require this as broad
            # table evidence rather than as a universal per-row field pair.
            if "MinPrice" in fields and "MaxPrice" in fields:
                fish_price_rows += 1

        for row in spawner_rows:
            fields = field_map(row)
            if fields.get("SpawnerClassName"):
                populated_spawners += 1

        for row in box_rows:
            fields = field_map(row)
            if fields.get("ItemDbName"):
                populated_boxes += 1

        for row in global_rows:
            fields = field_map(row)
            if fields.get("Fish") and fields.get("Probability"):
                populated_globals += 1

        return (
            has_fish_template
            and populated_fish >= 30
            and fish_price_rows >= 40
            and populated_spawners >= 10
            and populated_boxes >= 5
            and populated_globals >= 5
        )

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_skill_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the Sims 3 full Skills tuning table.

    This does NOT replace permanent XML health. The v1 marker deliberately
    targets the large GameplayData-style Skills table, not small custom-skill
    packages or unrelated XML resources.

    Required evidence:
    - root is exactly Skills
    - at least 50 direct SkillList rows
    - every SkillList has SkillName, Hex, and MaxSkillLevel fields
    - a broad default/template SkillList row is present
    - at least 50 populated skill rows are present
    - table-wide populated rows expose real skill progression/class/data fields
    - at least three referenced CustomDataSheet names also exist as direct
      skill-specific tables under the Skills root
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "Skills":
            return False

        direct_children = list(root)

        if not direct_children:
            return False

        skill_rows = [
            child
            for child in direct_children
            if local_name(child.tag) == "SkillList"
        ]

        if len(skill_rows) < 50:
            return False

        direct_table_names = {
            local_name(child.tag)
            for child in direct_children
            if local_name(child.tag) != "SkillList"
        }

        template_fields = {
            "SkillName",
            "SkillDescription",
            "Hex",
            "CodeVersion",
            "CustomClassName",
            "CustomDataSheet",
            "ThoughtBalloonTopic",
            "MaxSkillLevel",
            "Commodity",
            "Physical",
            "Mental",
            "Musical",
            "Creative",
            "Artistic",
            "Hidden",
            "HiddenWithSkillProgress",
            "Certificate",
            "CanDecay",
            "LogicSkillBoost",
            "AlwaysDisplayTNS",
            "Level_1",
            "Level_2",
            "Level_3",
            "Level_4",
            "Level_5",
            "Level_6",
            "Level_7",
            "Level_8",
            "Level_9",
            "Level_10",
            "DreamsAndPromisesIcon",
            "IconKey",
            "SkillUIIcon",
        }

        has_full_template = False
        populated_count = 0
        has_custom_class = False
        has_custom_data_sheet = False
        has_commodity = False
        has_available_age_species = False
        has_progression_level = False
        referenced_data_sheets = set()

        for skill in skill_rows:
            fields = {
                local_name(child.tag): (child.text or "").strip()
                for child in list(skill)
            }

            if not {"SkillName", "Hex", "MaxSkillLevel"}.issubset(fields):
                return False

            if (
                template_fields.issubset(set(fields))
                and not fields.get("SkillName")
                and not fields.get("Hex")
            ):
                has_full_template = True

            if fields.get("SkillName") or fields.get("Hex"):
                if not (
                    fields.get("SkillName")
                    and fields.get("Hex")
                    and fields.get("MaxSkillLevel")
                ):
                    return False

                populated_count += 1

                if fields.get("CustomClassName"):
                    has_custom_class = True

                if fields.get("CustomDataSheet"):
                    has_custom_data_sheet = True
                    referenced_data_sheets.add(fields["CustomDataSheet"])

                if fields.get("Commodity"):
                    has_commodity = True

                if fields.get("AvailableAgeSpecies"):
                    has_available_age_species = True

                if any(
                    fields.get(name)
                    for name in (
                        "Level_1",
                        "Level_2",
                        "Level_3",
                        "Level_4",
                        "Level_5",
                        "Level_6",
                        "Level_7",
                        "Level_8",
                        "Level_9",
                        "Level_10",
                    )
                ):
                    has_progression_level = True

        matching_data_sheets = referenced_data_sheets.intersection(
            direct_table_names
        )

        return (
            has_full_template
            and populated_count >= 50
            and has_custom_class
            and has_custom_data_sheet
            and has_commodity
            and has_available_age_species
            and has_progression_level
            and len(matching_data_sheets) >= 3
        )

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_social_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the Sims 3 SocialData_* XML tables.

    This does NOT replace permanent XML health. The v1 marker is deliberately
    conservative and targets full social-data tables rather than tiny custom
    Social XML fragments.

    Required evidence:
    - root is Social
    - at least 100 direct Action records
    - every direct child is Action with non-empty key + com attributes
    - every Action has Text, AD, AA, Jazz, and Rules children
    - every AD has a val attribute
    - table-wide evidence includes AGE + Available under AA,
      Graph under Jazz, and both LHS + RHS under Rules
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "Social":
            return False

        actions = list(root)

        if len(actions) < 100:
            return False

        required_action_children = {
            "Text",
            "AD",
            "AA",
            "Jazz",
            "Rules",
        }

        has_age = False
        has_available = False
        has_graph = False
        has_lhs = False
        has_rhs = False

        for action in actions:
            if local_name(action.tag) != "Action":
                return False

            if not (action.attrib.get("key") or "").strip():
                return False

            if not (action.attrib.get("com") or "").strip():
                return False

            child_map = {}

            for child in list(action):
                child_map.setdefault(local_name(child.tag), []).append(child)

            if not required_action_children.issubset(set(child_map)):
                return False

            ad = child_map["AD"][0]

            if not (ad.attrib.get("val") or "").strip():
                return False

            aa = child_map["AA"][0]
            jazz = child_map["Jazz"][0]
            rules = child_map["Rules"][0]

            aa_names = {
                local_name(child.tag)
                for child in list(aa)
            }
            jazz_names = {
                local_name(child.tag)
                for child in list(jazz)
            }
            rules_names = {
                local_name(child.tag)
                for child in list(rules)
            }

            if "AGE" in aa_names:
                has_age = True

            if "Available" in aa_names:
                has_available = True

            if "Graph" in jazz_names:
                has_graph = True

            if "LHS" in rules_names:
                has_lhs = True

            if "RHS" in rules_names:
                has_rhs = True

        return (
            has_age
            and has_available
            and has_graph
            and has_lhs
            and has_rhs
        )

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_buff_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the Sims 3 full Buffs tuning table.

    This does NOT replace permanent XML health. The marker is deliberately
    strict and based on the observed full-table structure.

    Required evidence:
    - root is exactly buffs
    - every direct child is BuffList
    - one template/default BuffList exposes the broad core field set
    - every BuffList contains the characteristic identity/display fields
    - at least 100 populated buff rows are present
    - populated rows include real mood/effect/duration/class data
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "buffs":
            return False

        direct_children = list(root)

        if not direct_children:
            return False

        if any(
            local_name(child.tag) != "BuffList"
            for child in direct_children
        ):
            return False

        universal_fields = {
            "Hex",
            "BuffName",
            "BuffDescription",
            "ThumbFilename",
            "AxisEffected",
        }

        template_fields = {
            "Hex",
            "SKU",
            "BuffName",
            "BuffDescription",
            "BuffHelpText",
            "ThumbFilename",
            "AxisEffected",
            "EffectValue",
            "DelayTimer",
            "TimeoutLength",
            "PermaMoodlet",
            "PermaMoodletColor",
            "PolarityOverride",
            "SolveCommodity",
            "AttemptAutoSolve",
            "SolveTime",
            "Stackable",
            "IsExtreme",
            "FacialIdle",
            "IncreasedEffectiveness",
            "ReducedEffectiveness",
            "CustomClassName",
            "Topic",
            "ShowBallon",
            "Travel",
            "DisallowedOccults",
            "JazzStateSuffix",
            "SpeciesAvailability",
        }

        has_full_template = False
        populated_count = 0
        has_effect_value = False
        has_timeout = False
        has_custom_class = False

        for buff in direct_children:
            fields = {
                local_name(child.tag): (child.text or "").strip()
                for child in list(buff)
            }

            field_names = set(fields)

            if not universal_fields.issubset(field_names):
                return False

            if (
                template_fields.issubset(field_names)
                and not fields.get("Hex")
                and not fields.get("BuffName")
            ):
                has_full_template = True

            if fields.get("Hex") and fields.get("BuffName"):
                populated_count += 1

                if "EffectValue" in fields:
                    has_effect_value = True

                if "TimeoutLength" in fields:
                    has_timeout = True

                if fields.get("CustomClassName"):
                    has_custom_class = True

        return (
            has_full_template
            and populated_count >= 100
            and has_effect_value
            and has_timeout
            and has_custom_class
        )

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_boarding_school_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the Sims 3 BoardingSchool tuning table.

    This does NOT replace permanent XML health. The marker is deliberately
    strict and based on the observed full-table schema.

    Required evidence:
    - root is BoardingSchool
    - every direct child is BoardingSchool
    - one default/template row exposes the full core configuration fields
    - at least five populated school rows are present
    - populated rows contain characteristic school/skill/career fields
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "BoardingSchool":
            return False

        direct_children = list(root)

        if not direct_children:
            return False

        if any(
            local_name(child.tag) != "BoardingSchool"
            for child in direct_children
        ):
            return False

        template_fields = {
            "GUID",
            "Name",
            "Cost",
            "SchoolTraits",
            "OffensiveTraits",
            "MaxTraitsReplacedChild",
            "MaxTraitsReplacedTeen",
            "LearnableSkills",
            "MaxSkillsLearnedChild",
            "MaxSkillsLearnedTeen",
            "MaxSkillLevelChild",
            "MaxSkillLevelTeen",
            "RecommendedOccupations",
            "RecommendedOccupationsStartLevel",
            "CarType",
            "BuffOrigin",
            "ConfirmationText",
            "ContinueEnrollmentText",
            "EnrollmentTNS",
            "HighSchoolGraduationText",
            "RemovalConfirmationText",
            "PhoneCallFlavorTNS",
            "PhoneCallHomesickTNS",
            "PhoneCallNeedMoneyTNS",
            "NeedMoneyRange",
            "PhoneCallRewardTNS",
            "RewardMoneyRange",
            "TNSIconName",
            "CareerPanelIconName",
            "CareerPanelText",
            "GraduationUniform",
            "GraduationSting",
        }

        populated_required_fields = {
            "GUID",
            "Name",
            "Cost",
            "SchoolTraits",
            "OffensiveTraits",
            "LearnableSkills",
            "RecommendedOccupations",
            "BuffOrigin",
            "ConfirmationText",
            "ContinueEnrollmentText",
            "EnrollmentTNS",
            "HighSchoolGraduationText",
            "GraduationUniform",
            "GraduationSting",
        }

        has_full_template = False
        populated_school_count = 0

        for school in direct_children:
            fields = {
                local_name(child.tag): (child.text or "").strip()
                for child in list(school)
            }

            field_names = set(fields)

            if template_fields.issubset(field_names):
                has_full_template = True

            if fields.get("GUID"):
                if not populated_required_fields.issubset(field_names):
                    return False

                if not all(
                    fields.get(name)
                    for name in (
                        "GUID",
                        "Name",
                        "Cost",
                        "SchoolTraits",
                        "LearnableSkills",
                        "RecommendedOccupations",
                        "BuffOrigin",
                        "GraduationUniform",
                    )
                ):
                    return False

                populated_school_count += 1

        return has_full_template and populated_school_count >= 5

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_scoring_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the NRaas-style Scoring XML framework.

    This does NOT replace permanent XML health. The marker is deliberately
    strict and schema-based.

    Required evidence:
    - root is Scoring
    - one or more populated Methods definitions exist
    - populated Methods entries contain Name + FullClassName
    - every named Method has a matching direct scoring table
    - every non-Methods direct child is one of those named scoring tables
    - each named scoring table contains at least one record with FullClassName
    - at least one scoring record has a populated FullClassName plus a
      characteristic scoring payload field
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "Scoring":
            return False

        direct_children = list(root)

        if not direct_children:
            return False

        method_names = set()

        for method in direct_children:
            if local_name(method.tag) != "Methods":
                continue

            fields = {
                local_name(child.tag): (child.text or "").strip()
                for child in list(method)
            }

            name = fields.get("Name", "")
            full_class_name = fields.get("FullClassName", "")

            # Empty template/default rows are allowed.
            if not name and not full_class_name:
                continue

            if not name or not full_class_name:
                return False

            method_names.add(name)

        if not method_names:
            return False

        non_method_children = [
            child
            for child in direct_children
            if local_name(child.tag) != "Methods"
        ]

        if not non_method_children:
            return False

        direct_scoring_names = {
            local_name(child.tag)
            for child in non_method_children
        }

        # Methods is the authoritative list of scoring tables.
        if direct_scoring_names != method_names:
            return False

        scoring_payload_fields = {
            "Trait",
            "Hit",
            "Miss",
            "HitMinimum",
            "HitMaximum",
            "Scoring",
            "LTW",
            "AgeGender",
            "Scale",
            "Score",
        }

        has_real_scoring_entry = False

        for method_name in method_names:
            matching = [
                child
                for child in non_method_children
                if local_name(child.tag) == method_name
            ]

            if not matching:
                return False

            has_full_class_field = False

            for scoring_record in matching:
                fields = {
                    local_name(child.tag): (child.text or "").strip()
                    for child in list(scoring_record)
                }

                if "FullClassName" in fields:
                    has_full_class_field = True

                if (
                    fields.get("FullClassName")
                    and any(
                        field_name in fields
                        for field_name in scoring_payload_fields
                    )
                ):
                    has_real_scoring_entry = True

            if not has_full_class_field:
                return False

        return has_real_scoring_entry

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_dreams_and_promises_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the Sims 3 DreamsAndPromisesNodes table.

    This does NOT replace permanent XML health. The marker is deliberately
    strict and based on the observed built-in-style primitive table.

    Required evidence:
    - root is DreamsAndPromisesNodes
    - every direct child is Primitives
    - every Primitives record contains the core Name, Category, Id, and
      RequiredProductVersions fields
    - at least one record contains TriggerEvent
    - at least one record contains CheckFunction or FeedbackFunction
    - at least one record contains a life-event field
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "DreamsAndPromisesNodes":
            return False

        direct_children = list(root)

        if not direct_children:
            return False

        core_fields = {
            "Name",
            "Category",
            "Id",
            "RequiredProductVersions",
        }

        has_trigger_event = False
        has_feedback_or_check = False
        has_life_event_field = False

        for primitive in direct_children:
            if local_name(primitive.tag) != "Primitives":
                return False

            field_names = {
                local_name(child.tag)
                for child in list(primitive)
            }

            if not core_fields.issubset(field_names):
                return False

            if "TriggerEvent" in field_names:
                has_trigger_event = True

            if (
                "CheckFunction" in field_names
                or "FeedbackFunction" in field_names
            ):
                has_feedback_or_check = True

            if (
                "IsLifeEvent" in field_names
                or "LifeEventNature" in field_names
                or "LifeEventDefaultSnapshot" in field_names
            ):
                has_life_event_field = True

        return (
            has_trigger_event
            and has_feedback_or_check
            and has_life_event_field
        )

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_dumpster_diving_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the Sims 3 DumpsterDiveData tuning table.

    This does NOT replace permanent XML health. The marker is deliberately
    strict and based on the observed table structure.

    Required evidence:
    - root is DumpsterDiveData
    - every direct child is RewardObjects
    - one RewardObjects record exposes the full default/template field set
    - at least one RewardObjects record has populated MedatorName, Quality,
      and ProductVersion values
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "DumpsterDiveData":
            return False

        direct_children = list(root)

        if not direct_children:
            return False

        if any(
            local_name(child.tag) != "RewardObjects"
            for child in direct_children
        ):
            return False

        full_template_fields = {
            "MedatorName",
            "Quality",
            "ProductVersion",
            "MaterialState",
            "Broken",
            "CanBeCleaned",
            "VFXSlot0",
            "VFXSlot1",
            "VFXSlot2",
            "VFXSlot3",
            "VFX0",
            "VFX1",
            "VFX2",
            "VFX3",
        }

        has_full_template = False
        has_populated_reward = False

        for reward in direct_children:
            fields = {
                local_name(child.tag): (child.text or "").strip()
                for child in list(reward)
            }

            if full_template_fields.issubset(set(fields)):
                has_full_template = True

            if all(
                fields.get(name)
                for name in (
                    "MedatorName",
                    "Quality",
                    "ProductVersion",
                )
            ):
                has_populated_reward = True

        return has_full_template and has_populated_reward

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_book_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the Sims 3 Books tuning table.

    This does NOT replace permanent XML health. The marker is deliberately
    strict and based on the observed built-in-style Books table.

    Required evidence:
    - root is Books
    - all direct records belong to established book-record families
    - BookGeneral, BookSkill, and SheetMusic records are all present
    - representative records expose their characteristic schema fields
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "Books":
            return False

        direct_children = list(root)

        if not direct_children:
            return False

        allowed_record_types = {
            "BookGeneral",
            "BookSkill",
            "BookComics",
            "BookRecipe",
            "SheetMusic",
            "AcademicTextBook",
            "BookFish",
            "BookToddler",
            "WrittenBookTitles",
            "MedicalJournal",
            "BookAlchemyRecipe",
            "WrittenBookMaterials",
        }

        direct_names = [local_name(child.tag) for child in direct_children]

        if any(name not in allowed_record_types for name in direct_names):
            return False

        if not {"BookGeneral", "BookSkill", "SheetMusic"}.issubset(set(direct_names)):
            return False

        required_fields = {
            "BookGeneral": {
                "Title",
                "Author",
                "Genre",
                "Length",
                "Value",
                "CodeVersion",
            },
            "BookSkill": {
                "Title",
                "Skill",
                "MinSkill",
                "MaxSkill",
                "ChildCanRead",
                "CodeVersion",
            },
            "SheetMusic": {
                "Title",
                "SkillName",
                "SkillLevel",
                "AudioClip",
                "ChildCanRead",
                "CodeVersion",
            },
        }

        for record_type, fields_needed in required_fields.items():
            matching = [
                child
                for child in direct_children
                if local_name(child.tag) == record_type
            ]

            if not matching:
                return False

            found_schema_match = False

            for record in matching:
                field_names = {
                    local_name(child.tag)
                    for child in list(record)
                }

                if fields_needed.issubset(field_names):
                    found_schema_match = True
                    break

            if not found_schema_match:
                return False

        return True

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_career_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the standard Sims 3 Careers XML table.

    This does NOT replace permanent XML health. It deliberately distinguishes
    the built-in-style Careers table from the existing custom CareerFiles
    loader schema used by the permanent "Career" classifier.

    Required evidence:
    - root is Careers
    - at least one direct CareerList record
    - at least one populated CareerList has CareerName, FullClassName, Guid,
      and TableName
    - at least one direct per-career level block has the characteristic
      Title/BasePay/Level/BranchName/Metric1 fields
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "Careers":
            return False

        direct_children = list(root)

        if not direct_children:
            return False

        career_lists = [
            child
            for child in direct_children
            if local_name(child.tag) == "CareerList"
        ]

        if not career_lists:
            return False

        populated_career_record = False

        for career_list in career_lists:
            fields = {
                local_name(child.tag): (child.text or "").strip()
                for child in list(career_list)
            }

            if all(
                fields.get(name)
                for name in (
                    "CareerName",
                    "FullClassName",
                    "Guid",
                    "TableName",
                )
            ):
                populated_career_record = True
                break

        if not populated_career_record:
            return False

        # Career level blocks are direct children whose tag is the career/table
        # name rather than CareerList. Their records expose the standard level
        # fields below.
        required_level_fields = {
            "Title",
            "BasePay",
            "Level",
            "BranchName",
            "Metric1",
        }

        for child in direct_children:
            if local_name(child.tag) == "CareerList":
                continue

            child_fields = {
                local_name(grandchild.tag)
                for grandchild in list(child)
            }

            if required_level_fields.issubset(child_fields):
                return True

        return False

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False



def detect_career_overhaul_tuning_xml_marker(package, entry):
    """
    Classification-only detector for the validated MFC career-overhaul XML
    tables that are not part of the standard Careers/CareerList schema.

    Returns:
      - "ActiveCareers"
      - "SkillBasedCareers"
      - None

    This does NOT replace permanent XML health. It requires a specific root
    plus multiple characteristic schema fields so a generic NMAP + XML tuning
    package cannot become Career Tuning from a filename or a single word.
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return None

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        root_name = local_name(root.tag)
        descendant_names = {
            local_name(element.tag)
            for element in root.iter()
        }

        if root_name == "ActiveCareers":
            required = {
                "ActiveCareer",
                "FullClassName",
                "Highest_Level",
                "XP_Gain_Multiplier_Contribution_Max",
            }

            if required.issubset(descendant_names):
                return "ActiveCareers"

        if root_name == "SkillBasedCareers":
            required = {
                "SkilledProfession",
                "Skill_Name",
                "Minimum_Skill_Level",
                "XP_Gain_Multiplier",
                "Highest_Level",
            }

            if required.issubset(descendant_names):
                return "SkillBasedCareers"

        return None

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return None

def detect_academic_degree_tuning_xml_marker(package, entry):
    """
    Classification-only detector for Sims 3 University AcademicDegrees XML.

    This does NOT replace permanent XML health. It deliberately requires the
    observed AcademicDegrees schema: the root must be AcademicDegrees, every
    direct child must be a Degrees record, and every record must expose the
    core degree identity / credit-hours / cost fields.
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "AcademicDegrees":
            return False

        degree_records = list(root)

        if not degree_records:
            return False

        if any(local_name(child.tag) != "Degrees" for child in degree_records):
            return False

        required_fields = {
            "AcademicDegreeEnum",
            "AcademicDegreeRequiredCreditHours",
            "AcademicDegreeCostPerCredit",
        }

        for degree in degree_records:
            field_names = {
                local_name(child.tag)
                for child in list(degree)
            }

            if not required_fields.issubset(field_names):
                return False

            enum_node = next(
                (
                    child for child in list(degree)
                    if local_name(child.tag) == "AcademicDegreeEnum"
                ),
                None
            )

            if enum_node is None or not (enum_node.text or "").strip():
                return False

        return True

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_opportunity_tuning_xml_marker(package, entry):
    """
    Classification-only detector for Sims 3 Opportunity definition XML.

    This does NOT participate in permanent XML structural health. The marker
    is deliberately strict and based on the observed Opportunity resource
    schema: an Opportunities_* root containing repeated, balanced direct
    blocks of OpportunitiesSetup, Names, OpportunitiesRequirements, and
    OpportunitiesCompletion. At least one setup must contain a non-empty GUID.
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        root = ET.fromstring(decoded.strip("\x00 \t\r\n"))

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        root_name = local_name(root.tag)

        if not root_name.startswith("Opportunities_"):
            return False

        required_blocks = {
            "OpportunitiesSetup",
            "Names",
            "OpportunitiesRequirements",
            "OpportunitiesCompletion",
        }

        direct_children = list(root)

        if not direct_children:
            return False

        direct_names = [local_name(child.tag) for child in direct_children]

        if any(name not in required_blocks for name in direct_names):
            return False

        counts = {
            name: direct_names.count(name)
            for name in required_blocks
        }

        if any(counts[name] <= 0 for name in required_blocks):
            return False

        # The observed Opportunity format is made of one matching block of
        # each type per opportunity record.
        if len(set(counts.values())) != 1:
            return False

        has_nonempty_guid = False

        for child in direct_children:
            if local_name(child.tag) != "OpportunitiesSetup":
                continue

            for setup_child in list(child):
                if (
                    local_name(setup_child.tag) == "GUID"
                    and (setup_child.text or "").strip()
                ):
                    has_nonempty_guid = True
                    break

            if has_nonempty_guid:
                break

        return has_nonempty_guid

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_tuning_xml_marker(package, entry):
    """
    Classification-only detector for Sims 3 XML tuning resources.

    This deliberately does NOT participate in XML structural health. The
    permanent XML health checker remains unchanged. A resource counts as
    tuning evidence only when the parsed XML has a <base> root containing a
    non-empty <Current_Tuning> section.
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        cleaned = decoded.strip("\x00 \t\r\n")
        root = ET.fromstring(cleaned)

        root_tag = str(root.tag).split("}")[-1]

        if root_tag != "base":
            return False

        current_tuning = None

        for child in list(root):
            child_tag = str(child.tag).split("}")[-1]

            if child_tag == "Current_Tuning":
                current_tuning = child
                break

        if current_tuning is None:
            return False

        # Empty tuning containers are not enough evidence by themselves.
        if len(list(current_tuning)) == 0:
            return False

        return True

    except Exception:
        return False



def detect_interaction_tuning_marker(package, entry):
    """
    Classification-only detector for Sims 3 ITUN / Interaction Tuning.

    Resource type 0x03B33DDF is the Sims 3 ITUN resource type. To avoid
    classifying a malformed or unrelated payload from the type ID alone, the
    resource must also parse as the established Sims 3 ITUN XML structure:
    <base> with direct <Interaction name=...>, <Object name=...>, and a
    non-empty <Current_Tuning> section. <CodeVersion> is deliberately not
    required because valid older ITUN resources may omit it.
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        cleaned = decoded.strip("\x00 \t\r\n")
        root = ET.fromstring(cleaned)

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "base":
            return False

        interaction = None
        obj = None
        current_tuning = None

        for child in list(root):
            child_tag = local_name(child.tag)

            if child_tag == "Interaction" and interaction is None:
                interaction = child
            elif child_tag == "Object" and obj is None:
                obj = child
            elif child_tag == "Current_Tuning" and current_tuning is None:
                current_tuning = child

        if interaction is None or not (interaction.attrib.get("name") or "").strip():
            return False

        if obj is None or not (obj.attrib.get("name") or "").strip():
            return False

        if current_tuning is None or len(list(current_tuning)) == 0:
            return False

        return True

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def detect_random_event_module_xml_marker(package, entry):
    """
    Classification-only detector for custom RandomEvents module XML.

    This does NOT participate in XML structural health. The marker is
    intentionally strict: the XML root must be <RandomEvents>, it must have
    at least one direct <Event> child, and an Event must contain the core
    structural fields used by this event-module schema.
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        raw = bytes(resource.AsBytes)

        decode_attempts = []

        if raw.startswith(b"\xEF\xBB\xBF"):
            decode_attempts.append("utf-8-sig")
        elif raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            decode_attempts.append("utf-16")

        decode_attempts.extend([
            "utf-8",
            "utf-16-le",
            "utf-16-be",
        ])

        decoded = None

        for codec in decode_attempts:
            try:
                candidate = raw.decode(codec)

                if codec not in ("utf-16", "utf-16-le", "utf-16-be"):
                    nul_count = candidate.count("\x00")
                    if nul_count > max(4, len(candidate) // 4):
                        continue

                if not candidate.replace("\x00", "").strip():
                    continue

                decoded = candidate
                break

            except Exception:
                continue

        if decoded is None:
            return False

        cleaned = decoded.strip("\x00 \t\r\n")
        root = ET.fromstring(cleaned)

        def local_name(tag):
            return str(tag).rsplit("}", 1)[-1]

        if local_name(root.tag) != "RandomEvents":
            return False

        required_event_fields = {
            "TypeEvent",
            "Rarity",
            "GeneralMethod",
            "STBLCaptionKey",
            "STBLTitleKey",
        }

        for child in list(root):
            if local_name(child.tag) != "Event":
                continue

            event_fields = {
                local_name(event_child.tag)
                for event_child in list(child)
            }

            if required_event_fields.issubset(event_fields):
                return True

        return False

    except Exception:
        # Classification must never turn a readable package into BROKEN.
        return False


def check_stbl_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 STBL / String Table.

    STBL has a real S3PI parser. Any exception while deserializing the
    resource is treated as structural corruption.

    No fixed table size, language count, or entry count is assumed.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"STBL HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_nmap_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 NMAP / Name Map
    resources (0x0166038C; displayed as _KEY in S3PE).

    S3PI has a dedicated NameMapResource parser. Any exception while
    deserializing the resource is treated as structural corruption.

    No fixed file size or entry count is assumed.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"NMAP HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_rslt_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 RSLT / Slot Definition
    resources (0xD3044521).

    RSLT is parsed through S3PI's GenericRCOLResource handling.
    Any exception while deserializing the resource is treated as
    structural corruption.

    No fixed version or byte-size requirement is imposed.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"RSLT HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_ftpt_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 FTPT / Model Footprint.

    FTPT is handled through S3PI RCOL parsing. Parser failures caused by
    truncated/unreadable structure are treated as corruption.

    No fixed version or byte-size requirement is imposed.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"FTPT HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_rig_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 RIG resources
    (0x8EAF13DE).

    The scanner explicitly loads s3pi.RigResource.dll so clear-format RIGs
    are validated by the dedicated RigResource parser.

    Any parser exception while deserializing a recognized RIG structure is
    treated as structural corruption. No fixed bone count or byte size is
    required.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"RIG HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_jazz_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 JAZZ / State Machine
    resources (0x02D5DF13).

    The scanner explicitly loads s3pi.JazzResource.dll. JAZZ resources are
    deserialized through S3PI's RCOL handling with the JAZZ chunk handlers
    available. Any parser exception is treated as structural corruption.

    No fixed byte size or exact state-machine complexity is required.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"JAZZ HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_s3sa_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 S3SA / Script Assembly
    resources (0x073FAA07).

    The scanner explicitly loads s3pi.ScriptResource.dll so S3SA resources
    are validated by the dedicated ScriptResource parser.

    Any parser exception is treated as structural corruption.
    No fixed byte size, hash value, or script assembly length is required.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"S3SA HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_lite_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 LITE resources
    (0x03B4C61D).

    LITE is handled through S3PI GenericRCOLResource. Any parser exception
    while deserializing the resource is treated as structural corruption.

    No fixed byte size or exact light-data complexity is required.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"LITE HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_clip_health(package, entry, resource_objects):
    """
    Permanent conservative raw structural health check for Sims 3 CLIP
    animation resources (0x6B20C4F3).

    S3PI falls back to DefaultResource for CLIP in this installation, so
    parser-exception validation is unavailable.

    Confirmed conservative rules:
      * resource must be at least 32 bytes
      * first DWORD must equal little-endian CLIP magic/type 0x6B20C4F3

    Other header DWORDs are intentionally NOT hard-coded yet because valid
    animation layouts may vary.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        raw = bytes(resource.AsBytes)

    except Exception as e:
        reason = f"CLIP read failure: {type(e).__name__}: {e}"
        print(f"CLIP HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason

    byte_count = len(raw)

    if byte_count < 32:
        reason = f"CLIP truncated header: {byte_count} bytes"
        print(f"CLIP HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason

    magic = int.from_bytes(raw[0:4], "little", signed=False)

    if magic != CLIP_TYPE:
        reason = (
            f"CLIP bad magic: 0x{magic:08X}; "
            f"expected 0x{CLIP_TYPE:08X}"
        )
        print(f"CLIP HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason

    return True, ""



def get_clip_framecount_for_classification(package, entry, resource_objects):
    """
    Classification-only conservative reader for Sims 3 CLIP framecount.
    Returns None when the expected minimum CLIP structure is not present.
    """
    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)
        raw = bytes(resource.AsBytes)

        if len(raw) < 32:
            return None

        if int.from_bytes(raw[0:4], "little", signed=False) != CLIP_TYPE:
            return None

        clip_relative_offset = int.from_bytes(
            raw[12:16],
            "little",
            signed=False
        )
        clip_header_start = 12 + clip_relative_offset

        if clip_header_start + 22 > len(raw):
            return None

        if raw[clip_header_start:clip_header_start + 8] != b"_pilC3S_":
            return None

        return int.from_bytes(
            raw[clip_header_start + 20:clip_header_start + 22],
            "little",
            signed=False
        )

    except Exception:
        return None


def check_matd_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 MATD / Material Definition
    resources (0x01D0E75D).

    MATD is handled as an RCOL resource by S3PI. Any exception while
    WrapperDealer / GenericRCOLResource deserializes the resource is treated
    as structural corruption.

    No fixed byte size or exact MATD layout is required.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"MATD HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_mtst_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 MTST resources
    (0x02019972).

    MTST is handled as an RCOL resource by S3PI. Any exception while
    WrapperDealer / GenericRCOLResource deserializes it is treated as
    structural corruption.

    No fixed resource size, version, or DWORD layout is required.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"MTST HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_mdlr_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 MDLR resources
    (resource type 0xCF9A4ACE).

    The dedicated s3pi.ModularResource.dll wrapper is used. Any exception
    while ModularResource.ModularResource deserializes the resource is
    treated as structural corruption.

    Valid examples observed in the user's collection include multiple
    different sizes (58, 90, 106 and 122 bytes), so no fixed size/version
    rule is used.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"MDLR HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_cwal_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 CWAL resources
    (Wall / Floor Pattern Catalog Resource, type 0x515CA4CD).

    S3PI handles CWAL through
    CatalogResource.WallFloorPatternCatalogResource.

    Any parser exception is treated as structural corruption.
    No fixed resource size/version rule is used.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"CWAL HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_cfen_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 CFEN resources
    (Catalog Fence Resource, type 0x0418FE2A).

    S3PI handles CFEN through CatalogResource.FenceCatalogResource.

    Any parser exception is treated as structural corruption.
    No fixed resource size/version rule is used.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"CFEN HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_ctpt_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 CTPT resources
    (Terrain Paint / Paving Catalog Resource, type 0x04ED4BB2).

    S3PI handles CTPT through CatalogResource.TerrainPaintBrushCatalogResource.

    Any parser exception is treated as structural corruption.
    No fixed resource size/version/count rule is used.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"CTPT HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_ngmp_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 NGMP resources
    (type 0xF3A38370).

    S3PI handles NGMP through
    NGMPHashMapResource.NGMPHashMapResource.

    Any parser exception is treated as structural corruption.
    No fixed resource-size rule is used.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"NGMP HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_crmt_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 CRMT resources
    (Catalog Roof Pattern Resource, type 0xF1EDBD86).

    S3PI handles CRMT through CatalogResource.RoofPatternCatalogResource.

    Any parser exception is treated as structural corruption.
    No fixed resource size/version/TGI-count rule is used.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"CRMT HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_spt2_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 SPT2 resources
    (SpeedTree / TREE data, type 0x021D7E8C).

    S3PI handles SPT2 through GenericRCOLResource.
    Any parser exception is treated as structural corruption.

    No fixed resource size or TREE payload length is enforced.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"SPT2 HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_spt_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 _SPT resources
    (type 0x00B552EA).

    S3PI exposes _SPT only as DefaultResource, so validation is raw.

    Conservative rules:
      - initial header must exist
      - declared signature length must be plausible and fit in the stream
      - signature must belong to the "__IdvSpt_" family
      - the early declared string/path length field must exist
      - its declared payload must fit in the stream

    No fixed total resource size, exact version number, exact signature
    suffix, exact path length, or exact path contents are required.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        if byte_count < 8:
            reason = f"resource too short for _SPT header ({byte_count} bytes)"
            print(f"_SPT HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        sig_len = int.from_bytes(
            raw[4:8],
            "little",
            signed=False
        )

        if sig_len == 0 or sig_len > 256:
            reason = f"invalid declared _SPT signature length {sig_len}"
            print(f"_SPT HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        sig_start = 8
        sig_end = sig_start + sig_len

        if sig_end > byte_count:
            reason = (
                f"_SPT signature truncated: need through byte {sig_end}, "
                f"have {byte_count}"
            )
            print(f"_SPT HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        sig_bytes = raw[sig_start:sig_end]

        try:
            signature = sig_bytes.decode("ascii")
        except Exception:
            reason = "_SPT signature is not valid ASCII"
            print(f"_SPT HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if not signature.startswith("__IdvSpt_"):
            reason = f"unexpected _SPT signature {signature!r}"
            print(f"_SPT HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        # Observed family structure:
        # [signature] + DWORD + DWORD + DWORD declared early payload length
        path_len_field = sig_end + 8
        payload_start = sig_end + 12

        if payload_start > byte_count:
            reason = "_SPT truncated before early declared payload"
            print(f"_SPT HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        payload_len = int.from_bytes(
            raw[path_len_field:path_len_field + 4],
            "little",
            signed=False
        )

        # Avoid absurd lengths without assuming a specific clean value.
        if payload_len > 16 * 1024 * 1024:
            reason = f"implausible _SPT early payload length {payload_len}"
            print(f"_SPT HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        payload_end = payload_start + payload_len

        if payload_end > byte_count:
            reason = (
                f"_SPT early payload truncated: need through byte "
                f"{payload_end}, have {byte_count}"
            )
            print(f"_SPT HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"_SPT HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_cral_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 CRAL resources
    (Catalog Railing Resource, type 0x04C58103).

    S3PI handles CRAL through CatalogResource.RailingCatalogResource.

    Any parser exception is treated as structural corruption.
    No fixed resource size/version/material/TGI-count rule is used.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"CRAL HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_layo_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 LAYO resources
    (UI Layout Resource, type 0x025C95B6).

    S3PI exposes LAYO as DefaultResource, so parser success is not useful.
    Validation is based on the raw layout text:

      - must decode as UTF-8/UTF-16 text
      - must contain parseable XML
      - root element must be <graph>
      - graph class/type, when present, must identify Layout

    No fixed byte size, control count, object count, or resource-ID rule.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        if byte_count == 0:
            reason = "empty LAYO resource"
            print(f"LAYO HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        decoded = None
        used_encoding = None

        for enc in ("utf-8-sig", "utf-16-le", "utf-16-be"):
            try:
                candidate = raw.decode(enc)
            except Exception:
                continue

            # LAYO is XML-like; reject accidental garbage decodes.
            if "<" in candidate and ">" in candidate:
                decoded = candidate
                used_encoding = enc
                break

        if decoded is None:
            reason = "LAYO text decode failed"
            print(f"LAYO HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        cleaned = decoded.lstrip("\ufeff\x00 \t\r\n")
        lt = cleaned.find("<")

        if lt < 0:
            reason = "LAYO contains no XML element"
            print(f"LAYO HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        xml_text = cleaned[lt:]

        try:
            root = ET.fromstring(xml_text)
        except Exception as e:
            reason = f"malformed LAYO XML: {type(e).__name__}: {e}"
            print(f"LAYO HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        root_tag = str(root.tag).split("}")[-1].lower()

        if root_tag == "metadata":
            # Compatibility: a known-good main-menu UI mod uses an intentionally
            # empty <metadata> LAYO as a stub alongside its normal Layout graph.
            # Keep this narrow: attributes, child elements, or non-whitespace
            # text are not accepted by this compatibility path.
            metadata_text = str(root.text or "").strip()
            metadata_children = list(root)

            if not root.attrib and not metadata_children and not metadata_text:
                print(
                    f"LAYO EMPTY METADATA COMPATIBILITY: {resource_id} | "
                    f"Bytes={byte_count:,} | Encoding={used_encoding} | Status=OK"
                )
                return True, ""

            reason = "non-empty LAYO metadata stub is not validated"
            print(f"LAYO HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if root_tag != "graph":
            reason = f"unexpected LAYO root element {root.tag!r}"
            print(f"LAYO HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        graph_class = str(root.attrib.get("class", "")).strip()
        graph_type = str(root.attrib.get("type", "")).strip()

        if graph_class and graph_class.lower() != "layout":
            reason = f"unexpected LAYO graph class {graph_class!r}"
            print(f"LAYO HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if graph_type and graph_type.lower() != "layout":
            reason = f"unexpected LAYO graph type {graph_type!r}"
            print(f"LAYO HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"LAYO HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_hinf_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 HINF resources
    (Hair Information, type 0x063261DA).

    S3PI exposes HINF as DefaultResource, so validation is raw.

    Conservative observed structure:
      - at least 8-byte header
      - first DWORD acts as a record-count-like field
      - resource must contain at least 8 + (count * 8) bytes

    No fixed total size, fixed count, or fixed second DWORD is required.
    Extra trailing bytes are allowed.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        if byte_count < 8:
            reason = (
                f"HINF too short for header: "
                f"have {byte_count}, need at least 8"
            )
            print(f"HINF HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        record_count = int.from_bytes(
            raw[0:4],
            "little",
            signed=False
        )

        # Defensive guard against absurd/corrupt count values.
        if record_count > 1000000:
            reason = f"implausible HINF record count {record_count}"
            print(f"HINF HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        expected_min = 8 + (record_count * 8)

        if expected_min > byte_count:
            reason = (
                f"HINF declared records truncated: "
                f"need at least {expected_min} bytes, have {byte_count}"
            )
            print(f"HINF HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"HINF HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_2ary_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 2ARY resources
    (type 0x05FF6BA4).

    S3PI exposes 2ARY as DefaultResource, so validation is raw.

    Proven structure from intact world resources:
      - at least 12-byte header
      - DWORD 0 is a non-zero header/version-like field
      - DWORD 1 and DWORD 2 are non-zero array dimensions
      - remaining payload represents a rectangular array
      - payload size must be divisible by (DimA * DimB)

    The derived bytes-per-cell value is NOT hardcoded to 1 or 4.
    A broad sanity guard of 1..64 bytes per cell is used.
    No fixed total size or fixed dimensions are required.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        if byte_count < 12:
            reason = (
                f"2ARY too short for header: "
                f"have {byte_count}, need at least 12"
            )
            print(f"2ARY HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        field0 = int.from_bytes(raw[0:4], "little", signed=False)
        dim_a = int.from_bytes(raw[4:8], "little", signed=False)
        dim_b = int.from_bytes(raw[8:12], "little", signed=False)
        payload_bytes = byte_count - 12

        if field0 == 0:
            reason = "2ARY zero header/version field"
            print(f"2ARY HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if dim_a == 0 or dim_b == 0:
            reason = f"2ARY zero dimension: {dim_a} x {dim_b}"
            print(f"2ARY HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if dim_a > 1000000 or dim_b > 1000000:
            reason = f"implausible 2ARY dimensions: {dim_a} x {dim_b}"
            print(f"2ARY HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        cell_count = dim_a * dim_b

        if payload_bytes <= 0:
            reason = "2ARY has no array payload"
            print(f"2ARY HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if payload_bytes % cell_count != 0:
            reason = (
                f"2ARY payload does not fit declared dimensions: "
                f"{payload_bytes} payload bytes for {dim_a} x {dim_b}"
            )
            print(f"2ARY HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        element_size = payload_bytes // cell_count

        if element_size < 1 or element_size > 64:
            reason = (
                f"implausible 2ARY element size "
                f"{element_size} bytes"
            )
            print(f"2ARY HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"2ARY HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_complate_health(package, entry, resource_objects):
    """
    Permanent health check for Sims 3 Complate resources
    (type 0x044AE110).

    Dedicated wrapper:
      ComplateResource.ComplateResource

    Conservative rule:
      any wrapper/parser exception => BROKEN

    No fixed total size, version, count, or internal-value rule.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        _ = resource.AsBytes
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"COMPLATE HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_63a33ea7_health(package, entry, resource_objects):
    """
    Permanent health check for Sims 3 resource type 0x63A33EA7.

    Confirmed wrapper:
      s3pi.GenericRCOLResource.GenericRCOLResource

    Conservative rule:
      any GenericRCOL parser/wrapper exception => BROKEN

    No fixed byte size, version, RCOL tag, count, or internal-value rule.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        _ = resource.AsBytes
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"63A33EA7 HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_fireplace_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 Fireplace Catalog resources
    (type 0x04F3CC01).

    Confirmed dedicated wrapper:
      CatalogResource.FireplaceCatalogResource

    Diagnostic validation:
      - deliberately truncated 8-byte resource -> EndOfStreamException
      - untouched 405-byte control -> parses normally

    Permanent rule:
      - dedicated wrapper/parser exception = BROKEN
      - no fixed total size, version, TGI count, or internal layout rule
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        # Touch serialized bytes so lazy wrapper failures cannot be hidden.
        _ = resource.AsBytes

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(
            f"FIREPLACE HEALTH FAILURE: "
            f"{resource_id} | {reason}"
        )
        return False, reason



def _validate_known_bmp_thumbnail(raw):
    """
    Conservative permanent validator for the real BMP thumbnail variant discovered in
    pitheinfinite_contour_01.package on 2026-08-22.

    Intentionally accepts only the exact conventional variant now proven:
      - BITMAPFILEHEADER + 40-byte BITMAPINFOHEADER
      - 32 bits per pixel
      - BI_RGB / uncompressed
      - pixel data begins at byte 54
      - declared file size exactly equals resource length
      - computed pixel array ends exactly at EOF

    Returns: (ok, reason, width, height)
    """
    byte_count = len(raw)

    if byte_count < 54:
        return False, f"BMP truncated: {byte_count} bytes (minimum supported header is 54)", None, None

    if raw[:2] != b"BM":
        return False, "BMP signature mismatch", None, None

    declared_size = int.from_bytes(raw[2:6], "little", signed=False)
    reserved1 = int.from_bytes(raw[6:8], "little", signed=False)
    reserved2 = int.from_bytes(raw[8:10], "little", signed=False)
    pixel_offset = int.from_bytes(raw[10:14], "little", signed=False)
    dib_size = int.from_bytes(raw[14:18], "little", signed=False)
    width = int.from_bytes(raw[18:22], "little", signed=True)
    height = int.from_bytes(raw[22:26], "little", signed=True)
    planes = int.from_bytes(raw[26:28], "little", signed=False)
    bits_per_pixel = int.from_bytes(raw[28:30], "little", signed=False)
    compression = int.from_bytes(raw[30:34], "little", signed=False)
    image_size = int.from_bytes(raw[34:38], "little", signed=False)

    if declared_size != byte_count:
        return False, (
            f"BMP declared file size mismatch: header={declared_size}, "
            f"resource={byte_count}"
        ), None, None

    if reserved1 != 0 or reserved2 != 0:
        return False, (
            f"BMP reserved fields are non-zero: {reserved1}, {reserved2}"
        ), None, None

    if dib_size != 40:
        return False, (
            f"BMP unsupported DIB header size: {dib_size} "
            f"(only proven 40-byte BITMAPINFOHEADER accepted)"
        ), None, None

    if pixel_offset != 54:
        return False, (
            f"BMP unsupported pixel-data offset: {pixel_offset} "
            f"(only proven offset 54 accepted)"
        ), None, None

    if width <= 0 or height == 0:
        return False, f"BMP invalid dimensions: {width}x{height}", None, None

    if planes != 1:
        return False, f"BMP invalid planes value: {planes}", None, None

    if bits_per_pixel != 32:
        return False, (
            f"BMP unsupported bit depth: {bits_per_pixel} "
            f"(only proven 32-bit variant accepted)"
        ), None, None

    if compression != 0:
        return False, (
            f"BMP unsupported compression: {compression} "
            f"(only proven BI_RGB/uncompressed variant accepted)"
        ), None, None

    abs_height = abs(height)
    row_stride = ((width * bits_per_pixel + 31) // 32) * 4
    pixel_bytes = row_stride * abs_height
    expected_end = pixel_offset + pixel_bytes

    if image_size not in (0, pixel_bytes):
        return False, (
            f"BMP image-size field inconsistent: header={image_size}, "
            f"computed={pixel_bytes}"
        ), None, None

    if expected_end != byte_count:
        return False, (
            f"BMP pixel data does not end at resource EOF: "
            f"expected_end={expected_end}, resource={byte_count}"
        ), None, None

    return True, "", width, abs_height



def check_thum_health(package, entry, resource_objects):
    """
    Structural health check for Sims 3 THUM type 0x626F60CD (and the existing
    compatible THUM route that reuses this checker).

    Accepted image forms remain PNG, JPEG and the already-supported BMP form.
    A PNG whose sole defect is a missing final IEND is accepted as recoverable
    only after full chunk/CRC/IDAT validation.

    Real-world 0x626F60CD compatibility variant:
    some CAS packages contain a non-image 0x626F60CD binary preset record with
    a well-formed UTF-16LE <preset><complate ...> block. That narrow structure
    is valid metadata and must not make the package BROKEN.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"
    resource_type_text = f"0x{int(entry.ResourceType):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = b"\x89PNG\r\n\x1a\n"
        jpeg_soi = b"\xFF\xD8"
        jpeg_eoi = b"\xFF\xD9"
        bmp_sig = b"BM"

        if raw.startswith(png_sig):
            if byte_count < 24:
                reason = (
                    f"THUM {resource_type_text} PNG truncated: {byte_count} bytes "
                    f"(too short for PNG structure)"
                )
                print(f"THUM HEALTH FAILURE: {resource_id} | Group={group_id} | {reason}")
                return False, reason

            width = int.from_bytes(raw[16:20], "big", signed=False)
            height = int.from_bytes(raw[20:24], "big", signed=False)
            if width <= 0 or height <= 0:
                reason = f"THUM {resource_type_text} PNG invalid dimensions: {width}x{height}"
                print(f"THUM HEALTH FAILURE: {resource_id} | Group={group_id} | {reason}")
                return False, reason

            if not _png_has_terminal_iend(raw):
                recoverable, checked_width, checked_height = _validate_png_missing_iend_only(raw)
                if recoverable:
                    width = checked_width
                    height = checked_height
                    print(
                        f"THUM HEALTH WARNING: {resource_id} | Group={group_id} | "
                        f"THUM {resource_type_text} PNG missing IEND terminator only; "
                        f"chunk/CRC/IDAT data validated; Status=OK (recoverable)"
                    )
                    return True, ""

                reason = f"THUM {resource_type_text} PNG missing/invalid IEND terminator"
                print(f"THUM HEALTH FAILURE: {resource_id} | Group={group_id} | {reason}")
                return False, reason

            print(
                f"{resource_type_text} THUM CHECK: {resource_id} | Group={group_id} | "
                f"Format=PNG | Bytes={byte_count:,} | Size={width}x{height} | Status=OK"
            )
            return True, ""

        if raw.startswith(jpeg_soi):
            if byte_count < 4:
                reason = f"THUM {resource_type_text} JPEG truncated: {byte_count} bytes"
                print(f"THUM HEALTH FAILURE: {resource_id} | Group={group_id} | {reason}")
                return False, reason
            if not raw.endswith(jpeg_eoi):
                reason = f"THUM {resource_type_text} JPEG missing EOI terminator"
                print(f"THUM HEALTH FAILURE: {resource_id} | Group={group_id} | {reason}")
                return False, reason

            width = height = None
            try:
                width, height = _diagnose_jpeg_dimensions_626f60cc(raw)
            except Exception:
                pass
            size_text = f"{width}x{height}" if width and height else "unknown"
            print(
                f"{resource_type_text} THUM CHECK: {resource_id} | Group={group_id} | "
                f"Format=JPEG | Bytes={byte_count:,} | Size={size_text} | Status=OK"
            )
            return True, ""

        if raw.startswith(bmp_sig):
            ok, bmp_reason, width, height = _validate_known_bmp_thumbnail(raw)
            if not ok:
                reason = f"THUM {resource_type_text} {bmp_reason}"
                print(f"THUM HEALTH FAILURE: {resource_id} | Group={group_id} | {reason}")
                return False, reason
            print(
                f"{resource_type_text} THUM CHECK: {resource_id} | Group={group_id} | "
                f"Format=BMP | Bytes={byte_count:,} | Size={width}x{height} | BPP=32 | Status=OK"
            )
            return True, ""

        if int(entry.ResourceType) == THUM_TYPE and _validate_626f60cd_preset_metadata_variant(raw):
            print(
                f"THUM COMPATIBILITY WARNING: {resource_id} | Group={group_id} | "
                f"Type={resource_type_text} | Format=CAS preset metadata variant | "
                f"Status=OK (non-image 0x626F60CD structure validated)"
            )
            return True, ""

        reason = (
            f"THUM {resource_type_text} unsupported/signature mismatch: "
            f"{raw[:16].hex().upper()}"
        )
        print(f"THUM HEALTH FAILURE: {resource_id} | Group={group_id} | {reason}")
        return False, reason

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"THUM HEALTH FAILURE: {resource_id} | Group={group_id} | {reason}")
        return False, reason



def check_626f60ce_health(package, entry, resource_objects):
    """
    Structural health test for Sims 3 resource type 0x626F60CE.

    Previously validated PNG route:
      - multiple clean PNG samples
      - deliberately truncated PNG control detected
      - PNG signature, dimensions and terminal IEND validated

    New 2026-08-20 real-world evidence:
      - lip_15_ulito4ka.package contains a 0x626F60CE payload beginning
        FF D8 FF E1 ... Exif, proving this type can also contain JPEG.

    Permanent PNG/JPEG/BMP rules:
      PNG:
        - at least 24 bytes
        - standard PNG signature
        - non-zero width and height
        - valid terminal IEND marker
      JPEG:
        - starts FF D8 (SOI)
        - at least 4 bytes
        - ends FF D9 (EOI)
        - SOF dimensions are best-effort metadata only

    Any payload that is neither PNG nor JPEG is a health failure.
    PNG/JPEG/BMP handling is permanent.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = b"\x89PNG\r\n\x1a\n"
        jpeg_soi = b"\xFF\xD8"
        jpeg_eoi = b"\xFF\xD9"
        bmp_sig = b"BM"

        # PNG route -- already permanently validated.
        if raw.startswith(png_sig):
            if byte_count < 24:
                reason = (
                    f"0x626F60CE PNG truncated: {byte_count} bytes "
                    f"(too short for PNG structure)"
                )
                print(
                    f"626F60CE HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            width = int.from_bytes(raw[16:20], "big", signed=False)
            height = int.from_bytes(raw[20:24], "big", signed=False)

            if width <= 0 or height <= 0:
                reason = f"0x626F60CE PNG invalid dimensions: {width}x{height}"
                print(
                    f"626F60CE HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            iend_ok = (
                byte_count >= 12
                and raw[-12:-8] == b"\x00\x00\x00\x00"
                and raw[-8:-4] == b"IEND"
            )

            if not iend_ok:
                reason = "0x626F60CE PNG missing/invalid IEND terminator"
                print(
                    f"626F60CE HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            print(
                f"626F60CE CHECK: {resource_id} | Group={group_id} | "
                f"Format=PNG | Bytes={byte_count:,} | Size={width}x{height} | Status=OK"
            )
            return True, ""

        # JPEG route -- newly discovered real-world variant, under test.
        if raw.startswith(jpeg_soi):
            if byte_count < 4:
                reason = (
                    f"0x626F60CE JPEG truncated: {byte_count} bytes "
                    f"(too short for JPEG structure)"
                )
                print(
                    f"626F60CE HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            if not raw.endswith(jpeg_eoi):
                reason = "0x626F60CE JPEG missing EOI terminator"
                print(
                    f"626F60CE HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            width = None
            height = None
            try:
                width, height = _diagnose_jpeg_dimensions_626f60cc(raw)
            except Exception:
                width, height = None, None

            size_text = (
                f"{width}x{height}"
                if width and height
                else "unknown"
            )

            print(
                f"626F60CE CHECK: {resource_id} | Group={group_id} | "
                f"Format=JPEG | Bytes={byte_count:,} | Size={size_text} | Status=OK"
            )
            return True, ""

        # Permanent BMP compatibility route.
        if raw.startswith(bmp_sig):
            ok, bmp_reason, width, height = _validate_known_bmp_thumbnail(raw)
            if not ok:
                reason = f"0x626F60CE {bmp_reason}"
                print(
                    f"626F60CE HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            print(
                f"626F60CE CHECK: {resource_id} | Group={group_id} | "
                f"Format=BMP | Bytes={byte_count:,} | "
                f"Size={width}x{height} | BPP=32 | Status=OK"
            )
            return True, ""

        reason = (
            "0x626F60CE unsupported/signature mismatch: "
            f"{raw[:16].hex().upper()}"
        )
        print(
            f"626F60CE HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(
            f"626F60CE HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_icon_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 ICON resources
    (type 0x2E75C765).

    Project evidence:
      - research notes identify 0x2E75C765 as ICON
      - five independent samples from the full resource analysis were PNG
      - deliberately truncated 8-byte PNG-signature-only control was detected
      - untouched 19,619-byte control passed as a valid 256x256 PNG with IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed dimensions, filesize, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = b"\x89PNG\r\n\x1a\n"
        jpeg_soi = b"\xFF\xD8"
        jpeg_eoi = b"\xFF\xD9"

        # Permanent JPEG compatibility route.
        if raw.startswith(jpeg_soi):
            if byte_count < 4:
                reason = f"ICON 0x2E75C765 JPEG truncated: {byte_count} bytes"
                print(
                    f"ICON HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            if not raw.endswith(jpeg_eoi):
                reason = "ICON 0x2E75C765 JPEG missing EOI terminator"
                print(
                    f"ICON HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            width = None
            height = None
            try:
                width, height = _diagnose_jpeg_dimensions_626f60cc(raw)
            except Exception:
                pass

            size_text = (
                f"{width}x{height}"
                if width and height
                else "unknown"
            )

            print(
                f"ICON CHECK: {resource_id} | Group={group_id} | "
                f"Format=JPEG | Bytes={byte_count:,} | "
                f"Size={size_text} | Status=OK"
            )
            return True, ""

        if byte_count < 24:
            reason = (
                f"ICON PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "ICON PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"ICON PNG invalid dimensions: {width}x{height}"
            print(
                f"ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == b"\x00\x00\x00\x00"
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "ICON PNG missing/invalid IEND terminator"
            print(
                f"ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(
            f"ICON HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_0580a2b5_thum_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 THUM resource type
    0x0580A2B5.

    Evidence:
      - identified as THUM
      - five independent clean samples in the full resource analysis
      - all sampled payloads are PNG
      - sampled clean examples are 64x64
      - deliberately truncated 8-byte PNG-signature-only control was detected
      - untouched 2,910-byte control passed as a valid 64x64 PNG with IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed dimensions, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = b"\x89PNG\r\n\x1a\n"
        jpeg_soi = b"\xFF\xD8"
        jpeg_eoi = b"\xFF\xD9"

        # Permanent JPEG compatibility route.
        if raw.startswith(jpeg_soi):
            if byte_count < 4:
                reason = f"THUM 0x0580A2B5 JPEG truncated: {byte_count} bytes"
                print(
                    f"0580A2B5 THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            if not raw.endswith(jpeg_eoi):
                reason = "THUM 0x0580A2B5 JPEG missing EOI terminator"
                print(
                    f"0580A2B5 THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            width = None
            height = None
            try:
                width, height = _diagnose_jpeg_dimensions_626f60cc(raw)
            except Exception:
                pass

            size_text = (
                f"{width}x{height}"
                if width and height
                else "unknown"
            )

            print(
                f"0580A2B5 THUM CHECK: {resource_id} | Group={group_id} | "
                f"Format=JPEG | Bytes={byte_count:,} | "
                f"Size={size_text} | Status=OK"
            )
            return True, ""

        if byte_count < 24:
            reason = (
                f"THUM 0x0580A2B5 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"0580A2B5 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "THUM 0x0580A2B5 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"0580A2B5 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = (
                f"THUM 0x0580A2B5 PNG invalid dimensions: "
                f"{width}x{height}"
            )
            print(
                f"0580A2B5 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == b"\x00\x00\x00\x00"
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            recoverable, checked_width, checked_height = _validate_png_missing_iend_only(raw)
            if recoverable:
                width = checked_width
                height = checked_height
                print(
                    f"0580A2B5 THUM HEALTH WARNING: {resource_id} | "
                    f"Group={group_id} | THUM 0x0580A2B5 PNG missing IEND "
                    f"terminator only; chunk/CRC/IDAT data validated; "
                    f"Status=OK (recoverable)"
                )
                return True, ""

            reason = "THUM 0x0580A2B5 PNG missing/invalid IEND terminator"
            print(
                f"0580A2B5 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(
            f"0580A2B5 THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_0580a2b4_thum_health(package, entry, resource_objects):
    """Permanent PNG health check for THUM type 0x0580A2B4."""
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)
        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = b"\x89PNG\r\n\x1a\n"
        jpeg_soi = b"\xFF\xD8"
        jpeg_eoi = b"\xFF\xD9"

        # Permanent JPEG compatibility route.
        if raw.startswith(jpeg_soi):
            if byte_count < 4:
                reason = f"THUM 0x0580A2B4 JPEG truncated: {byte_count} bytes"
                print(
                    f"0580A2B4 THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            if not raw.endswith(jpeg_eoi):
                reason = "THUM 0x0580A2B4 JPEG missing EOI terminator"
                print(
                    f"0580A2B4 THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            width = None
            height = None
            try:
                width, height = _diagnose_jpeg_dimensions_626f60cc(raw)
            except Exception:
                pass

            size_text = (
                f"{width}x{height}"
                if width and height
                else "unknown"
            )

            print(
                f"0580A2B4 THUM CHECK: {resource_id} | Group={group_id} | "
                f"Format=JPEG | Bytes={byte_count:,} | "
                f"Size={size_text} | Status=OK"
            )
            return True, ""

        if byte_count < 24:
            reason = (
                f"THUM 0x0580A2B4 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"0580A2B4 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "THUM 0x0580A2B4 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"0580A2B4 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"THUM 0x0580A2B4 PNG invalid dimensions: {width}x{height}"
            print(
                f"0580A2B4 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == b"\x00\x00\x00\x00"
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            recoverable, checked_width, checked_height = _validate_png_missing_iend_only(raw)
            if recoverable:
                width = checked_width
                height = checked_height
                print(
                    f"0580A2B4 THUM HEALTH WARNING: {resource_id} | "
                    f"Group={group_id} | THUM 0x0580A2B4 PNG missing IEND "
                    f"terminator only; chunk/CRC/IDAT data validated; "
                    f"Status=OK (recoverable)"
                )
                return True, ""

            reason = "THUM 0x0580A2B4 PNG missing/invalid IEND terminator"
            print(
                f"0580A2B4 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(
            f"0580A2B4 THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_0580a2b6_thum_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 THUM resource type
    0x0580A2B6.

    Evidence:
      - identified as THUM
      - many clean controls in ICOLOPOLIS_GrandeLama_newCC.package
      - observed clean payloads are PNG
      - sampled clean payloads are 128x128
      - deliberately truncated 8-byte PNG-signature-only replacement
        was uniquely detected while surrounding B6 THUMs remained valid

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 128x128 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])
        jpeg_soi = b"\xFF\xD8"
        jpeg_eoi = b"\xFF\xD9"

        # Permanent JPEG compatibility route.
        if raw.startswith(jpeg_soi):
            if byte_count < 4:
                reason = f"THUM 0x0580A2B6 JPEG truncated: {byte_count} bytes"
                print(
                    f"0580A2B6 THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            if not raw.endswith(jpeg_eoi):
                reason = "THUM 0x0580A2B6 JPEG missing EOI terminator"
                print(
                    f"0580A2B6 THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            width = None
            height = None
            try:
                width, height = _diagnose_jpeg_dimensions_626f60cc(raw)
            except Exception:
                pass

            size_text = (
                f"{width}x{height}"
                if width and height
                else "unknown"
            )

            print(
                f"0580A2B6 THUM CHECK: {resource_id} | Group={group_id} | "
                f"Format=JPEG | Bytes={byte_count:,} | "
                f"Size={size_text} | Status=OK"
            )
            return True, ""

        if byte_count < 24:
            reason = (
                f"THUM 0x0580A2B6 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"0580A2B6 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "THUM 0x0580A2B6 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"0580A2B6 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"THUM 0x0580A2B6 PNG invalid dimensions: {width}x{height}"
            print(
                f"0580A2B6 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            recoverable, checked_width, checked_height = _validate_png_missing_iend_only(raw)
            if recoverable:
                width = checked_width
                height = checked_height
                print(
                    f"0580A2B6 THUM HEALTH WARNING: {resource_id} | "
                    f"Group={group_id} | THUM 0x0580A2B6 PNG missing IEND "
                    f"terminator only; chunk/CRC/IDAT data validated; "
                    f"Status=OK (recoverable)"
                )
                return True, ""

            reason = "THUM 0x0580A2B6 PNG missing/invalid IEND terminator"
            print(
                f"0580A2B6 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"0580A2B6 THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_0580a2cd_snap_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 SNAP resource type 0x0580A2CD.

    Evidence:
      - confirmed SNAP type
      - clean sampled payloads are PNG
      - sampled clean examples are 64x64
      - deliberately truncated 8-byte PNG-signature-only replacement was detected
      - untouched 6,060-byte control parsed as 64x64 PNG with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 64x64 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"SNAP 0x0580A2CD PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"0580A2CD SNAP HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "SNAP 0x0580A2CD PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"0580A2CD SNAP HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"SNAP 0x0580A2CD PNG invalid dimensions: {width}x{height}"
            print(
                f"0580A2CD SNAP HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "SNAP 0x0580A2CD PNG missing/invalid IEND terminator"
            print(
                f"0580A2CD SNAP HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"0580A2CD SNAP HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_0580a2cf_snap_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 SNAP resource type 0x0580A2CF.

    Evidence:
      - confirmed SNAP type
      - clean sampled payloads are PNG
      - sampled clean examples are 512x512
      - deliberately truncated 8-byte PNG-signature-only replacement was detected
      - another 0x0580A2CF entry with the same instance but a different group
        remained valid, confirming entries must be checked independently

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 512x512 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"SNAP 0x0580A2CF PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"0580A2CF SNAP HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "SNAP 0x0580A2CF PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"0580A2CF SNAP HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"SNAP 0x0580A2CF PNG invalid dimensions: {width}x{height}"
            print(
                f"0580A2CF SNAP HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "SNAP 0x0580A2CF PNG missing/invalid IEND terminator"
            print(
                f"0580A2CF SNAP HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"0580A2CF SNAP HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_2e75c764_icon_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 ICON resource type 0x2E75C764.

    Evidence:
      - confirmed ICON type (small object thumbnail)
      - multiple project samples are DefaultResource PNGs
      - sampled clean examples are 32x32
      - deliberately truncated 8-byte PNG-signature-only replacement was detected
      - untouched 1,353-byte control passed as valid PNG with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 32x32 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"ICON 0x2E75C764 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"2E75C764 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "ICON 0x2E75C764 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"2E75C764 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"ICON 0x2E75C764 PNG invalid dimensions: {width}x{height}"
            print(
                f"2E75C764 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "ICON 0x2E75C764 PNG missing/invalid IEND terminator"
            print(
                f"2E75C764 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"2E75C764 ICON HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_2e75c766_icon_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 ICON resource type 0x2E75C766.

    Evidence:
      - confirmed ICON family (large object thumbnail)
      - multiple project samples are DefaultResource PNGs
      - sampled clean examples are 128x128
      - deliberately truncated 8-byte PNG-signature-only replacement was detected
      - untouched 17,041-byte control passed as valid PNG with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 128x128 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"ICON 0x2E75C766 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"2E75C766 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "ICON 0x2E75C766 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"2E75C766 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"ICON 0x2E75C766 PNG invalid dimensions: {width}x{height}"
            print(
                f"2E75C766 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "ICON 0x2E75C766 PNG missing/invalid IEND terminator"
            print(
                f"2E75C766 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"2E75C766 ICON HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_2e75c767_icon_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 ICON resource type 0x2E75C767.

    Evidence:
      - confirmed ICON family (very large object thumbnail)
      - project samples are DefaultResource PNGs
      - sampled clean example is 256x256
      - deliberately truncated 8-byte PNG-signature-only replacement was detected
      - untouched 102,811-byte control passed as valid PNG with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed dimensions, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"ICON 0x2E75C767 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"2E75C767 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "ICON 0x2E75C767 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"2E75C767 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"ICON 0x2E75C767 PNG invalid dimensions: {width}x{height}"
            print(
                f"2E75C767 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "ICON 0x2E75C767 PNG missing/invalid IEND terminator"
            print(
                f"2E75C767 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"2E75C767 ICON HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def _diagnose_jpeg_dimensions_626f60cc(raw):
    """Best-effort JPEG SOF dimension reader for diagnostic output only."""
    try:
        i = 2
        n = len(raw)
        sof_markers = {
            0xC0, 0xC1, 0xC2, 0xC3,
            0xC5, 0xC6, 0xC7,
            0xC9, 0xCA, 0xCB,
            0xCD, 0xCE, 0xCF,
        }

        while i + 1 < n:
            if raw[i] != 0xFF:
                i += 1
                continue

            while i < n and raw[i] == 0xFF:
                i += 1
            if i >= n:
                break

            marker = raw[i]
            i += 1

            if marker in {0x01, 0xD8, 0xD9} or 0xD0 <= marker <= 0xD7:
                continue

            if i + 1 >= n:
                break

            seg_len = (raw[i] << 8) | raw[i + 1]
            if seg_len < 2 or i + seg_len > n:
                break

            if marker in sof_markers and seg_len >= 7:
                height = (raw[i + 3] << 8) | raw[i + 4]
                width = (raw[i + 5] << 8) | raw[i + 6]
                return width, height

            i += seg_len

    except Exception:
        pass

    return None, None


def check_626f60cc_thum_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 THUM type 0x626F60CC.

    Runtime-confirmed project evidence:
      - valid 0x626F60CC resources can be PNG OR JPEG
      - truncated 8-byte PNG replacement was detected
      - untouched 109,948-byte PNG control passed (256x256, terminal IEND)
      - truncated 2-byte JPEG SOI replacement was detected
      - untouched 72,239-byte JPEG control passed (256x256, terminal EOI)

    PNG permanent rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width/height
      - terminal IEND

    JPEG permanent rules:
      - starts FF D8 (SOI)
      - at least 4 bytes
      - ends FF D9 (EOI)
      - SOF dimensions are read only as best-effort metadata and are NOT
        required, avoiding false positives on valid JPEG variants

    Any payload that is neither PNG nor JPEG is a health failure.
    No fixed dimensions, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])
        jpeg_soi = bytes([0xFF, 0xD8])
        jpeg_eoi = bytes([0xFF, 0xD9])

        # PNG route
        if raw.startswith(png_sig):
            if byte_count < 24:
                reason = (
                    f"THUM 0x626F60CC PNG truncated: {byte_count} bytes "
                    f"(too short for PNG structure)"
                )
                print(
                    f"626F60CC THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            width = int.from_bytes(raw[16:20], "big", signed=False)
            height = int.from_bytes(raw[20:24], "big", signed=False)

            if width <= 0 or height <= 0:
                reason = (
                    f"THUM 0x626F60CC PNG invalid dimensions: "
                    f"{width}x{height}"
                )
                print(
                    f"626F60CC THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            iend_ok = (
                byte_count >= 12
                and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
                and raw[-8:-4] == b"IEND"
            )

            if not iend_ok:
                reason = "THUM 0x626F60CC PNG missing/invalid IEND terminator"
                print(
                    f"626F60CC THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            return True, ""

        # JPEG route
        if raw.startswith(jpeg_soi):
            if byte_count < 4:
                reason = (
                    f"THUM 0x626F60CC JPEG truncated: {byte_count} bytes "
                    f"(too short for JPEG structure)"
                )
                print(
                    f"626F60CC THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            if not raw.endswith(jpeg_eoi):
                reason = "THUM 0x626F60CC JPEG missing EOI terminator"
                print(
                    f"626F60CC THUM HEALTH FAILURE: {resource_id} | "
                    f"Group={group_id} | {reason}"
                )
                return False, reason

            return True, ""

        reason = (
            "THUM 0x626F60CC unsupported/signature mismatch: "
            f"{raw[:16].hex().upper()}"
        )
        print(
            f"626F60CC THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"626F60CC THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_05b1b524_thum_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 THUM type 0x05B1B524.

    Identified as the small Terrain Paint Brush catalog thumbnail.

    Runtime-confirmed project evidence:
      - clean resource: 1,695-byte DefaultResource PNG, observed 32x32
      - deliberately truncated 8-byte PNG-signature-only replacement detected
      - clean control remained valid with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 32x32 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"THUM 0x05B1B524 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"05B1B524 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "THUM 0x05B1B524 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"05B1B524 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"THUM 0x05B1B524 PNG invalid dimensions: {width}x{height}"
            print(
                f"05B1B524 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "THUM 0x05B1B524 PNG missing/invalid IEND terminator"
            print(
                f"05B1B524 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"05B1B524 THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_05b1b525_thum_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 THUM type 0x05B1B525.

    Identified as the medium Terrain Paint Brush catalog thumbnail.

    Runtime-confirmed project evidence:
      - clean control: 4,157-byte DefaultResource PNG, observed 54x54
      - other clean project samples include 64x64
      - deliberately truncated 8-byte PNG-signature-only replacement detected
      - untouched control passed with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed dimensions, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"THUM 0x05B1B525 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"05B1B525 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "THUM 0x05B1B525 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"05B1B525 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"THUM 0x05B1B525 PNG invalid dimensions: {width}x{height}"
            print(
                f"05B1B525 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "THUM 0x05B1B525 PNG missing/invalid IEND terminator"
            print(
                f"05B1B525 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"05B1B525 THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_05b1b526_thum_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 THUM type 0x05B1B526.

    Identified as the large Terrain Paint Brush catalog thumbnail.

    Runtime-confirmed project evidence:
      - clean control: 24,068-byte DefaultResource PNG, observed 128x128
      - deliberately truncated 8-byte PNG-signature-only replacement detected
      - untouched control passed with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 128x128 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"THUM 0x05B1B526 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"05B1B526 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "THUM 0x05B1B526 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"05B1B526 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"THUM 0x05B1B526 PNG invalid dimensions: {width}x{height}"
            print(
                f"05B1B526 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "THUM 0x05B1B526 PNG missing/invalid IEND terminator"
            print(
                f"05B1B526 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"05B1B526 THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_2653e3c9_thum_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 THUM type 0x2653E3C9.

    Identified as a Fence catalog thumbnail type.

    Runtime-confirmed project evidence:
      - clean control: 5,381-byte DefaultResource PNG, observed 54x54
      - deliberately truncated 8-byte PNG-signature-only replacement detected
      - untouched control passed with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 54x54 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"THUM 0x2653E3C9 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"2653E3C9 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "THUM 0x2653E3C9 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"2653E3C9 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"THUM 0x2653E3C9 PNG invalid dimensions: {width}x{height}"
            print(
                f"2653E3C9 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "THUM 0x2653E3C9 PNG missing/invalid IEND terminator"
            print(
                f"2653E3C9 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"2653E3C9 THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_fceab65b_thum_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 THUM type 0xFCEAB65B.

    Identified as the thumbnail for a packaged custom colourswatch.

    Runtime-confirmed project evidence:
      - clean control: 78,745-byte DefaultResource PNG, observed 256x256
      - four additional clean project samples are also PNG
      - deliberately truncated 8-byte PNG-signature-only replacement detected
      - untouched control passed with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 256x256 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"THUM 0xFCEAB65B PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"FCEAB65B THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "THUM 0xFCEAB65B PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"FCEAB65B THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"THUM 0xFCEAB65B PNG invalid dimensions: {width}x{height}"
            print(
                f"FCEAB65B THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "THUM 0xFCEAB65B PNG missing/invalid IEND terminator"
            print(
                f"FCEAB65B THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"FCEAB65B THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_d84e7fc5_icon_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 ICON type 0xD84E7FC5.

    Identified as the small lot image / lot thumbnail ICON.

    Runtime-confirmed project evidence:
      - clean control: 22,020-byte DefaultResource PNG, observed 128x128
      - deliberately truncated 8-byte PNG-signature-only replacement detected
      - untouched control passed with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 128x128 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"ICON 0xD84E7FC5 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"D84E7FC5 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "ICON 0xD84E7FC5 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"D84E7FC5 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"ICON 0xD84E7FC5 PNG invalid dimensions: {width}x{height}"
            print(
                f"D84E7FC5 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "ICON 0xD84E7FC5 PNG missing/invalid IEND terminator"
            print(
                f"D84E7FC5 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"D84E7FC5 ICON HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_d84e7fc6_icon_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 ICON type 0xD84E7FC6.

    Identified as the medium lot image / lot thumbnail ICON.

    Runtime-confirmed project evidence:
      - clean control: 80,247-byte DefaultResource PNG, observed 256x256
      - deliberately truncated 8-byte PNG-signature-only replacement detected
      - untouched control passed with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 256x256 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"ICON 0xD84E7FC6 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"D84E7FC6 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "ICON 0xD84E7FC6 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"D84E7FC6 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"ICON 0xD84E7FC6 PNG invalid dimensions: {width}x{height}"
            print(
                f"D84E7FC6 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "ICON 0xD84E7FC6 PNG missing/invalid IEND terminator"
            print(
                f"D84E7FC6 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"D84E7FC6 ICON HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_d84e7fc7_icon_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 ICON type 0xD84E7FC7.

    Identified as the large lot image / lot thumbnail ICON.

    Runtime-confirmed project evidence:
      - damaged group 0x00000000 replacement: 8-byte PNG signature only -> detected
      - untouched group 0x00000000 control: 188,854-byte PNG, 512x512 -> passed
      - same instance also runtime-confirmed clean across groups 0x00000001..0x0000000A
      - all observed clean resources are DefaultResource PNGs with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 512x512 requirement, filesize, group, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"ICON 0xD84E7FC7 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"D84E7FC7 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "ICON 0xD84E7FC7 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"D84E7FC7 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"ICON 0xD84E7FC7 PNG invalid dimensions: {width}x{height}"
            print(
                f"D84E7FC7 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "ICON 0xD84E7FC7 PNG missing/invalid IEND terminator"
            print(
                f"D84E7FC7 ICON HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"D84E7FC7 ICON HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_twni_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 TWNI type 0x0668F635.

    TWNI = World Description Image / World Metadata Image.

    Runtime-confirmed project evidence:
      - clean control: 170,892-byte DefaultResource PNG, observed 256x256
      - deliberately truncated 8-byte PNG-signature-only replacement detected
      - untouched control passed with terminal IEND
      - control world contained 6,048 total resources without unrelated failures

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 256x256 requirement, filesize, group, instance, package size,
    or world resource-count assumption.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"TWNI PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"TWNI HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "TWNI PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"TWNI HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"TWNI PNG invalid dimensions: {width}x{height}"
            print(
                f"TWNI HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "TWNI PNG missing/invalid IEND terminator"
            print(
                f"TWNI HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"TWNI HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_2d4284f1_thum_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 THUM type 0x2D4284F1.

    Identified as the medium Railing catalog thumbnail.

    Runtime-confirmed project evidence:
      - clean control: 94,943-byte DefaultResource PNG, observed 256x256
      - second independent clean package carried the same resource
      - deliberately truncated 8-byte PNG-signature-only replacement detected
      - untouched control passed with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No fixed 256x256 requirement, filesize, group, instance,
    colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"THUM 0x2D4284F1 PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"2D4284F1 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "THUM 0x2D4284F1 PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"2D4284F1 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"THUM 0x2D4284F1 PNG invalid dimensions: {width}x{height}"
            print(
                f"2D4284F1 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "THUM 0x2D4284F1 PNG missing/invalid IEND terminator"
            print(
                f"2D4284F1 THUM HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"2D4284F1 THUM HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_manifest_73e93eeb_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 package-manifest type 0x73E93EEB.

    Runtime-confirmed project evidence:
      - multiple independent clean samples use DefaultResource and XML payloads
      - clean control: 1,427 bytes, UTF-8, root <manifest>
      - deliberately malformed/truncated 47-byte XML replacement failed parsing
      - untouched clean control parsed successfully

    Permanent rules are intentionally conservative:
      - resource must not be empty
      - decode BOM-aware UTF-8 / UTF-16 text
      - XML must parse successfully

    No fixed root element, schema, filesize, group, instance, encoding,
    child-count, or attribute assumptions.
    """
    import xml.etree.ElementTree as ET

    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        if byte_count == 0:
            reason = "MANIFEST XML empty: 0 bytes"
            print(
                f"73E93EEB MANIFEST HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        try:
            if raw.startswith(bytes([0xFF, 0xFE])):
                decoded = raw.decode("utf-16-le")
            elif raw.startswith(bytes([0xFE, 0xFF])):
                decoded = raw.decode("utf-16-be")
            else:
                decoded = raw.decode("utf-8-sig")
        except Exception as ex:
            reason = f"MANIFEST XML decode failed: {type(ex).__name__}: {ex}"
            print(
                f"73E93EEB MANIFEST HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        try:
            ET.fromstring(decoded)
        except Exception as ex:
            reason = f"MANIFEST XML parse failed: {type(ex).__name__}: {ex}"
            print(
                f"73E93EEB MANIFEST HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"73E93EEB MANIFEST HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_8f78f74b_png_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 resource type 0x8F78F74B.

    Runtime-confirmed project evidence:
      - real clean sample uses s3pi.DefaultResource.DefaultResource
      - clean control: 153,079-byte PNG, observed 256x256
      - deliberately truncated 8-byte PNG-signature-only replacement detected
      - untouched control passed with terminal IEND

    Permanent PNG structural rules:
      - at least 24 bytes
      - standard PNG signature
      - non-zero width and height
      - valid terminal IEND marker

    No semantic family is assumed yet. No fixed 256x256 requirement,
    filesize, group, instance, colour type, or compression ratio.
    """
    resource_id = f"0x{int(entry.Instance):016X}"
    group_id = f"0x{int(entry.ResourceGroup):08X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)

        raw = bytes(resource.AsBytes)
        byte_count = len(raw)
        png_sig = bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])

        if byte_count < 24:
            reason = (
                f"0x8F78F74B PNG truncated: {byte_count} bytes "
                f"(too short for PNG structure)"
            )
            print(
                f"8F78F74B PNG HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        if not raw.startswith(png_sig):
            reason = (
                "0x8F78F74B PNG signature mismatch: "
                f"{raw[:16].hex().upper()}"
            )
            print(
                f"8F78F74B PNG HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        width = int.from_bytes(raw[16:20], "big", signed=False)
        height = int.from_bytes(raw[20:24], "big", signed=False)

        if width <= 0 or height <= 0:
            reason = f"0x8F78F74B PNG invalid dimensions: {width}x{height}"
            print(
                f"8F78F74B PNG HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        iend_ok = (
            byte_count >= 12
            and raw[-12:-8] == bytes([0x00, 0x00, 0x00, 0x00])
            and raw[-8:-4] == b"IEND"
        )

        if not iend_ok:
            reason = "0x8F78F74B PNG missing/invalid IEND terminator"
            print(
                f"8F78F74B PNG HEALTH FAILURE: {resource_id} | "
                f"Group={group_id} | {reason}"
            )
            return False, reason

        return True, ""

    except Exception as ex:
        reason = f"{type(ex).__name__}: {ex}"
        print(
            f"8F78F74B PNG HEALTH FAILURE: {resource_id} | "
            f"Group={group_id} | {reason}"
        )
        return False, reason


def check_simo_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 SIMO / Sim Outfit resources.

    SIMO has a dedicated S3PI parser. A deserialization exception is treated
    as structural corruption. No fixed version or byte-size requirement is
    imposed because valid SIMO resources vary substantially in size.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"SIMO HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_txtf_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 TXTF / Fabric Compositor.

    TXTF is handled by the S3PI compositor parser. Any deserialization
    exception is treated as structural corruption.

    No fixed version or byte-size requirement is imposed because valid TXTF
    resources vary substantially in size.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"TXTF HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_txtc_health(package, entry, resource_objects):
    """
    Permanent structural health check for Sims 3 TXTC / Texture Compositor.

    Important compatibility safeguard:
    S3PI's TXTC wrapper does not understand every property used by every
    custom resource. An "Unexpected property ID ... enumType ..." exception
    is therefore treated as an unsupported-wrapper warning, NOT definitive
    corruption.

    Clear structural parser failures such as EndOfStream/truncation still
    mark the TXTC as BROKEN.
    """
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        reason_lower = reason.lower()

        # Narrow compatibility exception confirmed from a real untouched
        # package and reproduced directly in S3PE/S3PI:
        # property 0x64399EC5 rejected for enumType SingleProperties.
        #
        # Do NOT suppress arbitrary "Unexpected property ID" errors, because
        # those could still represent genuinely damaged TXTC data.
        known_wrapper_limitation = (
            "unexpected property id 0x64399ec5" in reason_lower
            and "enumtype singleproperties" in reason_lower
        )

        if known_wrapper_limitation:
            print(
                f"TXTC COMPATIBILITY WARNING: {resource_id} | "
                f"S3PI/S3PE cannot interpret property 0x64399EC5 "
                f"for SingleProperties; not classifying as BROKEN | {reason}"
            )
            return True, ""

        print(f"TXTC HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_skin_tone_health(package, entry, resource_objects):
    """Conservative structural health check for Sims 3 TONE/skintone resources."""
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(0, package, entry)
        resource_objects.append(resource)
        raw = bytes(resource.AsBytes)
        byte_count = len(raw)

        if byte_count < 16:
            reason = f"TONE truncated: {byte_count} bytes (too short for minimum structure)"
            print(f"TONE HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if raw[:12] == b"\x00" * 12:
            reason = "TONE invalid zero header"
            print(f"TONE HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        version = int.from_bytes(raw[0:4], "little", signed=False)
        tgi_offset = int.from_bytes(raw[4:8], "little", signed=False)
        tgi_size = int.from_bytes(raw[8:12], "little", signed=False)
        tgi_start = 8 + tgi_offset
        tgi_end = tgi_start + tgi_size

        if tgi_start > byte_count:
            reason = f"TONE invalid TGI offset: start {tgi_start} exceeds resource size {byte_count} (version={version})"
            print(f"TONE HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        if tgi_end > byte_count:
            reason = f"TONE truncated TGI block: end {tgi_end} exceeds resource size {byte_count} (version={version})"
            print(f"TONE HEALTH FAILURE: {resource_id} | {reason}")
            return False, reason

        return True, ""

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        print(f"TONE HEALTH FAILURE: {resource_id} | {reason}")
        return False, reason


def check_build_buy_mesh_health(package, entry, resource_objects, resource_label):
    """Return structural parse health for MODL/MLOD resources."""
    resource_id = f"0x{int(entry.Instance):016X}"

    try:
        resource = WrapperDealer.GetResource(
            0,
            package,
            entry
        )
        resource_objects.append(resource)
        return True, ""

    except Exception as e:
        error = f"{type(e).__name__}: {e}"
        print(
            f"{resource_label} HEALTH FAILURE: {resource_id} | {error}"
        )
        return False, error

def check_package_storage_health(package, resources, filename):
    """
    Permanent DBPF/package-storage corruption check.

    Resolve the concrete:
      s3pi.Package.Package.GetResource(IResourceIndexEntry)

    through .NET reflection, because Python sees the package through
    the IPackage interface.

    This tests the PACKAGE STORAGE / DECOMPRESSION layer before any
    resource-specific wrapper/parser.

    Confirmed corruption signatures include:
      - InvalidDataException: unexpected end of file
      - OverflowException from impossible decompression sizes
      - resource data size does not match the package index

    Conservative safeguards:
      - zero-byte resources alone are allowed
      - no resource type / instance pattern is assumed
      - no fixed filesize, memsize, or compression ratio is assumed
      - if GetResource cannot be resolved, this check is skipped rather
        than falsely marking the package broken
    """
    import System
    from System.Reflection import BindingFlags

    failures = []
    failed_ids = []

    package_type = package.GetType()
    get_resource_method = None
    fallback_method = None

    methods = package_type.GetMethods(
        BindingFlags.Instance |
        BindingFlags.Public |
        BindingFlags.NonPublic
    )

    for method in methods:
        try:
            if method.Name != "GetResource":
                continue

            params = method.GetParameters()

            if len(params) != 1:
                continue

            if fallback_method is None:
                fallback_method = method

            param_type_name = str(
                params[0].ParameterType.FullName or ""
            )

            if "IResourceIndexEntry" in param_type_name:
                get_resource_method = method
                break

        except Exception:
            continue

    if get_resource_method is None:
        get_resource_method = fallback_method

    if get_resource_method is None:
        print(
            f"PACKAGE STORAGE CHECK UNAVAILABLE: "
            f"could not resolve GetResource on {package_type.FullName}"
        )
        return failures, failed_ids

    for i in range(resources.Count):
        entry = resources[i]
        stream = None

        try:
            args = System.Array[System.Object]([entry])
            stream = get_resource_method.Invoke(package, args)

            # Invocation performs the DBPF read/decompression.
            # ReadByte() is harmless for a valid zero-byte stream (-1).
            if stream is not None and stream.CanRead:
                _ = stream.ReadByte()

        except Exception as e:

            actual = getattr(e, "InnerException", None)

            if actual is None:
                actual = e

            rid = f"0x{int(entry.Instance):016X}"
            rtype = f"0x{int(entry.ResourceType):08X}"

            try:
                error_name = str(actual.GetType().Name)
            except Exception:
                error_name = type(actual).__name__

            reason = f"{error_name}: {actual}"
            id_text = f"{rtype}/{rid}"

            failures.append(
                f"PACKAGE STORAGE {id_text}: {reason}"
            )

            failed_ids.append(id_text)

            print(
                f"PACKAGE STORAGE HEALTH FAILURE: "
                f"Type={rtype} | Instance={rid} | {reason}"
            )

        finally:
            try:
                if stream is not None:
                    stream.Dispose()
            except Exception:
                try:
                    if stream is not None:
                        stream.Close()
                except Exception:
                    pass

    if failures:
        print(
            f"PACKAGE STORAGE HEALTH SUMMARY: "
            f"{os.path.basename(filename)} | "
            f"BROKEN_ENTRIES={len(failures)} | "
            f"TOTAL={resources.Count}"
        )

    return failures, failed_ids



def _health_issue_instance_and_text(raw_error):
    """Split scanner health error into optional instance ID and raw exception text."""
    text = str(raw_error or "").strip()
    match = re.match(r"^(0x[0-9A-Fa-f]{8,16})\s*:\s*(.*)$", text, re.S)
    if match:
        return "0x" + match.group(1)[2:].upper(), match.group(2).strip()
    return "", text


def _condense_health_exception(raw_error, grouped=False):
    """Return a compact human-readable reason while raw detail stays in the console log."""
    text = str(raw_error or "").strip()
    if not text:
        return "malformed/unreadable resource"

    # Sims 3 resource-specific messages that are much clearer than the raw
    # S3PI/.NET exception chain. These affect spreadsheet wording only.
    lite_match = re.search(
        r"Invalid Tag read:.*?expected:\s*['\"]LITE['\"]",
        text,
        re.I | re.S,
    )
    if lite_match:
        return "malformed resource: missing/invalid LITE tag"

    if re.search(r"XML empty resource", text, re.I):
        return "empty resource"

    xml_parse_match = re.search(
        r"XML parse failed:\s*([^\r\n]+)",
        text,
        re.I,
    )
    if xml_parse_match:
        detail = xml_parse_match.group(1).strip()
        mismatch = re.search(r"mismatched tag:\s*line\s*(\d+),\s*column\s*(\d+)", detail, re.I)
        if mismatch:
            return f"malformed XML: mismatched tag at line {mismatch.group(1)}, column {mismatch.group(2)}"
        junk = re.search(r"junk after document element:\s*line\s*(\d+),\s*column\s*(\d+)", detail, re.I)
        if junk:
            return f"malformed XML: extra content after root element at line {junk.group(1)}"
        unclosed = re.search(r"unclosed token:\s*line\s*(\d+),\s*column\s*(\d+)", detail, re.I)
        if unclosed:
            return f"malformed XML: unclosed token at line {unclosed.group(1)}"
        invalid = re.search(r"not well-formed \(invalid token\):\s*line\s*(\d+),\s*column\s*(\d+)", detail, re.I)
        if invalid:
            return f"malformed XML: invalid token at line {invalid.group(1)}, column {invalid.group(2)}"
        if len(detail) > 150:
            detail = detail[:147].rstrip() + "..."
        return f"malformed XML: {detail}"

    if re.search(r"Hit unexpected end of file", text, re.I):
        return "truncated/corrupt data: unexpected end of file"

    if re.search(r"Array dimensions exceeded supported range", text, re.I):
        return "invalid compressed size/resource index"

    if re.search(r"Resource data indicates size does not match index", text, re.I):
        return "compressed data size/index mismatch"

    # Common parser failures first, because outer TargetInvocationException text
    # otherwise hides the useful inner cause.
    byte_match = re.search(
        r"Expected\s+(0x[0-9A-Fa-f]+|\d+)\s+bytes?\s+of\s+data,\s*read\s+"
        r"(0x[0-9A-Fa-f]+|\d+)",
        text,
        re.I,
    )
    if byte_match:
        if "ShaderDataList" in text or ".MTNF.Parse" in text:
            if grouped:
                return "truncated ShaderData"
            return (
                "truncated ShaderData "
                f"(expected {byte_match.group(1)}, read {byte_match.group(2)})"
            )
        if grouped:
            return "truncated/malformed data"
        return (
            "truncated/malformed data "
            f"(expected {byte_match.group(1)}, read {byte_match.group(2)})"
        )

    if re.search(r"Unable to read beyond the end of the stream", text, re.I):
        return "truncated resource (unexpected end of data)"

    if re.search(r"EndOfStreamException", text, re.I):
        return "truncated resource (unexpected end of data)"

    # Preserve useful innermost parser messages but drop stack traces and
    # wrapper-only TargetInvocationException layers.
    exception_matches = re.findall(
        r"(?:System\.[A-Za-z0-9_.]+\.)?([A-Za-z0-9_]+Exception)\s*:\s*([^\r\n]+)",
        text,
    )
    meaningful = [
        (name, message.strip())
        for name, message in exception_matches
        if name not in {"TargetInvocationException"}
        and message.strip()
    ]
    if meaningful:
        name, message = meaningful[-1]
        message = message.split(" ---> ", 1)[0].strip()
        if len(message) > 180:
            message = message[:177].rstrip() + "..."
        short_name = name[:-9] if name.endswith("Exception") else name
        return f"{short_name}: {message}"

    # If there was no recognisable exception class, use only the first useful
    # line and never copy a .NET stack trace into the spreadsheet.
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("at ") or line.startswith("---"):
            continue
        line = re.sub(
            r"^(?:System\.)?(?:Reflection\.)?TargetInvocationException:\s*"
            r"Exception has been thrown by the target of an invocation\.?(?:\s*--->\s*)?",
            "",
            line,
            flags=re.I,
        ).strip()
        if line:
            if len(line) > 180:
                line = line[:177].rstrip() + "..."
            return line

    return "malformed/unreadable resource"


def _health_reason_group_key(raw_error):
    """Stable reason key used to collapse repeated failures in one resource family."""
    reason = _condense_health_exception(raw_error, grouped=True).lower()
    reason = re.sub(r"0x[0-9a-f]+", "<hex>", reason)
    reason = re.sub(r"\b\d+\b", "<n>", reason)
    return reason


def _condense_resource_health_group(resource_label, error_lines):
    """Compact one resource family's errors for the spreadsheet Health Issues cell."""
    parsed = []
    for raw in (error_lines or []):
        instance_id, error_text = _health_issue_instance_and_text(raw)
        parsed.append((instance_id, error_text))

    if not parsed:
        return []

    # One or two failures: retain the exact instance IDs because they are useful
    # and still fit comfortably in the review cell.
    if len(parsed) <= 2:
        result = []
        for instance_id, error_text in parsed:
            reason = _condense_health_exception(error_text, grouped=False)
            if instance_id:
                result.append(f"{resource_label} {instance_id} — {reason}")
            else:
                result.append(f"{resource_label} — {reason}")
        return result

    # Larger groups: collapse repeated parser failures. Exact IDs remain in the
    # dedicated Broken <TYPE> IDs technical columns.
    grouped_reasons = {}
    reason_order = []
    for _instance_id, error_text in parsed:
        key = _health_reason_group_key(error_text)
        if key not in grouped_reasons:
            grouped_reasons[key] = {
                "count": 0,
                "display": _condense_health_exception(error_text, grouped=True),
            }
            reason_order.append(key)
        grouped_reasons[key]["count"] += 1

    if len(grouped_reasons) == 1:
        item = grouped_reasons[reason_order[0]]
        noun = "resource" if len(parsed) == 1 else "resources"
        return [
            f"{resource_label} — {len(parsed)} malformed {noun}: {item['display']}"
        ]

    result = []
    for key in reason_order:
        item = grouped_reasons[key]
        noun = "resource" if item["count"] == 1 else "resources"
        result.append(
            f"{resource_label} — {item['count']} malformed {noun}: {item['display']}"
        )
    return result


def _build_condensed_health_issues(resource_error_groups):
    """Build short Health Issues lines from (label, raw_error_list) pairs."""
    output = []
    for resource_label, error_lines in resource_error_groups:
        output.extend(_condense_resource_health_group(resource_label, error_lines))
    return output


def _condense_package_health_issue(error_text):
    """Compact a package-level scan failure for spreadsheet display."""
    text = str(error_text or "").strip()
    if not text:
        return "PACKAGE — unreadable package"
    if re.search(r"Empty package file|0 bytes", text, re.I):
        return "PACKAGE — zero-byte file"
    if re.search(r"Hit unexpected end of file", text, re.I) and "CheckHeader" in text:
        return "PACKAGE — truncated/corrupt DBPF header: unexpected end of file"
    return "PACKAGE — " + _condense_health_exception(text, grouped=False)


def scan_package(filename):

    package = None
    resources = None
    resource_objects = []

    try:

        package = Package.OpenPackage(0, filename)

        # -------------------------
        # INITIAL VALUES
        # -------------------------

        img_count = 0
        valid_texture_count = 0
        fourk = 0
        two_k = 0

        content_type = ""
        cc_type = "Unknown"
        package_format = ""
        career_marker = False
        tuning_marker = False
        random_event_module_marker = False
        interaction_tuning_marker = False
        script_mod_marker = False
        external_script_class_marker = False
        external_script_class_names = []
        objk_script_class_names = []
        s3sa_defined_type_names = set()
        scripted_object_subtype = ""
        ui_layout_marker = False
        ini_mod_marker = False
        hair_color_preset_marker = False
        thumbnail_mod_candidate = True
        thumbnail_mod_resource_count = 0
        font_mod_candidate = True
        font_mod_nmap_count = 0
        font_mod_css_count = 0
        font_mod_ttf_count = 0
        lot_ldes_marker = False
        lot_preview_count = 0
        pose_candidate = True
        pose_nmap_count = 0
        pose_clip_count = 0
        pose_one_frame_clip_count = 0
        pose_multi_frame_clip_count = 0
        string_table_mod_candidate = True
        string_table_mod_stbl_count = 0
        string_table_mod_nmap_count = 0
        collection_file_candidate = True
        collection_file_count = 0
        footprint_mod_candidate = True
        footprint_mod_ftpt_count = 0
        opportunity_tuning_candidate = True
        opportunity_tuning_nmap_count = 0
        opportunity_tuning_xml_count = 0
        opportunity_tuning_valid_xml_count = 0
        jazz_mod_candidate = True
        jazz_mod_nmap_count = 0
        jazz_mod_jazz_count = 0
        academic_degree_tuning_candidate = True
        academic_degree_tuning_xml_count = 0
        academic_degree_tuning_valid_xml_count = 0
        career_tuning_candidate = True
        career_tuning_nmap_count = 0
        career_tuning_xml_count = 0
        career_tuning_valid_xml_count = 0
        book_tuning_candidate = True
        book_tuning_nmap_count = 0
        book_tuning_xml_count = 0
        book_tuning_valid_xml_count = 0
        dumpster_diving_tuning_candidate = True
        dumpster_diving_tuning_nmap_count = 0
        dumpster_diving_tuning_xml_count = 0
        dumpster_diving_tuning_valid_xml_count = 0
        dreams_and_promises_tuning_candidate = True
        dreams_and_promises_tuning_xml_count = 0
        dreams_and_promises_tuning_valid_xml_count = 0
        scoring_tuning_candidate = True
        scoring_tuning_nmap_count = 0
        scoring_tuning_xml_count = 0
        scoring_tuning_valid_xml_count = 0
        boarding_school_tuning_candidate = True
        boarding_school_tuning_nmap_count = 0
        boarding_school_tuning_xml_count = 0
        boarding_school_tuning_valid_xml_count = 0
        buff_tuning_candidate = True
        buff_tuning_nmap_count = 0
        buff_tuning_xml_count = 0
        buff_tuning_valid_xml_count = 0
        social_tuning_candidate = True
        social_tuning_nmap_count = 0
        social_tuning_xml_count = 0
        social_tuning_valid_xml_count = 0
        skill_tuning_candidate = True
        skill_tuning_nmap_count = 0
        skill_tuning_xml_count = 0
        skill_tuning_valid_xml_count = 0
        texture_mod_candidate = True
        texture_mod_nmap_count = 0
        texture_mod_img_count = 0
        fishing_tuning_candidate = True
        fishing_tuning_nmap_count = 0
        fishing_tuning_xml_count = 0
        fishing_tuning_stbl_count = 0
        fishing_tuning_valid_xml_count = 0
        cooking_tuning_candidate = True
        cooking_tuning_nmap_count = 0
        cooking_tuning_xml_count = 0
        cooking_tuning_stbl_count = 0
        cooking_tuning_png_count = 0
        cooking_tuning_store_ingredients_count = 0
        cooking_tuning_valid_xml_count = 0
        cooking_tuning_valid_store_ingredients_count = 0
        cooking_tuning_recipe_process_count = 0
        cooking_tuning_recipe_master_count = 0
        cooking_tuning_ingredients_count = 0
        cooking_tuning_plants_count = 0
        cooking_tuning_bars_count = 0
        cooking_tuning_shop_tabs_count = 0
        terrain_paint_candidate = True
        terrain_paint_ctpt_count = 0
        terrain_paint_stbl_count = 0
        terrain_paint_img_count = 0
        terrain_paint_icon_count = 0
        terrain_paint_nmap_count = 0
        terrain_paint_brush_thum_count = 0

        biggest = 0
        biggest_size = ""

        casp = 0
        geom = 0
        presets = 0
        preset_type = ""
        simo = 0
        sime = 0
        famd = 0
        objd = 0
        objk = 0
        modl = 0
        mlod = 0

        # EA reference counters:
        # ea_matches = exact TGI + content-hash matches
        # ea_tgi_matches = known EA TGI regardless of payload
        # ea_modified_matches = known EA TGI with different payload
        ea_matches = 0
        ea_tgi_matches = 0
        ea_modified_matches = 0
        ea_unverifiable_matches = 0
        # Temporary V4.2 diagnostics: record why EA Candidate packages are mixed.
        ea_exact_details = []
        ea_modified_details = []
        ea_unverifiable_details = []
        ea_non_ea_type_counts = {}
        ea_hash_errors = 0
        ea_casp_matches = 0
        ea_geom_matches = 0
        ea_img_matches = 0

        geom_total = 0
        geom_largest = 0

        # Structural resource health is separate from performance risk.
        broken_geom_ids = []
        broken_geom_errors = []
        broken_casp_ids = []
        broken_casp_errors = []
        broken_img_ids = []
        broken_img_errors = []
        broken_vpxy_ids = []
        broken_vpxy_errors = []
        broken_bone_ids = []
        broken_bone_errors = []
        broken_bond_ids = []
        broken_bond_errors = []
        broken_morph_ids = []
        broken_morph_errors = []
        broken_build_buy_ids = []
        broken_build_buy_errors = []
        broken_preset_ids = []
        broken_preset_errors = []
        broken_tone_ids = []
        broken_tone_errors = []
        broken_imag_ids = []
        broken_imag_errors = []
        broken_txtc_ids = []
        broken_txtc_errors = []
        broken_ptrn_ids = []
        broken_ptrn_errors = []
        broken_simo_ids = []
        broken_simo_errors = []
        broken_txtf_ids = []
        broken_txtf_errors = []
        broken_famd_ids = []
        broken_famd_errors = []
        broken_objd_ids = []
        broken_objk_ids = []
        broken_xml_ids = []
        broken_xml_errors = []
        broken_stbl_ids = []
        broken_stbl_errors = []
        broken_nmap_ids = []
        broken_nmap_errors = []
        broken_rslt_ids = []
        broken_ftpt_ids = []
        broken_rig_ids = []
        broken_rig_errors = []
        broken_jazz_ids = []
        broken_jazz_errors = []
        broken_s3sa_ids = []
        broken_s3sa_errors = []
        broken_lite_ids = []
        broken_lite_errors = []
        broken_clip_ids = []
        broken_clip_errors = []
        broken_matd_ids = []
        broken_mtst_ids = []
        broken_mdlr_ids = []
        broken_cwal_ids = []
        broken_cfen_ids = []
        broken_ctpt_ids = []
        broken_ngmp_ids = []
        broken_ngmp_errors = []
        broken_crmt_ids = []
        broken_spt2_ids = []
        broken_spt_ids = []
        broken_cral_ids = []
        broken_layo_ids = []
        broken_layo_errors = []
        broken_hinf_ids = []
        broken_hinf_errors = []
        broken_2ary_ids = []
        broken_2ary_errors = []
        broken_complate_ids = []
        broken_complate_errors = []
        broken_63a33ea7_ids = []
        broken_63a33ea7_errors = []
        broken_fireplace_ids = []
        broken_fireplace_errors = []
        broken_thum_ids = []
        broken_thum_errors = []
        broken_626f60ce_ids = []
        broken_626f60ce_errors = []
        broken_icon_ids = []
        broken_icon_errors = []
        broken_twni_ids = []
        broken_twni_errors = []
        broken_manifest_ids = []
        broken_manifest_errors = []
        broken_8f78f74b_ids = []
        broken_8f78f74b_errors = []
        broken_0580a2b5_ids = []
        broken_0580a2b5_errors = []
        broken_0580a2b4_ids = []
        broken_0580a2b4_errors = []
        broken_0580a2b6_ids = []
        broken_0580a2b6_errors = []
        broken_0580a2cd_ids = []
        broken_0580a2cd_errors = []
        broken_0580a2cf_ids = []
        broken_0580a2cf_errors = []

        vpxy = 0
        vpxy_instances = []

        bones = 0

        # CAS SLIDER RESOURCES
        face = 0
        fbln = 0
        stbl_instances = []
        bgeo_instances = []
        face_instances = []
        fbln_instances = []
        key_instances = []

        total_vertices = 0
        highest_polys = 0
        highest_geom_id = ""

        poly_warning = "Normal"

        resolutions = {}
        img_sizes = []
        casp_instances = []
        casp_records = []
        geom_instances = []
        img_instances = []
        unknown_resources = []
        unknown_reason = ""
        slider_count = 0

        # -------------------------
        # GET RESOURCES
        # -------------------------

        resources = package.GetResourceList

        # -------------------------
        # PACKAGE STORAGE / DECOMPRESSION HEALTH
        # -------------------------

        (
            broken_package_storage_errors,
            broken_package_storage_ids
        ) = check_package_storage_health(
            package,
            resources,
            filename
        )

        # -------------------------
        # SCANNING STARTS HERE
        # -------------------------

        known_types = [
            IMG_TYPE,
            IMAG_JPEG_TYPE,
            IMAG_PNG_TYPE,
            IMAG_TGA_TYPE,
            TONE_TYPE,
            GEOM_TYPE,
            VPXY_TYPE,
            BONE_TYPE,
            BOND_TYPE,
            PRESET_TYPE,
            STBL_TYPE,
            COLL_TYPE,
            SIMO_TYPE,
            SIME_TYPE,
            FAMD_TYPE,
            KEY_TYPE,
            BGEO_TYPE,
            BBLN_TYPE,
            FACE_TYPE,
            FBLN_TYPE,
            TXTC_TYPE,
            TXTF_TYPE,
            PTRN_TYPE,
            XML_TYPE,
            CSS_TYPE,
            TTF_TYPE,
            LDES_TYPE,
            THUM_AD366F95_TYPE,
            ITUN_TYPE,
            MODL_TYPE,
            MLOD_TYPE,
            OBJK_TYPE,
            OBJD_TYPE,
            MATD_TYPE,
            MTST_TYPE,
            MDLR_TYPE,
            CWAL_TYPE,
            CFEN_TYPE,
            CTPT_TYPE,
            NGMP_TYPE,
            CRMT_TYPE,
            SPT2_TYPE,
            SPT_TYPE,
            CRAL_TYPE,
            LAYO_TYPE,
            HINF_TYPE,
            TWOARY_TYPE,
            COMPLATE_TYPE,
            TYPE_63A33EA7
        ]

        known_types.extend(CASP_TYPES)

        for i in range(resources.Count):

            if broken_package_storage_ids:
                break

            entry = resources[i]

            rtype = int(entry.ResourceType)
            resource_instance = int(entry.Instance)

            # Terrain Paint TEST structural candidate:
            # Expanded only for the validated Kale + Heaven terrain-paint
            # families. MANIFEST is optional metadata. Substantive resources
            # must remain within CTPT + STBL + IMG, with only the validated
            # NMAP / generic ICON / terrain-brush THUM auxiliaries allowed.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == CTPT_TYPE:
                terrain_paint_ctpt_count += 1
            elif rtype == STBL_TYPE:
                terrain_paint_stbl_count += 1
            elif rtype == IMG_TYPE:
                terrain_paint_img_count += 1
            elif rtype == ICON_TYPE:
                terrain_paint_icon_count += 1
            elif rtype == NMAP_TYPE:
                terrain_paint_nmap_count += 1
            elif rtype in (
                THUM_05B1B524_TYPE,
                THUM_05B1B525_TYPE,
                THUM_05B1B526_TYPE
            ):
                terrain_paint_brush_thum_count += 1
            else:
                terrain_paint_candidate = False

            # Thumbnail Mod V2 TEST structural candidate:
            # MANIFEST and NMAP are auxiliary metadata. Every other
            # substantive resource must be one of the established THUM
            # image types used by this classifier.
            if rtype in (MANIFEST_TYPE, NMAP_TYPE):
                pass
            elif rtype in (
                THUM_TYPE,
                PNG_626F60CE_TYPE,
                THUM_AD366F95_TYPE
            ):
                thumbnail_mod_resource_count += 1
            else:
                thumbnail_mod_candidate = False

            # Font Mod TEST structural candidate:
            # ignore only the standard MANIFEST auxiliary; every other
            # substantive resource must be NMAP, _CSS, or _TTF.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                font_mod_nmap_count += 1
            elif rtype == CSS_TYPE:
                font_mod_css_count += 1
            elif rtype == TTF_TYPE:
                font_mod_ttf_count += 1
            else:
                font_mod_candidate = False

            # Lot TEST supporting structural evidence:
            # Sims 3's dedicated lot preview image types.
            if rtype in (
                ICON_D84E7FC5_TYPE,
                ICON_D84E7FC6_TYPE,
                ICON_D84E7FC7_TYPE
            ):
                lot_preview_count += 1

            # Pose TEST structural candidate:
            # MANIFEST is auxiliary metadata; every substantive resource
            # must otherwise be NMAP or CLIP.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                pose_nmap_count += 1
            elif rtype == CLIP_TYPE:
                pose_clip_count += 1
            else:
                pose_candidate = False

            # String Table Mod structural candidate:
            # MANIFEST and NMAP are auxiliary metadata; every other
            # substantive resource must be STBL.  NMAP commonly names the
            # localized Strings_* tables and does not change the category.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                string_table_mod_nmap_count += 1
            elif rtype == STBL_TYPE:
                string_table_mod_stbl_count += 1
            else:
                string_table_mod_candidate = False

            # Collection File structural candidate:
            # MANIFEST is auxiliary; all substantive resources must be the
            # dedicated Sims 3 COLL resource type 0x0C07456D.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == COLL_TYPE:
                collection_file_count += 1
            else:
                collection_file_candidate = False

            # Footprint Mod TEST structural candidate:
            # ignore only the standard MANIFEST auxiliary; every other
            # substantive resource must be FTPT.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == FTPT_TYPE:
                footprint_mod_ftpt_count += 1
            else:
                footprint_mod_candidate = False

            # Opportunity Tuning TEST structural candidate:
            # MANIFEST is auxiliary metadata; otherwise this conservative
            # classifier permits only NMAP + XML resources.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                opportunity_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                opportunity_tuning_xml_count += 1
            else:
                opportunity_tuning_candidate = False

            # Jazz Mod TEST structural candidate:
            # Match the observed package conservatively: MANIFEST is auxiliary;
            # every other resource must be NMAP or JAZZ.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                jazz_mod_nmap_count += 1
            elif rtype == JAZZ_TYPE:
                jazz_mod_jazz_count += 1
            else:
                jazz_mod_candidate = False

            # Academic Degree Tuning TEST structural candidate:
            # Match the observed package conservatively: MANIFEST is auxiliary;
            # every other substantive resource must be XML.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == XML_TYPE:
                academic_degree_tuning_xml_count += 1
            else:
                academic_degree_tuning_candidate = False

            # Career Tuning TEST structural candidate:
            # MANIFEST is auxiliary; otherwise only the observed NMAP + XML
            # family is permitted.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                career_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                career_tuning_xml_count += 1
            else:
                career_tuning_candidate = False

            # Book Tuning TEST structural candidate:
            # MANIFEST is auxiliary; otherwise only the observed NMAP + XML
            # family is permitted.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                book_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                book_tuning_xml_count += 1
            else:
                book_tuning_candidate = False

            # Dumpster Diving Tuning TEST structural candidate:
            # MANIFEST is auxiliary; otherwise only the observed NMAP + XML
            # family is permitted.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                dumpster_diving_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                dumpster_diving_tuning_xml_count += 1
            else:
                dumpster_diving_tuning_candidate = False

            # Dreams and Promises Tuning TEST structural candidate:
            # Match the observed package conservatively: MANIFEST is auxiliary;
            # every other substantive resource must be XML.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == XML_TYPE:
                dreams_and_promises_tuning_xml_count += 1
            else:
                dreams_and_promises_tuning_candidate = False

            # Scoring Tuning TEST structural candidate:
            # MANIFEST is auxiliary; otherwise only NMAP + XML are permitted.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                scoring_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                scoring_tuning_xml_count += 1
            else:
                scoring_tuning_candidate = False

            # Boarding School Tuning TEST structural candidate:
            # MANIFEST is auxiliary; otherwise only NMAP + XML are permitted.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                boarding_school_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                boarding_school_tuning_xml_count += 1
            else:
                boarding_school_tuning_candidate = False

            # Buff Tuning TEST structural candidate:
            # MANIFEST is auxiliary; otherwise only NMAP + XML are permitted.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                buff_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                buff_tuning_xml_count += 1
            else:
                buff_tuning_candidate = False

            # Social Tuning TEST structural candidate:
            # MANIFEST is auxiliary; otherwise only NMAP + XML are permitted.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                social_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                social_tuning_xml_count += 1
            else:
                social_tuning_candidate = False

            # Skill Tuning TEST structural candidate:
            # MANIFEST is auxiliary; otherwise only NMAP + XML are permitted.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                skill_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                skill_tuning_xml_count += 1
            else:
                skill_tuning_candidate = False

            # Texture Mod TEST structural candidate:
            # MANIFEST is auxiliary; every substantive resource must be
            # NMAP or IMG. Final classification additionally requires every
            # IMG to have passed the existing DDS texture health/size check.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                texture_mod_nmap_count += 1
            elif rtype == IMG_TYPE:
                texture_mod_img_count += 1
            else:
                texture_mod_candidate = False

            # Fishing Tuning TEST structural candidate:
            # MANIFEST and STBL are allowed auxiliaries; otherwise only
            # NMAP + XML are permitted. STBL cannot trigger this category.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                fishing_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                fishing_tuning_xml_count += 1
            elif rtype == STBL_TYPE:
                fishing_tuning_stbl_count += 1
            else:
                fishing_tuning_candidate = False

            # Cooking Tuning TEST structural candidate:
            # MANIFEST is auxiliary. The observed overhaul is composed only
            # of NMAP, normal XML, localized STBL, small PNG IMAG resources,
            # and the dedicated Store Ingredients resource 0xE5105066.
            if rtype == MANIFEST_TYPE:
                pass
            elif rtype == NMAP_TYPE:
                cooking_tuning_nmap_count += 1
            elif rtype == XML_TYPE:
                cooking_tuning_xml_count += 1
            elif rtype == STBL_TYPE:
                cooking_tuning_stbl_count += 1
            elif rtype == IMAG_PNG_TYPE:
                cooking_tuning_png_count += 1
            elif rtype == STORE_INGREDIENTS_TYPE:
                cooking_tuning_store_ingredients_count += 1
            else:
                cooking_tuning_candidate = False

            if DEBUG:
                print("========== RESOURCE DEBUG ==========")
                print("ENTRY:", entry)
                print("ResourceType:", hex(rtype))
                print("ResourceGroup:", hex(int(entry.ResourceGroup)))
                print("Instance:", hex(resource_instance))
                print("====================================")
                print("")

                if rtype in [0x0333406C, 0x034AEEBC, 0x034AEECB]:
                    print("POSSIBLE CASP:", f"0x{rtype:08X}")

                if rtype in CASP_TYPES:
                    print("******** REAL CASP FOUND ********")
                    print("TYPE:", hex(rtype))
                    print("INSTANCE:", hex(resource_instance))
                    print("GROUP:", hex(int(entry.ResourceGroup)))
                    print("****************************")
                    print(entry)
                    print("INSTANCE:", hex(resource_instance))
                    print("INSTANCE type:", type(resource_instance))
                    print("")

                    print(
                        "RESOURCE:",
                        hex(rtype),
                        "CASP TYPES:",
                        [hex(x) for x in CASP_TYPES]
                    )
                    print(
                        "RESOURCE CHECK:",
                        hex(rtype),
                        "CASP?",
                        rtype in CASP_TYPES
                    )

            if rtype not in known_types:

                unknown_resources.append(
                    f"0x{rtype:08X}"
                )


            # -------------------------
            # RESOURCE PROGRESS
            # -------------------------

            if i > 0 and i % 1000 == 0:
                print(
                    f"{os.path.basename(filename)}: {i}/{resources.Count} resources"
                )

            # -------------------------
            # EA DATABASE CHECK
            # -------------------------

            if DEBUG:
                print(
                    hex(rtype),
                    "Group:",
                    hex(int(entry.ResourceGroup)),
                    "Instance:",
                    hex(int(entry.Instance))
                )
            if rtype == VPXY_TYPE:

                vpxy += 1

                vpxy_instances.append(
                    f"0x{int(entry.Instance):016X}"
                )

                vpxy_ok, vpxy_error = check_vpxy_health(
                    package,
                    entry,
                    resource_objects
                )

                if not vpxy_ok:
                    vpxy_id = f"0x{int(entry.Instance):016X}"
                    broken_vpxy_ids.append(vpxy_id)
                    broken_vpxy_errors.append(
                        f"{vpxy_id}: VPXY Parse Failed - {vpxy_error}"
                    )

                if DEBUG:
                    print(
                        "VPXY FOUND:",
                        f"0x{int(entry.Instance):016X}"
                    )

                    print(
                        "VPXY OBJECT:",
                        type(entry)
                    )

            # -------------------------
            # BONE / BOND HEALTH
            # -------------------------

            elif rtype == BONE_TYPE:

                bones += 1
                bone_id = f"0x{int(entry.Instance):016X}"
                bone_ok, bone_error = check_bone_bond_health(
                    package,
                    entry,
                    resource_objects,
                    "BONE"
                )

                if not bone_ok:
                    broken_bone_ids.append(bone_id)
                    broken_bone_errors.append(
                        f"BONE {bone_id}: {bone_error}"
                    )

            elif rtype == BOND_TYPE:

                bond_id = f"0x{int(entry.Instance):016X}"
                bond_ok, bond_error = check_bone_bond_health(
                    package,
                    entry,
                    resource_objects,
                    "BOND"
                )

                if not bond_ok:
                    broken_bond_ids.append(bond_id)
                    broken_bond_errors.append(
                        f"BOND {bond_id}: {bond_error}"
                    )

            # -------------------------
            # RIG
            # -------------------------

            elif rtype == RIG_TYPE:

                rig_id = f"0x{int(entry.Instance):016X}"

                rig_ok, rig_error = check_rig_health(
                    package,
                    entry,
                    resource_objects
                )

                if not rig_ok:
                    broken_rig_ids.append(rig_id)
                    broken_rig_errors.append(
                        f"RIG {rig_id}: {rig_error}"
                    )

            # -------------------------
            # JAZZ / STATE MACHINE
            # -------------------------

            elif rtype == JAZZ_TYPE:

                jazz_id = f"0x{int(entry.Instance):016X}"

                jazz_ok, jazz_error = check_jazz_health(
                    package,
                    entry,
                    resource_objects
                )

                if not jazz_ok:
                    broken_jazz_ids.append(jazz_id)
                    broken_jazz_errors.append(
                        f"JAZZ {jazz_id}: {jazz_error}"
                    )

            # -------------------------
            # S3SA / SCRIPT ASSEMBLY
            # -------------------------

            elif rtype == S3SA_TYPE:

                s3sa_id = f"0x{int(entry.Instance):016X}"

                s3sa_ok, s3sa_error = check_s3sa_health(
                    package,
                    entry,
                    resource_objects
                )

                if s3sa_ok:
                    if not script_mod_marker:
                        script_mod_marker = True
                        print(
                            "SCRIPT MOD MARKER:",
                            s3sa_id,
                            "| Type=0x073FAA07",
                            "| S3SAParser=OK"
                        )
                    s3sa_defined_type_names.update(
                        extract_s3sa_defined_types(package, entry)
                    )
                else:
                    broken_s3sa_ids.append(s3sa_id)
                    broken_s3sa_errors.append(
                        f"S3SA {s3sa_id}: {s3sa_error}"
                    )

            # -------------------------
            # LITE / LIGHTING
            # -------------------------

            elif rtype == LITE_TYPE:

                lite_id = f"0x{int(entry.Instance):016X}"

                lite_ok, lite_error = check_lite_health(
                    package,
                    entry,
                    resource_objects
                )

                if not lite_ok:
                    broken_lite_ids.append(lite_id)
                    broken_lite_errors.append(
                        f"LITE {lite_id}: {lite_error}"
                    )

            # -------------------------
            # CLIP / ANIMATION
            # -------------------------

            elif rtype == CLIP_TYPE:

                clip_id = f"0x{int(entry.Instance):016X}"

                clip_ok, clip_error = check_clip_health(
                    package,
                    entry,
                    resource_objects
                )

                if not clip_ok:
                    broken_clip_ids.append(clip_id)
                    broken_clip_errors.append(
                        f"CLIP {clip_id}: {clip_error}"
                    )

                else:
                    clip_framecount = get_clip_framecount_for_classification(
                        package,
                        entry,
                        resource_objects
                    )

                    if clip_framecount is not None:
                        print(
                            "POSE CLIP FRAMECOUNT:",
                            clip_id,
                            "| Frames=",
                            clip_framecount
                        )

                        if clip_framecount == 1:
                            pose_one_frame_clip_count += 1
                        elif clip_framecount > 1:
                            pose_multi_frame_clip_count += 1

            # -------------------------
            # MATD / MATERIAL DEFINITION
            # -------------------------

            elif rtype == MATD_TYPE:

                matd_id = f"0x{int(entry.Instance):016X}"

                matd_ok, matd_error = check_matd_health(
                    package,
                    entry,
                    resource_objects
                )

                if not matd_ok:
                    broken_matd_ids.append(matd_id)

                    # MATD belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"MATD {matd_id}")
                    broken_build_buy_errors.append(
                        f"MATD {matd_id}: {matd_error}"
                    )

            # -------------------------
            # MTST
            # -------------------------

            elif rtype == MTST_TYPE:

                mtst_id = f"0x{int(entry.Instance):016X}"

                mtst_ok, mtst_error = check_mtst_health(
                    package,
                    entry,
                    resource_objects
                )

                if not mtst_ok:
                    broken_mtst_ids.append(mtst_id)

                    # MTST belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"MTST {mtst_id}")
                    broken_build_buy_errors.append(
                        f"MTST {mtst_id}: {mtst_error}"
                    )

            # -------------------------
            # MDLR — MODEL REFERENCE / LINK RESOURCE
            # -------------------------

            elif rtype == MDLR_TYPE:

                mdlr_id = f"0x{int(entry.Instance):016X}"

                mdlr_ok, mdlr_error = check_mdlr_health(
                    package,
                    entry,
                    resource_objects
                )

                if not mdlr_ok:
                    broken_mdlr_ids.append(mdlr_id)

                    # MDLR belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"MDLR {mdlr_id}")
                    broken_build_buy_errors.append(
                        f"MDLR {mdlr_id}: {mdlr_error}"
                    )

            # -------------------------
            # CWAL — WALL / FLOOR PATTERN CATALOG RESOURCE
            # -------------------------

            elif rtype == CWAL_TYPE:

                cwal_id = f"0x{int(entry.Instance):016X}"

                cwal_ok, cwal_error = check_cwal_health(
                    package,
                    entry,
                    resource_objects
                )

                if not cwal_ok:
                    broken_cwal_ids.append(cwal_id)

                    # CWAL belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"CWAL {cwal_id}")
                    broken_build_buy_errors.append(
                        f"CWAL {cwal_id}: {cwal_error}"
                    )

            # -------------------------
            # CFEN — CATALOG FENCE RESOURCE
            # -------------------------

            elif rtype == CFEN_TYPE:

                cfen_id = f"0x{int(entry.Instance):016X}"

                cfen_ok, cfen_error = check_cfen_health(
                    package,
                    entry,
                    resource_objects
                )

                if not cfen_ok:
                    broken_cfen_ids.append(cfen_id)

                    # CFEN belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"CFEN {cfen_id}")
                    broken_build_buy_errors.append(
                        f"CFEN {cfen_id}: {cfen_error}"
                    )

            # -------------------------
            # CTPT — TERRAIN PAINT / PAVING CATALOG RESOURCE
            # -------------------------

            elif rtype == CTPT_TYPE:

                ctpt_id = f"0x{int(entry.Instance):016X}"

                ctpt_ok, ctpt_error = check_ctpt_health(
                    package,
                    entry,
                    resource_objects
                )

                if not ctpt_ok:
                    broken_ctpt_ids.append(ctpt_id)

                    # CTPT belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"CTPT {ctpt_id}")
                    broken_build_buy_errors.append(
                        f"CTPT {ctpt_id}: {ctpt_error}"
                    )

            # -------------------------
            # NGMP — HASH MAP RESOURCE
            # -------------------------

            elif rtype == NGMP_TYPE:

                ngmp_id = f"0x{int(entry.Instance):016X}"

                ngmp_ok, ngmp_error = check_ngmp_health(
                    package,
                    entry,
                    resource_objects
                )

                if not ngmp_ok:
                    broken_ngmp_ids.append(ngmp_id)
                    broken_ngmp_errors.append(
                        f"NGMP {ngmp_id}: {ngmp_error}"
                    )

            # -------------------------
            # CRMT — CATALOG ROOF PATTERN RESOURCE
            # -------------------------

            elif rtype == CRMT_TYPE:

                crmt_id = f"0x{int(entry.Instance):016X}"

                crmt_ok, crmt_error = check_crmt_health(
                    package,
                    entry,
                    resource_objects
                )

                if not crmt_ok:
                    broken_crmt_ids.append(crmt_id)

                    # CRMT belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"CRMT {crmt_id}")
                    broken_build_buy_errors.append(
                        f"CRMT {crmt_id}: {crmt_error}"
                    )

            # -------------------------
            # SPT2 — SPEEDTREE RESOURCE
            # -------------------------

            elif rtype == SPT2_TYPE:

                spt2_id = f"0x{int(entry.Instance):016X}"

                spt2_ok, spt2_error = check_spt2_health(
                    package,
                    entry,
                    resource_objects
                )

                if not spt2_ok:
                    broken_spt2_ids.append(spt2_id)

                    # SPT2 is a vegetation / Build-Buy rendering resource.
                    broken_build_buy_ids.append(f"SPT2 {spt2_id}")
                    broken_build_buy_errors.append(
                        f"SPT2 {spt2_id}: {spt2_error}"
                    )

            # -------------------------
            # _SPT — SEPARATE VEGETATION RESOURCE
            # -------------------------

            elif rtype == SPT_TYPE:

                spt_id = f"0x{int(entry.Instance):016X}"

                spt_ok, spt_error = check_spt_health(
                    package,
                    entry,
                    resource_objects
                )

                if not spt_ok:
                    broken_spt_ids.append(spt_id)

                    # _SPT is a vegetation / Build-Buy resource.
                    broken_build_buy_ids.append(f"_SPT {spt_id}")
                    broken_build_buy_errors.append(
                        f"_SPT {spt_id}: {spt_error}"
                    )

            # -------------------------
            # CRAL — CATALOG RAILING RESOURCE
            # -------------------------

            elif rtype == CRAL_TYPE:

                cral_id = f"0x{int(entry.Instance):016X}"

                cral_ok, cral_error = check_cral_health(
                    package,
                    entry,
                    resource_objects
                )

                if not cral_ok:
                    broken_cral_ids.append(cral_id)

                    # CRAL belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"CRAL {cral_id}")
                    broken_build_buy_errors.append(
                        f"CRAL {cral_id}: {cral_error}"
                    )

            # -------------------------
            # LAYO — UI LAYOUT RESOURCE
            # -------------------------

            elif rtype == LAYO_TYPE:

                layo_id = f"0x{int(entry.Instance):016X}"

                layo_ok, layo_error = check_layo_health(
                    package,
                    entry,
                    resource_objects
                )

                if layo_ok:
                    if not ui_layout_marker:
                        ui_layout_marker = True
                        print(
                            "UI LAYOUT MOD MARKER:",
                            layo_id,
                            "| Type=0x025C95B6",
                            "| LAYOParser=OK"
                        )
                else:
                    broken_layo_ids.append(layo_id)
                    broken_layo_errors.append(
                        f"LAYO {layo_id}: {layo_error}"
                    )

            # -------------------------
            # HINF — HAIR INFORMATION
            # -------------------------

            elif rtype == HINF_TYPE:

                hinf_id = f"0x{int(entry.Instance):016X}"

                hinf_ok, hinf_error = check_hinf_health(
                    package,
                    entry,
                    resource_objects
                )

                if hinf_ok:
                    if not hair_color_preset_marker:
                        hair_color_preset_marker = True
                        print(
                            "HAIR COLOR PRESET MARKER:",
                            hinf_id,
                            "| Type=0x063261DA",
                            "| HINFParser=OK"
                        )
                else:
                    broken_hinf_ids.append(hinf_id)
                    broken_hinf_errors.append(
                        f"HINF {hinf_id}: {hinf_error}"
                    )

            # -------------------------
            # 2ARY — OPAQUE / WORLD-SPECIFIC
            # -------------------------

            elif rtype == TWOARY_TYPE:

                twoary_id = f"0x{int(entry.Instance):016X}"

                twoary_ok, twoary_error = check_2ary_health(
                    package,
                    entry,
                    resource_objects
                )

                if not twoary_ok:
                    broken_2ary_ids.append(twoary_id)
                    broken_2ary_errors.append(
                        f"2ARY {twoary_id}: {twoary_error}"
                    )

            # -------------------------
            # COMPLATE
            # -------------------------

            elif rtype == COMPLATE_TYPE:

                complate_id = f"0x{int(entry.Instance):016X}"

                complate_ok, complate_error = check_complate_health(
                    package,
                    entry,
                    resource_objects
                )

                if not complate_ok:
                    broken_complate_ids.append(complate_id)
                    broken_complate_errors.append(
                        f"COMPLATE {complate_id}: {complate_error}"
                    )

            # -------------------------
            # 0x63A33EA7 — GENERIC RCOL
            # -------------------------

            elif rtype == TYPE_63A33EA7:

                type_63a33ea7_id = f"0x{int(entry.Instance):016X}"

                type_63a33ea7_ok, type_63a33ea7_error = check_63a33ea7_health(
                    package,
                    entry,
                    resource_objects
                )

                if not type_63a33ea7_ok:
                    broken_63a33ea7_ids.append(type_63a33ea7_id)
                    broken_63a33ea7_errors.append(
                        f"63A33EA7 {type_63a33ea7_id}: {type_63a33ea7_error}"
                    )

            # -------------------------
            # FIREPLACE CATALOG
            # -------------------------

            elif rtype == FIREPLACE_TYPE:

                fireplace_id = f"0x{int(entry.Instance):016X}"

                fireplace_ok, fireplace_error = check_fireplace_health(
                    package,
                    entry,
                    resource_objects
                )

                if not fireplace_ok:
                    broken_fireplace_ids.append(fireplace_id)
                    broken_fireplace_errors.append(
                        f"FIREPLACE {fireplace_id}: {fireplace_error}"
                    )

            # -------------------------
            # THUM PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # THUM 0x626F60CC MIXED PNG/JPEG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # THUM 0x05B1B524 PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # THUM 0x05B1B525 PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # THUM 0x05B1B526 PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # THUM 0x2653E3C9 FENCE PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # THUM 0xFCEAB65B CUSTOM COLOURSWATCH PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # ICON 0xD84E7FC5 SMALL LOT PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # ICON 0xD84E7FC6 MEDIUM LOT PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # ICON 0xD84E7FC7 LARGE LOT PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # TWNI 0x0668F635 WORLD DESCRIPTION IMAGE PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # THUM 0x2D4284F1 RAILING PNG DIAGNOSTIC
            # -------------------------

            # -------------------------
            # MANIFEST 0x73E93EEB XML DIAGNOSTIC
            # -------------------------

            # -------------------------
            # TYPE 0x8F78F74B PNG HEALTH
            # -------------------------

            elif rtype == TYPE_8F78F74B:

                type_8f78f74b_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x8F78F74B, Group=0x{int(entry.ResourceGroup):08X})"
                )

                type_8f78f74b_ok, type_8f78f74b_error = check_8f78f74b_png_health(
                    package,
                    entry,
                    resource_objects
                )

                if not type_8f78f74b_ok:
                    broken_8f78f74b_ids.append(type_8f78f74b_id)
                    broken_8f78f74b_errors.append(
                        f"8F78F74B {type_8f78f74b_id}: {type_8f78f74b_error}"
                    )

            elif rtype == MANIFEST_TYPE:

                manifest_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x73E93EEB, Group=0x{int(entry.ResourceGroup):08X})"
                )

                manifest_ok, manifest_error = check_manifest_73e93eeb_health(
                    package,
                    entry,
                    resource_objects
                )

                if not manifest_ok:
                    broken_manifest_ids.append(manifest_id)
                    broken_manifest_errors.append(
                        f"MANIFEST {manifest_id}: {manifest_error}"
                    )

            elif rtype == THUM_2D4284F1_TYPE:

                thum_2d4284f1_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x2D4284F1, Group=0x{int(entry.ResourceGroup):08X})"
                )

                thum_2d4284f1_ok, thum_2d4284f1_error = check_2d4284f1_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_2d4284f1_ok:
                    broken_thum_ids.append(thum_2d4284f1_id)
                    broken_thum_errors.append(
                        f"THUM {thum_2d4284f1_id}: {thum_2d4284f1_error}"
                    )

            elif rtype == TWNI_TYPE:

                twni_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x0668F635, Group=0x{int(entry.ResourceGroup):08X})"
                )

                twni_ok, twni_error = check_twni_health(
                    package,
                    entry,
                    resource_objects
                )

                if not twni_ok:
                    broken_twni_ids.append(twni_id)
                    broken_twni_errors.append(
                        f"TWNI {twni_id}: {twni_error}"
                    )

            elif rtype == ICON_D84E7FC7_TYPE:

                icon_d84e7fc7_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0xD84E7FC7, Group=0x{int(entry.ResourceGroup):08X})"
                )

                icon_d84e7fc7_ok, icon_d84e7fc7_error = check_d84e7fc7_icon_health(
                    package,
                    entry,
                    resource_objects
                )

                if not icon_d84e7fc7_ok:
                    broken_icon_ids.append(icon_d84e7fc7_id)
                    broken_icon_errors.append(
                        f"ICON {icon_d84e7fc7_id}: {icon_d84e7fc7_error}"
                    )

            elif rtype == ICON_D84E7FC6_TYPE:

                icon_d84e7fc6_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0xD84E7FC6, Group=0x{int(entry.ResourceGroup):08X})"
                )

                icon_d84e7fc6_ok, icon_d84e7fc6_error = check_d84e7fc6_icon_health(
                    package,
                    entry,
                    resource_objects
                )

                if not icon_d84e7fc6_ok:
                    broken_icon_ids.append(icon_d84e7fc6_id)
                    broken_icon_errors.append(
                        f"ICON {icon_d84e7fc6_id}: {icon_d84e7fc6_error}"
                    )

            elif rtype == ICON_D84E7FC5_TYPE:

                icon_d84e7fc5_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0xD84E7FC5, Group=0x{int(entry.ResourceGroup):08X})"
                )

                icon_d84e7fc5_ok, icon_d84e7fc5_error = check_d84e7fc5_icon_health(
                    package,
                    entry,
                    resource_objects
                )

                if not icon_d84e7fc5_ok:
                    broken_icon_ids.append(icon_d84e7fc5_id)
                    broken_icon_errors.append(
                        f"ICON {icon_d84e7fc5_id}: {icon_d84e7fc5_error}"
                    )

            elif rtype == THUM_FCEAB65B_TYPE:

                thum_fceab65b_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0xFCEAB65B, Group=0x{int(entry.ResourceGroup):08X})"
                )

                thum_fceab65b_ok, thum_fceab65b_error = check_fceab65b_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_fceab65b_ok:
                    broken_thum_ids.append(thum_fceab65b_id)
                    broken_thum_errors.append(
                        f"THUM {thum_fceab65b_id}: {thum_fceab65b_error}"
                    )

            elif rtype == THUM_2653E3C9_TYPE:

                thum_2653e3c9_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x2653E3C9, Group=0x{int(entry.ResourceGroup):08X})"
                )

                thum_2653e3c9_ok, thum_2653e3c9_error = check_2653e3c9_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_2653e3c9_ok:
                    broken_thum_ids.append(thum_2653e3c9_id)
                    broken_thum_errors.append(
                        f"THUM {thum_2653e3c9_id}: {thum_2653e3c9_error}"
                    )

            elif rtype == THUM_05B1B526_TYPE:

                thum_05b1b526_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x05B1B526, Group=0x{int(entry.ResourceGroup):08X})"
                )

                thum_05b1b526_ok, thum_05b1b526_error = check_05b1b526_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_05b1b526_ok:
                    broken_thum_ids.append(thum_05b1b526_id)
                    broken_thum_errors.append(
                        f"THUM {thum_05b1b526_id}: {thum_05b1b526_error}"
                    )

            elif rtype == THUM_05B1B525_TYPE:

                thum_05b1b525_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x05B1B525, Group=0x{int(entry.ResourceGroup):08X})"
                )

                thum_05b1b525_ok, thum_05b1b525_error = check_05b1b525_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_05b1b525_ok:
                    broken_thum_ids.append(thum_05b1b525_id)
                    broken_thum_errors.append(
                        f"THUM {thum_05b1b525_id}: {thum_05b1b525_error}"
                    )

            elif rtype == THUM_05B1B524_TYPE:

                thum_05b1b524_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x05B1B524, Group=0x{int(entry.ResourceGroup):08X})"
                )

                thum_05b1b524_ok, thum_05b1b524_error = check_05b1b524_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_05b1b524_ok:
                    broken_thum_ids.append(thum_05b1b524_id)
                    broken_thum_errors.append(
                        f"THUM {thum_05b1b524_id}: {thum_05b1b524_error}"
                    )

            elif rtype == THUM_626F60CC_TYPE:

                thum_626f60cc_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x626F60CC, Group=0x{int(entry.ResourceGroup):08X})"
                )

                thum_626f60cc_ok, thum_626f60cc_error = check_626f60cc_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_626f60cc_ok:
                    # 0x626F60CC is part of the existing THUM semantic family.
                    # Reuse broken_thum_* so it counts as THUM, not a new family.
                    broken_thum_ids.append(thum_626f60cc_id)
                    broken_thum_errors.append(
                        f"THUM {thum_626f60cc_id}: {thum_626f60cc_error}"
                    )

            elif rtype == THUM_TYPE:

                thum_id = f"0x{int(entry.Instance):016X}"

                thum_ok, thum_error = check_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_ok:
                    broken_thum_ids.append(thum_id)
                    broken_thum_errors.append(
                        f"THUM {thum_id}: {thum_error}"
                    )

            # -------------------------
            # THUM 0xAD366F95
            # -------------------------

            elif rtype == THUM_AD366F95_TYPE:

                thum_ad366f95_id = f"0x{int(entry.Instance):016X}"

                thum_ad366f95_ok, thum_ad366f95_error = check_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_ad366f95_ok:
                    broken_thum_ids.append(thum_ad366f95_id)
                    broken_thum_errors.append(
                        f"THUM {thum_ad366f95_id}: {thum_ad366f95_error}"
                    )

            # -------------------------
            # 0x626F60CE PNG DIAGNOSTIC
            # -------------------------

            elif rtype == PNG_626F60CE_TYPE:

                png_626f60ce_id = f"0x{int(entry.Instance):016X}"

                png_626f60ce_ok, png_626f60ce_error = check_626f60ce_health(
                    package,
                    entry,
                    resource_objects
                )

                if not png_626f60ce_ok:
                    broken_626f60ce_ids.append(png_626f60ce_id)
                    broken_626f60ce_errors.append(
                        f"626F60CE {png_626f60ce_id}: {png_626f60ce_error}"
                    )

            # -------------------------
            # ICON PNG DIAGNOSTIC
            # -------------------------

            elif rtype == ICON_TYPE:

                icon_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x2E75C765, Group=0x{int(entry.ResourceGroup):08X})"
                )

                icon_ok, icon_error = check_icon_health(
                    package,
                    entry,
                    resource_objects
                )

                if not icon_ok:
                    broken_icon_ids.append(icon_id)
                    broken_icon_errors.append(
                        f"ICON {icon_id}: {icon_error}"
                    )

            # -------------------------
            # ICON 0x2E75C764 PNG DIAGNOSTIC
            # -------------------------

            elif rtype == ICON_2E75C764_TYPE:

                icon_2e75c764_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x2E75C764, Group=0x{int(entry.ResourceGroup):08X})"
                )

                icon_2e75c764_ok, icon_2e75c764_error = check_2e75c764_icon_health(
                    package,
                    entry,
                    resource_objects
                )

                if not icon_2e75c764_ok:
                    broken_icon_ids.append(icon_2e75c764_id)
                    broken_icon_errors.append(
                        f"ICON {icon_2e75c764_id}: {icon_2e75c764_error}"
                    )

            # -------------------------
            # ICON 0x2E75C766 PNG DIAGNOSTIC
            # -------------------------

            elif rtype == ICON_2E75C766_TYPE:

                icon_2e75c766_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x2E75C766, Group=0x{int(entry.ResourceGroup):08X})"
                )

                icon_2e75c766_ok, icon_2e75c766_error = check_2e75c766_icon_health(
                    package,
                    entry,
                    resource_objects
                )

                if not icon_2e75c766_ok:
                    broken_icon_ids.append(icon_2e75c766_id)
                    broken_icon_errors.append(
                        f"ICON {icon_2e75c766_id}: {icon_2e75c766_error}"
                    )


            # -------------------------
            # ICON 0x2E75C767 PNG DIAGNOSTIC
            # -------------------------

            elif rtype == ICON_2E75C767_TYPE:

                icon_2e75c767_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Type=0x2E75C767, Group=0x{int(entry.ResourceGroup):08X})"
                )

                icon_2e75c767_ok, icon_2e75c767_error = check_2e75c767_icon_health(
                    package,
                    entry,
                    resource_objects
                )

                if not icon_2e75c767_ok:
                    broken_icon_ids.append(icon_2e75c767_id)
                    broken_icon_errors.append(
                        f"ICON {icon_2e75c767_id}: {icon_2e75c767_error}"
                    )


            # -------------------------
            # 0x0580A2B5 PNG DIAGNOSTIC
            # -------------------------

            elif rtype == PNG_0580A2B5_TYPE:

                thum_0580a2b5_id = f"0x{int(entry.Instance):016X}"

                thum_0580a2b5_ok, thum_0580a2b5_error = check_0580a2b5_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_0580a2b5_ok:
                    broken_0580a2b5_ids.append(thum_0580a2b5_id)
                    broken_0580a2b5_errors.append(
                        f"THUM 0580A2B5 {thum_0580a2b5_id}: "
                        f"{thum_0580a2b5_error}"
                    )

            # -------------------------
            # THUM 0x0580A2B4 PNG DIAGNOSTIC
            # -------------------------

            elif rtype == THUM_0580A2B4_TYPE:

                thum_0580a2b4_id = f"0x{int(entry.Instance):016X}"

                thum_0580a2b4_ok, thum_0580a2b4_error = check_0580a2b4_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_0580a2b4_ok:
                    broken_0580a2b4_ids.append(thum_0580a2b4_id)
                    broken_0580a2b4_errors.append(
                        f"THUM 0580A2B4 {thum_0580a2b4_id}: "
                        f"{thum_0580a2b4_error}"
                    )

            # -------------------------
            # THUM 0x0580A2B6 PNG DIAGNOSTIC
            # -------------------------

            elif rtype == THUM_0580A2B6_TYPE:

                thum_0580a2b6_id = f"0x{int(entry.Instance):016X}"

                thum_0580a2b6_ok, thum_0580a2b6_error = check_0580a2b6_thum_health(
                    package,
                    entry,
                    resource_objects
                )

                if not thum_0580a2b6_ok:
                    broken_0580a2b6_ids.append(thum_0580a2b6_id)
                    broken_0580a2b6_errors.append(
                        f"THUM 0580A2B6 {thum_0580a2b6_id}: "
                        f"{thum_0580a2b6_error}"
                    )

            # -------------------------
            # SNAP 0x0580A2CD PNG DIAGNOSTIC
            # -------------------------

            elif rtype == SNAP_0580A2CD_TYPE:

                snap_0580a2cd_id = f"0x{int(entry.Instance):016X}"

                snap_0580a2cd_ok, snap_0580a2cd_error = check_0580a2cd_snap_health(
                    package,
                    entry,
                    resource_objects
                )

                if not snap_0580a2cd_ok:
                    broken_0580a2cd_ids.append(snap_0580a2cd_id)
                    broken_0580a2cd_errors.append(
                        f"SNAP 0580A2CD {snap_0580a2cd_id}: "
                        f"{snap_0580a2cd_error}"
                    )

            # -------------------------
            # SNAP 0x0580A2CF PNG DIAGNOSTIC
            # -------------------------

            elif rtype == SNAP_0580A2CF_TYPE:

                snap_0580a2cf_id = (
                    f"0x{int(entry.Instance):016X} "
                    f"(Group=0x{int(entry.ResourceGroup):08X})"
                )

                snap_0580a2cf_ok, snap_0580a2cf_error = check_0580a2cf_snap_health(
                    package,
                    entry,
                    resource_objects
                )

                if not snap_0580a2cf_ok:
                    broken_0580a2cf_ids.append(snap_0580a2cf_id)
                    broken_0580a2cf_errors.append(
                        f"SNAP 0580A2CF {snap_0580a2cf_id}: "
                        f"{snap_0580a2cf_error}"
                    )

            # -------------------------
            # CAS SLIDERS
            # -------------------------

            elif rtype == STBL_TYPE:

                stbl_id = f"0x{int(entry.Instance):016X}"

                stbl_instances.append(stbl_id)

                stbl_ok, stbl_error = check_stbl_health(
                    package,
                    entry,
                    resource_objects
                )

                if not stbl_ok:
                    broken_stbl_ids.append(stbl_id)
                    broken_stbl_errors.append(
                        f"STBL {stbl_id}: {stbl_error}"
                    )


            elif rtype == BGEO_TYPE:

                bgeo_id = f"0x{int(entry.Instance):016X}"
                bgeo_ok, bgeo_error = check_morph_health(
                    package,
                    entry,
                    resource_objects,
                    "BGEO"
                )

                if not bgeo_ok:
                    broken_morph_ids.append(f"BGEO {bgeo_id}")
                    broken_morph_errors.append(
                        f"BGEO {bgeo_id}: {bgeo_error}"
                    )

                bgeo_instances.append(bgeo_id)


            elif rtype == BBLN_TYPE:

                bbln_id = f"0x{int(entry.Instance):016X}"
                bbln_ok, bbln_error = check_morph_health(
                    package,
                    entry,
                    resource_objects,
                    "BBLN"
                )

                if not bbln_ok:
                    broken_morph_ids.append(f"BBLN {bbln_id}")
                    broken_morph_errors.append(
                        f"BBLN {bbln_id}: {bbln_error}"
                    )


            elif rtype == FACE_TYPE:

                face_id = f"0x{int(entry.Instance):016X}"
                face_ok, face_error = check_morph_health(
                    package,
                    entry,
                    resource_objects,
                    "FACE"
                )

                if not face_ok:
                    broken_morph_ids.append(f"FACE {face_id}")
                    broken_morph_errors.append(
                        f"FACE {face_id}: {face_error}"
                    )

                face += 1
                face_instances.append(face_id)


            elif rtype == FBLN_TYPE:

                fbln_id = f"0x{int(entry.Instance):016X}"
                fbln_ok, fbln_error = check_morph_health(
                    package,
                    entry,
                    resource_objects,
                    "FBLN"
                )

                if not fbln_ok:
                    broken_morph_ids.append(f"FBLN {fbln_id}")
                    broken_morph_errors.append(
                        f"FBLN {fbln_id}: {fbln_error}"
                    )

                fbln += 1
                slider_count += 1
                fbln_instances.append(fbln_id)


            elif rtype == KEY_TYPE:

                nmap_id = f"0x{int(entry.Instance):016X}"

                key_instances.append(nmap_id)

                nmap_ok, nmap_error = check_nmap_health(
                    package,
                    entry,
                    resource_objects
                )

                if not nmap_ok:
                    broken_nmap_ids.append(nmap_id)
                    broken_nmap_errors.append(
                        f"NMAP {nmap_id}: {nmap_error}"
                    )

            # -------------------------
            # SKIN TONE (TONE)
            # -------------------------

            elif rtype == TONE_TYPE:

                tone_id = f"0x{int(entry.Instance):016X}"
                tone_ok, tone_error = check_skin_tone_health(
                    package,
                    entry,
                    resource_objects
                )

                if not tone_ok:
                    broken_tone_ids.append(tone_id)
                    broken_tone_errors.append(
                        f"TONE {tone_id}: {tone_error}"
                    )

                if cc_type == "Unknown":
                    cc_type = "Skin"

            # -------------------------
            # PTRN / PATTERN LIST
            # -------------------------

            elif rtype == PTRN_TYPE:

                ptrn_id = f"0x{int(entry.Instance):016X}"
                ptrn_ok, ptrn_error = check_ptrn_health(
                    package,
                    entry,
                    resource_objects
                )

                if not ptrn_ok:
                    broken_ptrn_ids.append(ptrn_id)
                    broken_ptrn_errors.append(
                        f"PTRN {ptrn_id}: {ptrn_error}"
                    )

                if cc_type == "Unknown":
                    cc_type = "Pattern"

            # -------------------------
            # TXTF / FABRIC COMPOSITOR
            # -------------------------

            elif rtype == TXTF_TYPE:

                txtf_id = f"0x{int(entry.Instance):016X}"
                txtf_ok, txtf_error = check_txtf_health(
                    package,
                    entry,
                    resource_objects
                )

                if not txtf_ok:
                    broken_txtf_ids.append(txtf_id)
                    broken_txtf_errors.append(
                        f"TXTF {txtf_id}: {txtf_error}"
                    )

            # -------------------------
            # TXTC / TEXTURE COMPOSITOR
            # -------------------------

            elif rtype == TXTC_TYPE:

                txtc_id = f"0x{int(entry.Instance):016X}"
                txtc_ok, txtc_error = check_txtc_health(
                    package,
                    entry,
                    resource_objects
                )

                if not txtc_ok:
                    broken_txtc_ids.append(txtc_id)
                    broken_txtc_errors.append(
                        f"TXTC {txtc_id}: {txtc_error}"
                    )

            # -------------------------
            # PRESETS
            # -------------------------

            elif rtype == PRESET_TYPE:

                presets += 1

                preset_id = f"0x{int(entry.Instance):016X}"
                preset_ok, preset_error = check_preset_health(
                    package,
                    entry,
                    resource_objects
                )

                if not preset_ok:
                    broken_preset_ids.append(preset_id)
                    broken_preset_errors.append(
                        f"PRESET {preset_id}: {preset_error}"
                    )

                if cc_type == "Unknown":

                    cc_type = "Preset"

            # -------------------------
            # SIM / HOUSEHOLD DATA
            # -------------------------

            elif rtype == SIMO_TYPE:

                simo += 1

                simo_id = f"0x{int(entry.Instance):016X}"
                simo_ok, simo_error = check_simo_health(
                    package,
                    entry,
                    resource_objects
                )

                if not simo_ok:
                    broken_simo_ids.append(simo_id)
                    broken_simo_errors.append(
                        f"SIMO {simo_id}: {simo_error}"
                    )

            elif rtype == SIME_TYPE:

                # Classification-only evidence for saved/exported Sim packages.
                # Do not introduce a speculative SIME health heuristic here.
                sime += 1

            elif rtype == FAMD_TYPE:

                famd += 1

                famd_id = f"0x{int(entry.Instance):016X}"
                famd_ok, famd_error = check_famd_health(
                    package,
                    entry,
                    resource_objects
                )

                if not famd_ok:
                    broken_famd_ids.append(famd_id)
                    broken_famd_errors.append(
                        f"FAMD {famd_id}: {famd_error}"
                    )

            elif rtype == XML_TYPE:

                xml_id = f"0x{int(entry.Instance):016X}"
                xml_ok, xml_error = check_xml_health(
                    package,
                    entry,
                    resource_objects
                )

                if not xml_ok:
                    broken_xml_ids.append(xml_id)
                    broken_xml_errors.append(
                        f"XML {xml_id}: {xml_error}"
                    )

                elif (
                    not career_marker
                    and detect_career_xml_marker(package, entry)
                ):
                    career_marker = True
                    print(
                        f"CAREER MARKER: {xml_id} | "
                        f"Root=CareerFiles | Status=FOUND"
                    )

                if xml_ok:
                    cooking_marker = detect_cooking_tuning_xml_marker(
                        package,
                        entry,
                        XML_TYPE
                    )

                    if cooking_marker is not None:
                        cooking_tuning_valid_xml_count += 1

                        if cooking_marker == "recipe_process":
                            cooking_tuning_recipe_process_count += 1
                        elif cooking_marker == "recipe_master":
                            cooking_tuning_recipe_master_count += 1
                        elif cooking_marker == "ingredients":
                            cooking_tuning_ingredients_count += 1
                        elif cooking_marker == "plants":
                            cooking_tuning_plants_count += 1
                        elif cooking_marker == "bars":
                            cooking_tuning_bars_count += 1
                        elif cooking_marker == "shop_tabs":
                            cooking_tuning_shop_tabs_count += 1

                        print(
                            f"COOKING TUNING XML MARKER: {xml_id} | "
                            f"Schema={cooking_marker}"
                        )

                if (
                    xml_ok
                    and detect_fishing_tuning_xml_marker(package, entry)
                ):
                    fishing_tuning_valid_xml_count += 1
                    print(
                        f"FISHING TUNING XML MARKER: {xml_id} | "
                        f"Root=Fishing | Fish/Spawner/BoxItems/GlobalSpawner=FOUND"
                    )

                if (
                    xml_ok
                    and detect_skill_tuning_xml_marker(package, entry)
                ):
                    skill_tuning_valid_xml_count += 1
                    print(
                        f"SKILL TUNING XML MARKER: {xml_id} | "
                        f"Root=Skills | SkillListSchema=FOUND"
                    )

                if (
                    xml_ok
                    and detect_social_tuning_xml_marker(package, entry)
                ):
                    social_tuning_valid_xml_count += 1
                    print(
                        f"SOCIAL TUNING XML MARKER: {xml_id} | "
                        f"Root=Social | ActionSchema=FOUND"
                    )

                if (
                    xml_ok
                    and detect_buff_tuning_xml_marker(package, entry)
                ):
                    buff_tuning_valid_xml_count += 1
                    print(
                        f"BUFF TUNING XML MARKER: {xml_id} | "
                        f"Root=buffs | BuffListSchema=FOUND"
                    )

                if (
                    xml_ok
                    and detect_boarding_school_tuning_xml_marker(package, entry)
                ):
                    boarding_school_tuning_valid_xml_count += 1
                    print(
                        f"BOARDING SCHOOL TUNING XML MARKER: {xml_id} | "
                        f"Root=BoardingSchool | SchoolSchema=FOUND"
                    )

                if (
                    xml_ok
                    and detect_scoring_tuning_xml_marker(package, entry)
                ):
                    scoring_tuning_valid_xml_count += 1
                    print(
                        f"SCORING TUNING XML MARKER: {xml_id} | "
                        f"Root=Scoring | MethodsSchema=FOUND"
                    )

                if (
                    xml_ok
                    and detect_dreams_and_promises_tuning_xml_marker(package, entry)
                ):
                    dreams_and_promises_tuning_valid_xml_count += 1
                    print(
                        f"DREAMS AND PROMISES TUNING XML MARKER: {xml_id} | "
                        f"Root=DreamsAndPromisesNodes | PrimitiveSchema=FOUND"
                    )

                if (
                    xml_ok
                    and detect_dumpster_diving_tuning_xml_marker(package, entry)
                ):
                    dumpster_diving_tuning_valid_xml_count += 1
                    print(
                        f"DUMPSTER DIVING TUNING XML MARKER: {xml_id} | "
                        f"Root=DumpsterDiveData | RewardObjectsSchema=FOUND"
                    )

                if (
                    xml_ok
                    and detect_book_tuning_xml_marker(package, entry)
                ):
                    book_tuning_valid_xml_count += 1
                    print(
                        f"BOOK TUNING XML MARKER: {xml_id} | "
                        f"Root=Books | BookSchema=FOUND"
                    )

                if xml_ok:
                    career_tuning_schema = None

                    if detect_career_tuning_xml_marker(package, entry):
                        career_tuning_schema = "Careers"
                    else:
                        career_tuning_schema = (
                            detect_career_overhaul_tuning_xml_marker(
                                package,
                                entry
                            )
                        )

                    if career_tuning_schema is not None:
                        career_tuning_valid_xml_count += 1
                        print(
                            f"CAREER TUNING XML MARKER: {xml_id} | "
                            f"Root={career_tuning_schema} | "
                            f"CareerSchema=FOUND"
                        )

                if (
                    xml_ok
                    and detect_academic_degree_tuning_xml_marker(package, entry)
                ):
                    academic_degree_tuning_valid_xml_count += 1
                    print(
                        f"ACADEMIC DEGREE TUNING XML MARKER: {xml_id} | "
                        f"Root=AcademicDegrees | Status=FOUND"
                    )

                if (
                    xml_ok
                    and detect_opportunity_tuning_xml_marker(package, entry)
                ):
                    opportunity_tuning_valid_xml_count += 1
                    print(
                        f"OPPORTUNITY TUNING XML MARKER: {xml_id} | "
                        f"Schema=Opportunities_* | Status=FOUND"
                    )

                if (
                    xml_ok
                    and not tuning_marker
                    and detect_tuning_xml_marker(package, entry)
                ):
                    tuning_marker = True
                    print(
                        f"TUNING MARKER: {xml_id} | "
                        f"Root=base | Current_Tuning=FOUND"
                    )

                if (
                    xml_ok
                    and not random_event_module_marker
                    and detect_random_event_module_xml_marker(package, entry)
                ):
                    random_event_module_marker = True
                    print(
                        f"RANDOM EVENT MODULE MARKER: {xml_id} | "
                        f"Root=RandomEvents | EventSchema=FOUND"
                    )

            elif rtype == STORE_INGREDIENTS_TYPE:

                store_ingredients_id = f"0x{int(entry.Instance):016X}"
                cooking_store_marker = detect_cooking_tuning_xml_marker(
                    package,
                    entry,
                    STORE_INGREDIENTS_TYPE
                )

                if cooking_store_marker == "store_ingredients":
                    cooking_tuning_valid_store_ingredients_count += 1
                    print(
                        f"COOKING TUNING STORE INGREDIENTS MARKER: "
                        f"{store_ingredients_id} | "
                        f"Root=Ingredients_store | Status=FOUND"
                    )

            # -------------------------
            # LDES / LOT DESCRIPTION - CLASSIFICATION ONLY
            # -------------------------
            # 0xD063545B is Sims 3's dedicated Lot Description resource.
            # Final classification additionally requires at least one of the
            # existing dedicated lot preview image resources.
            elif rtype == LDES_TYPE:

                if not lot_ldes_marker:
                    lot_ldes_marker = True
                    ldes_id = f"0x{int(entry.Instance):016X}"
                    print(
                        f"LOT LDES MARKER: {ldes_id} | "
                        f"Type=0x{LDES_TYPE:08X} | Resource=LDES"
                    )

            # -------------------------
            # _INI / CONFIGURATION RESOURCE - CLASSIFICATION ONLY
            # -------------------------
            # 0x1F886EAD is the Sims 3 _INI resource type. This marker is
            # intentionally classification-only: it does not add or weaken
            # any health check. Final classification is applied only if the
            # package remains Unknown after stronger established categories.

            elif rtype == INI_TYPE:

                if not ini_mod_marker:
                    ini_mod_marker = True
                    ini_id = f"0x{int(entry.Instance):016X}"
                    print(
                        f"INI MOD MARKER: {ini_id} | "
                        f"Type=0x{INI_TYPE:08X} | Resource=_INI"
                    )

            # -------------------------
            # ITUN / INTERACTION TUNING - CLASSIFICATION ONLY
            # -------------------------

            elif rtype == ITUN_TYPE:

                if (
                    not interaction_tuning_marker
                    and detect_interaction_tuning_marker(package, entry)
                ):
                    interaction_tuning_marker = True
                    itun_id = f"0x{int(entry.Instance):016X}"
                    print(
                        f"INTERACTION TUNING MARKER: {itun_id} | "
                        f"Type=0x{ITUN_TYPE:08X} | ITUNSchema=FOUND"
                    )

            # -------------------------
            # IMAG (JPEG / PNG / TGA) - DIAGNOSTIC ONLY
            # -------------------------

            elif rtype in IMAG_TYPES:

                imag_id = f"0x{int(entry.Instance):016X}"
                imag_ok, imag_error = check_imag_resource_health(
                    package,
                    entry,
                    resource_objects
                )

                if not imag_ok:
                    broken_imag_ids.append(imag_id)
                    broken_imag_errors.append(
                        f"IMAG {imag_id}: {imag_error}"
                    )

            # -------------------------
            # DDS IMG
            # -------------------------

            elif rtype == IMG_TYPE:

                img_count += 1
                img_instances.append(
                    f"0x{int(entry.Instance):016X}"
                )


                try:

                    size = get_texture_size(
                        package,
                        entry,
                        resource_objects,
                        broken_img_ids,
                        broken_img_errors
                    )

                    if DEBUG:
                        print("TEXTURE SIZE RESULT:", size)

                    if size:

                        valid_texture_count += 1

                        w, h = size

                        text = f"{w}x{h}"

                        img_sizes.append(text)

                        resolutions[text] = (
                            resolutions.get(text, 0) + 1
                        )

                        if w * h > biggest:

                            biggest = w * h
                            biggest_size = text

                        if w >= 4096 or h >= 4096:

                            fourk += 1

                        elif w >= 2048 or h >= 2048:

                            two_k += 1

                    size = None
                    

                except Exception as e:

                    print("IMG ERROR")

                    if DEBUG:
                        print(e)

            # -------------------------
            # GEOM
            # -------------------------

            elif rtype == GEOM_TYPE:

                if DEBUG:
                    print(
                        "FOUND GEOM:",
                        f"0x{int(entry.Instance):016X}"
                    )

                geom += 1
                geom_instances.append(
                    f"0x{int(entry.Instance):016X}"
                )

                vertices, faces, geom_ok, geom_error = get_geom_poly_count(
                    package,
                    entry,
                    resource_objects
                )

                if not geom_ok:
                    geom_id = f"0x{int(entry.Instance):016X}"
                    broken_geom_ids.append(geom_id)
                    broken_geom_errors.append(
                        f"{geom_id}: {geom_error}"
                    )

                if faces > highest_polys and faces > 100:

                    highest_polys = faces
                    total_vertices = vertices
                    highest_geom_id = str(entry.Instance)

                    if DEBUG:

                        print(
                            f"HIGHEST GEOM -> Vertices={vertices:,} Faces={faces:,}"
                        )

                try:

                    geom_size = int(entry.Filesize)

                except:

                    geom_size = 0


                geom_total += geom_size


                if geom_size > geom_largest:

                    geom_largest = geom_size

            # -------------------------
            # OBJD
            # -------------------------

            elif rtype == OBJD_TYPE:

                objd += 1

                objd_id = f"0x{int(entry.Instance):016X}"
                objd_ok, objd_error = check_objd_health(
                    package,
                    entry,
                    resource_objects
                )

                if not objd_ok:
                    broken_objd_ids.append(objd_id)

                    # OBJD belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"OBJD {objd_id}")
                    broken_build_buy_errors.append(
                        f"OBJD {objd_id}: {objd_error}"
                    )


            # -------------------------
            # OBJK
            # -------------------------

            elif rtype == OBJK_TYPE:

                objk += 1

                objk_id = f"0x{int(entry.Instance):016X}"
                objk_ok, objk_error = check_objk_health(
                    package,
                    entry,
                    resource_objects
                )

                if not objk_ok:
                    broken_objk_ids.append(objk_id)

                    # OBJK belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"OBJK {objk_id}")
                    broken_build_buy_errors.append(
                        f"OBJK {objk_id}: {objk_error}"
                    )
                else:
                    script_class_ref, script_class_name = (
                        detect_objk_script_class(package, entry)
                    )
                    if (
                        script_class_ref
                        and script_class_name
                        and script_class_name not in objk_script_class_names
                    ):
                        objk_script_class_names.append(script_class_name)

                    custom_script_ref, custom_script_class = (
                        detect_custom_objk_script_class(package, entry)
                    )
                    if custom_script_ref:
                        external_script_class_marker = True
                        if (
                            custom_script_class
                            and custom_script_class not in external_script_class_names
                        ):
                            external_script_class_names.append(custom_script_class)


            # -------------------------
            # MODL
            # -------------------------

            elif rtype == MODL_TYPE:

                modl_id = f"0x{int(entry.Instance):016X}"
                modl_ok, modl_error = check_build_buy_mesh_health(
                    package,
                    entry,
                    resource_objects,
                    "MODL"
                )

                if not modl_ok:
                    broken_build_buy_ids.append(f"MODL {modl_id}")
                    broken_build_buy_errors.append(
                        f"MODL {modl_id}: {modl_error}"
                    )

                modl += 1


            # -------------------------
            # MLOD
            # -------------------------

            elif rtype == MLOD_TYPE:

                mlod_id = f"0x{int(entry.Instance):016X}"
                mlod_ok, mlod_error = check_build_buy_mesh_health(
                    package,
                    entry,
                    resource_objects,
                    "MLOD"
                )

                if not mlod_ok:
                    broken_build_buy_ids.append(f"MLOD {mlod_id}")
                    broken_build_buy_errors.append(
                        f"MLOD {mlod_id}: {mlod_error}"
                    )

                mlod += 1

            # -------------------------
            # RSLT
            # -------------------------

            elif rtype == RSLT_TYPE:

                rslt_id = f"0x{int(entry.Instance):016X}"

                rslt_ok, rslt_error = check_rslt_health(
                    package,
                    entry,
                    resource_objects
                )

                if not rslt_ok:
                    broken_rslt_ids.append(rslt_id)

                    # RSLT belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"RSLT {rslt_id}")
                    broken_build_buy_errors.append(
                        f"RSLT {rslt_id}: {rslt_error}"
                    )

            # -------------------------
            # FTPT
            # -------------------------

            elif rtype == FTPT_TYPE:

                ftpt_id = f"0x{int(entry.Instance):016X}"

                ftpt_ok, ftpt_error = check_ftpt_health(
                    package,
                    entry,
                    resource_objects
                )

                if not ftpt_ok:
                    broken_ftpt_ids.append(ftpt_id)

                    # FTPT belongs to the BUILD_BUY corruption group.
                    broken_build_buy_ids.append(f"FTPT {ftpt_id}")
                    broken_build_buy_errors.append(
                        f"FTPT {ftpt_id}: {ftpt_error}"
                    )

            # -------------------------
            # CASP
            # -------------------------

            if DEBUG_RESOURCES:
                print(
                    "CHECK TYPE:",
                    rtype,
                    type(rtype),
                    hex(rtype) if isinstance(rtype, int) else ""
                )

            if rtype in CASP_TYPES:

                if DEBUG:
                    print(
                        "CASP MATCH FOUND:",
                        hex(rtype),
                        "INSTANCE:",
                        hex(int(entry.Instance))
                    )

                casp += 1

                casp_id = f"0x{int(entry.Instance):016X}"

                casp_instances.append(
                    casp_id
                )
                try:
                    # Parse the CASP itself separately from later classification.
                    # Only a real parser failure is a structural health failure.
                    try:
                        resource = WrapperDealer.GetResource(
                            0,
                            package,
                            entry
                        )
                    except Exception as e:
                        error_text = f"{type(e).__name__}: {e}"
                        broken_casp_ids.append(casp_id)
                        broken_casp_errors.append(
                            f"{casp_id}: {error_text}"
                        )
                        print(
                            f"CASP HEALTH FAILURE: {casp_id} | {error_text}"
                        )
                        resource = None
                        continue

                    # CASP DIAGNOSTIC: report successful parse details without
                    # adding any new health rule. Existing parser failures remain
                    # the permanent CASP structural failure condition.
                    try:
                        casp_byte_count = len(resource.AsBytes)
                    except Exception:
                        try:
                            casp_byte_count = int(entry.Filesize)
                        except Exception:
                            casp_byte_count = 0

                    print(
                        f"CASP CHECK: {casp_id} | "
                        f"Group=0x{int(entry.ResourceGroup):08X} | "
                        f"Bytes={casp_byte_count:,} | "
                        f"Wrapper={type(resource).__name__} | Status=OK"
                    )

                    if DEBUG:
                        print("RESOURCE TYPE:")
                        print(resource.GetType())
                        print("RESOURCE CLASS:")
                        print(resource.__class__)
                        print("RESOURCE DIR:")
                        print(dir(resource))

                        print("CASP RESOURCE OBJECT:")
                        print(resource)
                        print(type(resource))
                        print("CASP CONTENT FIELDS:")
                        print(resource.ContentFields)

                        print("CASP BYTES LENGTH:")
                        print(len(resource.AsBytes))

                    resource_objects.append(resource)


                    interesting = [
                        "Clothing",
                        "DataType",
                        "ClothingCategory",
                        "BodyType",
                        "BodyTypeFlags",
                        "CategoryFlags",
                        "OverlayPriority",
                        "AgeGender",
                        "CasPart1Index",
                        "CasPart2Index",
                        "LODInfo",
                    ]


                    for field in interesting:

                        try:

                            if DEBUG:
                                print(
                                    field,
                                    "=",
                                    resource[field]
                                )

                        except:

                            pass


                    clothing_text = str(resource["Clothing"])

                    datatype_text = str(
                        resource["DataType"]
                    ).lower()

                    try:
                        clothing_id = int(
                            clothing_text.split()[0],
                            16
                        )

                        cc_type = CAS_TYPE_NAMES.get(
                            clothing_id,
                            clothing_text
                        )

                    except:
                        cc_type = clothing_text

                    # -------------------------
                    # STORE CASP RECORD
                    # -------------------------

                    def safe_casp_field(resource, field):

                        try:
                            return str(resource[field])

                        except:
                            return ""


                    casp_records.append({
                        "CASP Instance": casp_id,
                        "Package": os.path.basename(filename),
                        "CC Type": cc_type,
                        "Data Type": datatype_text,
                        "Age Gender": safe_casp_field(resource, "AgeGender"),
                        "Category": safe_casp_field(resource, "ClothingCategory"),
                        "Body Type": safe_casp_field(resource, "BodyType"),
                        "Overlay Priority": safe_casp_field(resource, "OverlayPriority")
                    })

                    # -------------------------
                    # RELEASE CASP RESOURCE
                    # -------------------------

                    resource = None


                except Exception as e:

                    print("CASP CLASSIFICATION ERROR")

                    if DEBUG:
                        print(e)
            
            # -------------------------
            # EA DATABASE CHECK
            # -------------------------

            key_resource = (
                int(entry.ResourceType),
                int(entry.ResourceGroup),
                int(entry.Instance)
            )

            if key_resource in EA_LOOKUP:

                ea_tgi_matches += 1

                reference_methods = EA_HASH_LOOKUP.get(key_resource, {})
                exact_match = False
                method_failed = False
                method_attempted = False

                for hash_method, reference_hashes in reference_methods.items():
                    try:
                        resource_digest = _get_resource_sha256_by_method(
                            package,
                            entry,
                            hash_method
                        )
                    except Exception as e:
                        method_failed = True
                        if DEBUG:
                            print(
                                "EA HASH METHOD ERROR:",
                                hash_method,
                                "Type=0x%08X" % (int(entry.ResourceType) & 0xFFFFFFFF),
                                "Group=0x%08X" % (int(entry.ResourceGroup) & 0xFFFFFFFF),
                                "Instance=0x%016X" % (int(entry.Instance) & 0xFFFFFFFFFFFFFFFF),
                                str(e)
                            )
                        continue

                    method_attempted = True

                    if resource_digest in reference_hashes:
                        exact_match = True
                        break

                if exact_match:
                    ea_matches += 1
                    ea_exact_details.append(key_resource)

                    if rtype in CASP_TYPES:
                        ea_casp_matches += 1

                    elif rtype == GEOM_TYPE:
                        ea_geom_matches += 1

                    elif rtype == IMG_TYPE:
                        ea_img_matches += 1

                elif (
                    key_resource in EA_UNVERIFIABLE_TGIS
                    or method_failed
                    or not reference_methods
                    or not method_attempted
                ):
                    # Do not call a resource modified unless every EA content
                    # variant for this TGI was actually comparable.
                    ea_unverifiable_matches += 1
                    ea_unverifiable_details.append(key_resource)

                else:
                    ea_modified_matches += 1
                    ea_modified_details.append(key_resource)

            else:
                non_ea_type = int(entry.ResourceType) & 0xFFFFFFFF
                ea_non_ea_type_counts[non_ea_type] = ea_non_ea_type_counts.get(non_ea_type, 0) + 1

        # -------------------------
        # FINAL CHECKS (outside resource loop)
        # -------------------------

        package_format = detect_package_format(filename)


        # -------------------------
        # PACKAGE CASP TYPE
        # -------------------------

        known_casp_types = set(CAS_TYPE_NAMES.values())

        package_casp_types = {
            record["CC Type"]
            for record in casp_records
            if record.get("CC Type") in known_casp_types
        }

        if len(package_casp_types) == 1:

            cc_type = next(iter(package_casp_types))

        elif len(package_casp_types) > 1:

            cc_type = "Mixed CAS"

        elif casp_records:

            cc_type = "Unknown"


        # -------------------------
        # CAREER CLASSIFICATION TEST
        # -------------------------

        if (
            cc_type == "Unknown"
            and career_marker
        ):
            cc_type = "Career"
            print(
                "CAREER CLASSIFICATION:",
                os.path.basename(filename),
                "| Result= Career"
            )


        # -------------------------
        # TERRAIN PAINT CLASSIFICATION TEST
        # -------------------------

        if (
            cc_type == "Unknown"
            and casp == 0
            and terrain_paint_candidate
            and terrain_paint_ctpt_count == 2
            and terrain_paint_stbl_count > 0
            and terrain_paint_img_count > 0
            and not broken_ctpt_ids
            and not broken_stbl_ids
            and not broken_img_ids
            and not broken_icon_ids
            and not broken_thum_ids
            and not broken_nmap_ids
            and not broken_manifest_ids
        ):
            cc_type = "Terrain Paint"
            print(
                "TERRAIN PAINT CLASSIFICATION:",
                os.path.basename(filename),
                "| CTPT=", terrain_paint_ctpt_count,
                "| STBL=", terrain_paint_stbl_count,
                "| IMG=", terrain_paint_img_count,
                "| ICON=", terrain_paint_icon_count,
                "| NMAP=", terrain_paint_nmap_count,
                "| TERRAIN_THUM=", terrain_paint_brush_thum_count,
                "| Result= Terrain Paint"
            )

        # -------------------------
        # COLLECTION FILE CLASSIFICATION TEST
        # -------------------------
        # Dedicated structural rule: after ignoring optional MANIFEST
        # metadata, every substantive resource is COLL (0x0C07456D).
        if (
            cc_type == "Unknown"
            and casp == 0
            and collection_file_candidate
            and collection_file_count > 0
            and not broken_manifest_ids
            and not broken_package_storage_ids
        ):
            cc_type = "Collection File"
            print(
                "COLLECTION FILE CLASSIFICATION:",
                os.path.basename(filename),
                "| COLL=",
                collection_file_count,
                "| Result= Collection File"
            )

        # -------------------------
        # BUILD / BUY DETECTION
        # -------------------------

        filename_lower = os.path.basename(filename).lower()

        raw_build_name = os.path.splitext(
            os.path.basename(filename)
        )[0]

        build_name_parts = []

        for index, char in enumerate(raw_build_name):

            if (
                index > 0
                and char.isupper()
                and raw_build_name[index - 1].islower()
            ):
                build_name_parts.append(" ")

            build_name_parts.append(char)

        build_name = "".join(
            char.lower() if char.isalnum() else " "
            for char in "".join(build_name_parts)
        )

        build_name_tokens = set(
            build_name.split()
        )

        if (
            cc_type == "Unknown"
            and casp == 0
        ):

            if (
                "pattern" in build_name_tokens
                or "patterns" in build_name_tokens
            ):

                cc_type = "Build/Buy"

            elif (
                "walllight" in build_name_tokens
                or (
                    "wall" in build_name_tokens
                    and "light" in build_name_tokens
                )
            ):

                cc_type = "Build/Buy"

                if DEBUG:
                    print(
                        "BUILD DETECTED:",
                        filename,
                        cc_type
                    )

        if DEBUG:
            print("PACKAGE FORMAT:", package_format)

        # -------------------------
        # OBJECT PRIORITY
        # -------------------------

        if (
            cc_type == "Unknown"
            and casp == 0
            and vpxy > 0
            and fbln == 0
            and face == 0
        ):
            cc_type = "Object"

        # -------------------------
        # SLIDER DETECTION
        # -------------------------

        if (
            cc_type == "Unknown"
            and casp == 0
            and (
                fbln > 0
                or face > 0
            )
        ):
            cc_type = "Morph Slider"
        # -------------------------
        # POSSIBLE SKIN CHECK
        # -------------------------

        possible_skin = False
        skin_type = ""

        if (
            cc_type == "Unknown"
            and casp == 0
            and geom == 0
            and valid_texture_count >= 20
        ):

            raw_possible_skin_name = os.path.splitext(
                os.path.basename(filename)
            )[0]

            filename_lower = raw_possible_skin_name.lower()

            possible_skin_parts = []

            for index, char in enumerate(raw_possible_skin_name):

                if (
                    index > 0
                    and char.isupper()
                    and raw_possible_skin_name[index - 1].islower()
                ):
                    possible_skin_parts.append(" ")

                possible_skin_parts.append(char)

            possible_skin_name = "".join(
                char.lower() if char.isalnum() else " "
                for char in "".join(possible_skin_parts)
            )

            possible_skin_tokens = set(
                possible_skin_name.split()
            )

            if not any(word in possible_skin_tokens for word in [
                "wall",
                "wallpaper",
                "floor",
                "tile",
                "pattern",
                "poster",
                "art",
                "painting",
                "light",
                "merged",
                "collection"
            ]):

                possible_skin = True

                if (
                    "nondefault" in filename_lower
                    or "non-default" in filename_lower
                    or "nd" in possible_skin_tokens
                    or "_nd" in filename_lower
                    or "-nd" in filename_lower
                ):

                    skin_type = "Non-Default"


                elif any(word in filename_lower for word in [
                    "defaultskin",
                    "default",
                    "defskin",
                    "def_skin",
                    "def_",
                    "_def",
                    "-def",
                    "replacement",
                    "_dr_",
                    "-dr-",
                    "dr_",
                    "_dr"
                ]):

                    skin_type = "Default Replacement"


                else:

                    skin_type = "Unknown"

        # -------------------------
        # HQ TEXTURE CHECK
        # -------------------------

        if fourk > 0:

            hq_warning = "4K TEXTURE"


        elif two_k > 0:

            hq_warning = "2048 HIGH"


        else:

            hq_warning = "Normal"



        # -------------------------
        # POLY WARNING
        # -------------------------

        poly_warning = get_poly_warning(
            highest_polys,
            cc_type
        )



        # -------------------------
        # SPECIAL CC DETECTION

        special_type = detect_special_cc(
            filename,
            casp,
            geom,
            presets,
            cc_type,
            objd,
            objk,
            modl,
            mlod,
            simo,
            famd
        )
        
        if special_type:

            cc_type = special_type

            if special_type.endswith("Preset"):
                preset_type = special_type

            possible_skin = False

            if cc_type != "Skin":
                skin_type = ""


        # -------------------------
        # SCRIPTED OBJECT CLASSIFICATION TEST
        # -------------------------
        # Embedded Script now uses direct structural ownership evidence:
        # a healthy OBJK scriptClass must EXACTLY match a TypeDef declared by
        # one of this package's successfully parsed S3SA/.NET assemblies.
        # This distinguishes a real scripted object from a script package that
        # merely references ordinary EA object classes or carries object assets.
        #
        # External Script has no local S3SA to prove ownership, so it keeps a
        # deliberately conservative creator-namespace heuristic.  Importantly,
        # a lowercase final EA class name is not enough evidence.
        has_build_buy_object_resources = (
            objd > 0
            or objk > 0
            or modl > 0
            or mlod > 0
        )

        embedded_script_class_matches = [
            class_name
            for class_name in objk_script_class_names
            if class_name in s3sa_defined_type_names
        ]
        embedded_object_namespace_matches = [
            class_name
            for class_name in embedded_script_class_matches
            if (
                class_name.startswith("Sims3.Gameplay.Objects.")
                or class_name.startswith("Sims3.Store.Objects.")
            )
        ]
        # Some creator-authored scripted objects use their own root namespace
        # rather than Sims3.Gameplay.Objects.*.  Permit that only for a small
        # local assembly, where an exact OBJK->TypeDef match is much more likely
        # to describe the package's primary object rather than one ancillary
        # helper buried inside a large framework (for example NRaas Careers).
        small_creator_assembly_match = (
            bool(embedded_script_class_matches)
            and 0 < len(s3sa_defined_type_names) <= 100
        )

        embedded_scripted_object = (
            cc_type in ("Build/Buy", "Object")
            and script_mod_marker
            and objd > 0
            and (
                bool(embedded_object_namespace_matches)
                or small_creator_assembly_match
            )
            and casp == 0
            and has_build_buy_object_resources
        )

        external_scripted_object = (
            cc_type in ("Build/Buy", "Object")
            and not script_mod_marker
            and objd > 0
            and external_script_class_marker
            and casp == 0
            and has_build_buy_object_resources
        )

        script_mod_with_object_resources = (
            cc_type in ("Build/Buy", "Object")
            and script_mod_marker
            and not embedded_scripted_object
            and casp == 0
            and has_build_buy_object_resources
        )

        if embedded_scripted_object:
            cc_type = "Scripted Object"
            scripted_object_subtype = "Embedded Script"
            print(
                "SCRIPTED OBJECT CLASSIFICATION:",
                os.path.basename(filename),
                "| Subtype= Embedded Script",
                "| MatchedDefinedType=", "; ".join(embedded_script_class_matches[:3]),
                "| DefinedTypes=", len(s3sa_defined_type_names),
                "| S3SA=Valid",
                "| Tuning=", int(bool(tuning_marker)),
                "| ITUN=", int(bool(interaction_tuning_marker)),
                "| OBJD=", objd,
                "| OBJK=", objk,
                "| MODL=", modl,
                "| MLOD=", mlod,
                "| Result= Scripted Object"
            )

        elif external_scripted_object:
            cc_type = "Scripted Object"
            scripted_object_subtype = "External Script"
            print(
                "SCRIPTED OBJECT CLASSIFICATION:",
                os.path.basename(filename),
                "| Subtype= External Script",
                "| CustomClass=", "; ".join(external_script_class_names[:3]),
                "| S3SA=None",
                "| OBJD=", objd,
                "| OBJK=", objk,
                "| MODL=", modl,
                "| MLOD=", mlod,
                "| Result= Scripted Object"
            )

        elif script_mod_with_object_resources:
            cc_type = "Script Mod"
            scripted_object_subtype = ""
            print(
                "SCRIPT MOD WITH OBJECT RESOURCES CLASSIFICATION:",
                os.path.basename(filename),
                "| S3SA=Valid",
                "| OBJKScriptClass=",
                "; ".join(objk_script_class_names[:3]) or "None",
                "| DefinedTypeMatches=", len(embedded_script_class_matches),
                "| DefinedTypes=", len(s3sa_defined_type_names),
                "| OBJD=", objd,
                "| OBJK=", objk,
                "| MODL=", modl,
                "| MLOD=", mlod,
                "| Result= Script Mod"
            )


        # -------------------------
        # SIME + SIMO SAVED-SIM CLASSIFICATION TEST
        # -------------------------
        # Validated on the previously unidentified exported-Sim package. Keep
        # this conservative: require both saved-Sim evidence types and reject
        # CASP / Build-Buy model structures. Health checks remain independent.
        sime_simo_saved_sim = (
            cc_type == "Unknown"
            and sime > 0
            and simo > 0
            and casp == 0
            and objd == 0
            and objk == 0
            and modl == 0
            and mlod == 0
        )

        if sime_simo_saved_sim:
            print(
                "SIM PACKAGE CLASSIFICATION TEST:",
                os.path.basename(filename),
                "| Unknown -> Custom Sim",
                "| SIME=", sime,
                "| SIMO=", simo,
                "| Evidence=SIME+SIMO+NoCASP+NoBuildBuy"
            )
            cc_type = "Custom Sim"

        # -------------------------
        # STRUCTURAL SIM PACKAGE CLASSIFICATION (PERMANENT)
        # -------------------------
        # Validated on a 509-package diagnostic batch (9 intended creator-Sim
        # matches, 500 negatives, 0 false positives) and a targeted routing
        # test. Require Mixed CAS + SIMO + FACE and explicitly reject sliders.
        # Structural health checks remain independent and can still mark a
        # classified Sim package BROKEN.
        structural_sim_package = (
            cc_type == "Mixed CAS"
            and simo > 0
            and face > 0
            and slider_count == 0
        )

        if structural_sim_package:
            print(
                "SIM PACKAGE CLASSIFICATION:",
                os.path.basename(filename),
                "| Mixed CAS -> Custom Sim",
                "| Evidence=SIMO+FACE+MixedCAS+NoSlider"
            )
            cc_type = "Custom Sim"

        # -------------------------
        # RANDOM EVENT MODULE CLASSIFICATION TEST
        # -------------------------
        # More specific than the generic Current_Tuning marker, but still
        # lower priority than established CAS / Build-Buy / Object / Slider /
        # Skin / Preset / Career / Custom Sim classifications.
        if (
            cc_type == "Unknown"
            and random_event_module_marker
        ):
            cc_type = "Random Event Module"
            print(
                "RANDOM EVENT MODULE CLASSIFICATION:",
                os.path.basename(filename),
                "| Result= Random Event Module"
            )

        # -------------------------
        # INTERACTION TUNING CLASSIFICATION TEST
        # -------------------------
        # Resource-specific Sims 3 ITUN classification. Keep this below the
        # more specific Random Event Module rule and below all established
        # CAS / Build-Buy / Object / Slider / Skin / Preset / Career / Custom
        # Sim classifications by requiring the package to still be Unknown.
        # A package that also has the ordinary Current_Tuning XML marker keeps
        # the broader Tuning Mod classification instead of being stolen by ITUN.
        if (
            cc_type == "Unknown"
            and interaction_tuning_marker
            and not tuning_marker
        ):
            cc_type = "Interaction Tuning"
            print(
                "INTERACTION TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| Result= Interaction Tuning"
            )

        # -------------------------
        # OPPORTUNITY TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + XML
        # - at least one healthy NMAP and one XML are present
        # - every XML passed permanent XML health
        # - every XML matches the established Opportunities_* block schema
        if (
            cc_type == "Unknown"
            and opportunity_tuning_candidate
            and opportunity_tuning_nmap_count > 0
            and opportunity_tuning_xml_count > 0
            and opportunity_tuning_valid_xml_count == opportunity_tuning_xml_count
            and not broken_nmap_ids
            and not broken_xml_ids
        ):
            cc_type = "Opportunity Tuning"
            print(
                "OPPORTUNITY TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                opportunity_tuning_nmap_count,
                "| XML=",
                opportunity_tuning_xml_count,
                "| ValidOpportunityXML=",
                opportunity_tuning_valid_xml_count,
                "| Result= Opportunity Tuning"
            )

        # -------------------------
        # JAZZ MOD CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule grounded in the observed package:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, every resource is NMAP or JAZZ
        # - at least one NMAP and one JAZZ are present
        # - existing permanent NMAP/JAZZ health checks passed
        if (
            cc_type == "Unknown"
            and jazz_mod_candidate
            and jazz_mod_nmap_count > 0
            and jazz_mod_jazz_count > 0
            and not broken_nmap_ids
            and not broken_jazz_ids
        ):
            cc_type = "Jazz Mod"
            print(
                "JAZZ MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                jazz_mod_nmap_count,
                "| JAZZ=",
                jazz_mod_jazz_count,
                "| Result= Jazz Mod"
            )

        # -------------------------
        # TEXTURE MOD CLASSIFICATION TEST
        # -------------------------
        # Conservative pure-resource rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, every substantive resource is NMAP or IMG
        # - at least one healthy NMAP is present
        # - at least one IMG resource is present
        # - every IMG was recognized as a valid DDS by existing health logic
        # - no NMAP or IMG health failures occurred
        #
        # "Texture Mod" is intentionally broader than "Texture Override":
        # structure proves this is a texture-only resource package, but does
        # not by itself prove whether the textures replace EA, another mod,
        # or are consumed by another custom package.
        if (
            cc_type == "Unknown"
            and texture_mod_candidate
            and texture_mod_nmap_count > 0
            and texture_mod_img_count >= 1
            and valid_texture_count == texture_mod_img_count
            and not broken_nmap_ids
            and not broken_img_ids
        ):
            cc_type = "Texture Mod"
            print(
                "TEXTURE MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                texture_mod_nmap_count,
                "| IMG=",
                texture_mod_img_count,
                "| ValidDDS=",
                valid_texture_count,
                "| Result= Texture Mod"
            )

        # -------------------------
        # IMG-ONLY TEXTURE MOD CLASSIFICATION TEST
        # -------------------------
        # Some valid texture resource packs contain only DDS IMG resources
        # and no NMAP.  Keep this deliberately broad as Texture Mod rather
        # than guessing a semantic subtype from filenames.  Preserve the
        # existing Possible Skin path for large skin-like IMG sets.
        if (
            cc_type == "Unknown"
            and texture_mod_candidate
            and texture_mod_nmap_count == 0
            and texture_mod_img_count > 0
            and valid_texture_count == texture_mod_img_count
            and not broken_img_ids
            and not possible_skin
        ):
            cc_type = "Texture Mod"
            print(
                "IMG-ONLY TEXTURE MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| IMG=",
                texture_mod_img_count,
                "| ValidDDS=",
                valid_texture_count,
                "| Result= Texture Mod"
            )

        # -------------------------
        # COOKING TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow multi-table rule grounded in the inspected overhaul:
        # - package remains Unknown after stronger established classifiers
        # - only the observed cooking/localization resource families exist
        # - every normal XML passed permanent XML health AND matched a
        #   recognized cooking table schema
        # - the dedicated Store Ingredients resource parses completely and
        #   matches Ingredients_store
        # - the package carries the major master tables plus a large family
        #   of individual recipe-process tables
        # - NMAP/STBL/PNG resources must have no permanent health failures
        if (
            cc_type == "Unknown"
            and cooking_tuning_candidate
            and cooking_tuning_nmap_count > 0
            and cooking_tuning_xml_count >= 40
            and cooking_tuning_stbl_count > 0
            and cooking_tuning_png_count > 0
            and cooking_tuning_store_ingredients_count == 1
            and cooking_tuning_valid_xml_count == cooking_tuning_xml_count
            and cooking_tuning_valid_store_ingredients_count == 1
            and cooking_tuning_recipe_process_count >= 30
            and cooking_tuning_recipe_master_count >= 1
            and cooking_tuning_ingredients_count >= 1
            and cooking_tuning_plants_count >= 1
            and cooking_tuning_bars_count >= 1
            and cooking_tuning_shop_tabs_count >= 1
            and not broken_nmap_ids
            and not broken_xml_ids
            and not broken_stbl_ids
            and not broken_imag_ids
        ):
            cc_type = "Cooking Tuning"
            print(
                "COOKING TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                cooking_tuning_nmap_count,
                "| XML=",
                cooking_tuning_xml_count,
                "| STBL=",
                cooking_tuning_stbl_count,
                "| PNG=",
                cooking_tuning_png_count,
                "| StoreIngredients=",
                cooking_tuning_store_ingredients_count,
                "| RecipeProcess=",
                cooking_tuning_recipe_process_count,
                "| Result= Cooking Tuning"
            )

        # -------------------------
        # FISHING TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow full-table structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, only NMAP + XML + optional STBL exist
        # - at least one healthy NMAP and one XML are present
        # - every XML passed permanent XML health
        # - every XML matches the full Fishing multi-table schema
        # - any STBL localization resources also passed permanent STBL health
        if (
            cc_type == "Unknown"
            and fishing_tuning_candidate
            and fishing_tuning_nmap_count > 0
            and fishing_tuning_xml_count > 0
            and fishing_tuning_valid_xml_count == fishing_tuning_xml_count
            and not broken_nmap_ids
            and not broken_xml_ids
            and not broken_stbl_ids
        ):
            cc_type = "Fishing Tuning"
            print(
                "FISHING TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                fishing_tuning_nmap_count,
                "| XML=",
                fishing_tuning_xml_count,
                "| STBL=",
                fishing_tuning_stbl_count,
                "| ValidFishingXML=",
                fishing_tuning_valid_xml_count,
                "| Result= Fishing Tuning"
            )

        # -------------------------
        # SKILL TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow full-table structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + XML
        # - at least one healthy NMAP and one XML are present
        # - every XML passed permanent XML health
        # - every XML matches the full Skills / SkillList table schema
        if (
            cc_type == "Unknown"
            and skill_tuning_candidate
            and skill_tuning_nmap_count > 0
            and skill_tuning_xml_count > 0
            and skill_tuning_valid_xml_count == skill_tuning_xml_count
            and not broken_nmap_ids
            and not broken_xml_ids
        ):
            cc_type = "Skill Tuning"
            print(
                "SKILL TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                skill_tuning_nmap_count,
                "| XML=",
                skill_tuning_xml_count,
                "| ValidSkillsXML=",
                skill_tuning_valid_xml_count,
                "| Result= Skill Tuning"
            )

        # -------------------------
        # SOCIAL TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow full-table structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + XML
        # - at least one healthy NMAP and one XML are present
        # - every XML passed permanent XML health
        # - every XML matches the full Social / Action data schema
        if (
            cc_type == "Unknown"
            and social_tuning_candidate
            and social_tuning_nmap_count > 0
            and social_tuning_xml_count > 0
            and social_tuning_valid_xml_count == social_tuning_xml_count
            and not broken_nmap_ids
            and not broken_xml_ids
        ):
            cc_type = "Social Tuning"
            print(
                "SOCIAL TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                social_tuning_nmap_count,
                "| XML=",
                social_tuning_xml_count,
                "| ValidSocialXML=",
                social_tuning_valid_xml_count,
                "| Result= Social Tuning"
            )

        # -------------------------
        # BUFF TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow full-table structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + XML
        # - at least one healthy NMAP and one XML are present
        # - every XML passed permanent XML health
        # - every XML matches the full buffs / BuffList table schema
        if (
            cc_type == "Unknown"
            and buff_tuning_candidate
            and buff_tuning_nmap_count > 0
            and buff_tuning_xml_count > 0
            and buff_tuning_valid_xml_count == buff_tuning_xml_count
            and not broken_nmap_ids
            and not broken_xml_ids
        ):
            cc_type = "Buff Tuning"
            print(
                "BUFF TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                buff_tuning_nmap_count,
                "| XML=",
                buff_tuning_xml_count,
                "| ValidBuffsXML=",
                buff_tuning_valid_xml_count,
                "| Result= Buff Tuning"
            )

        # -------------------------
        # BOARDING SCHOOL TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + XML
        # - at least one healthy NMAP and one XML are present
        # - every XML passed permanent XML health
        # - every XML matches the BoardingSchool full-table schema
        if (
            cc_type == "Unknown"
            and boarding_school_tuning_candidate
            and boarding_school_tuning_nmap_count > 0
            and boarding_school_tuning_xml_count > 0
            and boarding_school_tuning_valid_xml_count
                == boarding_school_tuning_xml_count
            and not broken_nmap_ids
            and not broken_xml_ids
        ):
            cc_type = "Boarding School Tuning"
            print(
                "BOARDING SCHOOL TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                boarding_school_tuning_nmap_count,
                "| XML=",
                boarding_school_tuning_xml_count,
                "| ValidBoardingSchoolXML=",
                boarding_school_tuning_valid_xml_count,
                "| Result= Boarding School Tuning"
            )

        # -------------------------
        # SCORING TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + XML
        # - at least one healthy NMAP and one XML are present
        # - every XML passed permanent XML health
        # - every XML matches the Scoring / Methods framework
        if (
            cc_type == "Unknown"
            and scoring_tuning_candidate
            and scoring_tuning_nmap_count > 0
            and scoring_tuning_xml_count > 0
            and scoring_tuning_valid_xml_count == scoring_tuning_xml_count
            and not broken_nmap_ids
            and not broken_xml_ids
        ):
            cc_type = "Scoring Tuning"
            print(
                "SCORING TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                scoring_tuning_nmap_count,
                "| XML=",
                scoring_tuning_xml_count,
                "| ValidScoringXML=",
                scoring_tuning_valid_xml_count,
                "| Result= Scoring Tuning"
            )

        # -------------------------
        # DREAMS AND PROMISES TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, every substantive resource is XML
        # - at least one XML is present
        # - every XML passed permanent XML health
        # - every XML matches the DreamsAndPromisesNodes / Primitives schema
        if (
            cc_type == "Unknown"
            and dreams_and_promises_tuning_candidate
            and dreams_and_promises_tuning_xml_count > 0
            and dreams_and_promises_tuning_valid_xml_count
                == dreams_and_promises_tuning_xml_count
            and not broken_xml_ids
        ):
            cc_type = "Dreams and Promises Tuning"
            print(
                "DREAMS AND PROMISES TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| XML=",
                dreams_and_promises_tuning_xml_count,
                "| ValidDreamsAndPromisesXML=",
                dreams_and_promises_tuning_valid_xml_count,
                "| Result= Dreams and Promises Tuning"
            )

        # -------------------------
        # DUMPSTER DIVING TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + XML
        # - at least one healthy NMAP and one XML are present
        # - every XML passed permanent XML health
        # - every XML matches the DumpsterDiveData / RewardObjects schema
        if (
            cc_type == "Unknown"
            and dumpster_diving_tuning_candidate
            and dumpster_diving_tuning_nmap_count > 0
            and dumpster_diving_tuning_xml_count > 0
            and dumpster_diving_tuning_valid_xml_count
                == dumpster_diving_tuning_xml_count
            and not broken_nmap_ids
            and not broken_xml_ids
        ):
            cc_type = "Dumpster Diving Tuning"
            print(
                "DUMPSTER DIVING TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                dumpster_diving_tuning_nmap_count,
                "| XML=",
                dumpster_diving_tuning_xml_count,
                "| ValidDumpsterDiveDataXML=",
                dumpster_diving_tuning_valid_xml_count,
                "| Result= Dumpster Diving Tuning"
            )

        # -------------------------
        # BOOK TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + XML
        # - at least one healthy NMAP and one XML are present
        # - every XML passed permanent XML health
        # - every XML matches the established Books table schema
        if (
            cc_type == "Unknown"
            and book_tuning_candidate
            and book_tuning_nmap_count > 0
            and book_tuning_xml_count > 0
            and book_tuning_valid_xml_count == book_tuning_xml_count
            and not broken_nmap_ids
            and not broken_xml_ids
        ):
            cc_type = "Book Tuning"
            print(
                "BOOK TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                book_tuning_nmap_count,
                "| XML=",
                book_tuning_xml_count,
                "| ValidBooksXML=",
                book_tuning_valid_xml_count,
                "| Result= Book Tuning"
            )

        # -------------------------
        # CAREER TUNING CLASSIFICATION PERMANENT
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + XML
        # - at least one healthy NMAP and one XML are present
        # - every XML passed permanent XML health
        # - every XML matches either the standard Careers/CareerList schema
        #   or one of the validated ActiveCareers / SkillBasedCareers schemas
        if (
            cc_type == "Unknown"
            and career_tuning_candidate
            and career_tuning_nmap_count > 0
            and career_tuning_xml_count > 0
            and career_tuning_valid_xml_count == career_tuning_xml_count
            and not broken_nmap_ids
            and not broken_xml_ids
        ):
            cc_type = "Career Tuning"
            print(
                "CAREER TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                career_tuning_nmap_count,
                "| XML=",
                career_tuning_xml_count,
                "| ValidCareersXML=",
                career_tuning_valid_xml_count,
                "| Result= Career Tuning"
            )

        # -------------------------
        # ACADEMIC DEGREE TUNING CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, every substantive resource is XML
        # - at least one XML exists
        # - every XML passed permanent XML health
        # - every XML matches the established AcademicDegrees schema
        if (
            cc_type == "Unknown"
            and academic_degree_tuning_candidate
            and academic_degree_tuning_xml_count > 0
            and academic_degree_tuning_valid_xml_count == academic_degree_tuning_xml_count
            and not broken_xml_ids
        ):
            cc_type = "Academic Degree Tuning"
            print(
                "ACADEMIC DEGREE TUNING CLASSIFICATION:",
                os.path.basename(filename),
                "| XML=",
                academic_degree_tuning_xml_count,
                "| ValidAcademicDegreesXML=",
                academic_degree_tuning_valid_xml_count,
                "| Result= Academic Degree Tuning"
            )

        # -------------------------
        # SCRIPT MOD WITH EMBEDDED CAS CLASSIFICATION TEST
        # -------------------------
        # Some pure script mods bundle CASP resources for NPC/service outfits.
        # The package-level CAS aggregation above can therefore identify a
        # legitimate Hair/Body/etc. or Mixed CAS payload even though the
        # package itself is fundamentally a script mod.  Promote only when a
        # successfully parsed S3SA is accompanied by a real tuning marker and
        # there are no Build/Buy mesh/object resources.  CASP Master records
        # remain unchanged, so the embedded CAS types are still preserved.
        embedded_cas_script_mod = (
            (cc_type in known_casp_types or cc_type == "Mixed CAS")
            and script_mod_marker
            and (tuning_marker or interaction_tuning_marker)
            and casp > 0
            and objd == 0
            and objk == 0
            and modl == 0
            and mlod == 0
        )

        if embedded_cas_script_mod:
            embedded_cas_type = cc_type
            cc_type = "Script Mod"
            print(
                "SCRIPT MOD WITH EMBEDDED CAS CLASSIFICATION:",
                os.path.basename(filename),
                "| EmbeddedCAS=", embedded_cas_type,
                "| CASP=", casp,
                "| Evidence=ValidS3SA+Tuning/ITUN+NoBuildBuy",
                "| Result= Script Mod"
            )

        # -------------------------
        # TUNING MOD CLASSIFICATION TEST
        # -------------------------
        # Lower priority than established CAS / Build-Buy / Object / Slider /
        # Skin / Preset / Career classifications. Only otherwise-unidentified
        # packages can be promoted by this XML tuning marker.
        if (
            cc_type == "Unknown"
            and tuning_marker
        ):
            cc_type = "Tuning Mod"
            print(
                "TUNING MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| Result= Tuning Mod"
            )

        # -------------------------
        # SCRIPT MOD CLASSIFICATION TEST
        # -------------------------
        # S3SA (0x073FAA07) is the Sims 3 compiled script assembly resource.
        # Only classify a package that is still Unknown after all established
        # CAS / Build-Buy / Career / Random Event / Interaction Tuning /
        # Current_Tuning classifications, and only when the existing permanent
        # S3SA health parser successfully loaded at least one S3SA resource.
        if (
            cc_type == "Unknown"
            and script_mod_marker
        ):
            cc_type = "Script Mod"
            print(
                "SCRIPT MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| Result= Script Mod"
            )

        # -------------------------
        # UI LAYOUT MOD CLASSIFICATION TEST
        # -------------------------
        # LAYO (0x025C95B6) is the Sims 3 UI Layout resource type. Classify only
        # packages that are still Unknown after all established structural
        # classifications, and only when the existing permanent LAYO health
        # validator accepted at least one LAYO resource. This keeps mixed
        # tuning/script packages under their stronger existing classification.
        if (
            cc_type == "Unknown"
            and ui_layout_marker
        ):
            cc_type = "UI Layout Mod"
            print(
                "UI LAYOUT MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| Result= UI Layout Mod"
            )

        # -------------------------
        # HAIR COLOR PRESET CLASSIFICATION TEST
        # -------------------------
        # HINF (0x063261DA) is the Sims 3 Hair Information / haircolor
        # resource type. Classify only packages still Unknown after all
        # established stronger classifications, and only when the existing
        # permanent HINF health validator accepted at least one HINF resource.
        if (
            cc_type == "Unknown"
            and hair_color_preset_marker
        ):
            cc_type = "Hair Color Preset"
            print(
                "HAIR COLOR PRESET CLASSIFICATION:",
                os.path.basename(filename),
                "| Result= Hair Color Preset"
            )

        # -------------------------
        # THUMBNAIL MOD CLASSIFICATION TEST
        # -------------------------
        # V2 rule: after ignoring MANIFEST/NMAP auxiliary metadata,
        # every substantive resource must be one of the established Sims 3
        # thumbnail resource types 0x626F60CD/0x626F60CE/0xAD366F95.
        # Both types already have permanent structural image health checks.
        # Keep this low priority and require zero thumbnail-health failures.
        if (
            cc_type == "Unknown"
            and thumbnail_mod_candidate
            and thumbnail_mod_resource_count > 0
            and not broken_thum_ids
            and not broken_626f60ce_ids
        ):
            cc_type = "Thumbnail Mod"
            print(
                "THUMBNAIL MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| Resources=",
                thumbnail_mod_resource_count,
                "| Result= Thumbnail Mod"
            )

        # -------------------------
        # FONT MOD CLASSIFICATION TEST
        # -------------------------
        # Narrow v1 structural rule:
        # - package is still Unknown after stronger classifiers
        # - at least one valid NMAP plus at least one _CSS and one _TTF
        # - after ignoring MANIFEST, no other resource types are present
        # - the existing permanent NMAP health check must pass
        if (
            cc_type == "Unknown"
            and font_mod_candidate
            and font_mod_nmap_count > 0
            and font_mod_css_count > 0
            and font_mod_ttf_count > 0
            and not broken_nmap_ids
        ):
            cc_type = "Font Mod"
            print(
                "FONT MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                font_mod_nmap_count,
                "| CSS=",
                font_mod_css_count,
                "| TTF=",
                font_mod_ttf_count,
                "| Result= Font Mod"
            )

        # -------------------------
        # LOT CLASSIFICATION TEST
        # -------------------------
        # Require the dedicated LDES lot-description marker plus at least one
        # dedicated lot preview image. Keep it low priority so stronger
        # established classifications still win.
        if (
            cc_type == "Unknown"
            and lot_ldes_marker
            and lot_preview_count > 0
        ):
            cc_type = "Lot"
            print(
                "LOT CLASSIFICATION:",
                os.path.basename(filename),
                "| LDES=1",
                "| LotPreviews=",
                lot_preview_count,
                "| Result= Lot"
            )

        # -------------------------
        # POSE CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + CLIP
        # - every CLIP exposes framecount == 1 (static pose)
        # - existing form: at least one healthy NMAP + one or more healthy CLIPs
        # - V2 standalone form: no NMAP and exactly one healthy one-frame CLIP
        if (
            cc_type == "Unknown"
            and pose_candidate
            and pose_clip_count > 0
            and pose_one_frame_clip_count == pose_clip_count
            and (
                pose_nmap_count > 0
                or (
                    pose_nmap_count == 0
                    and pose_clip_count == 1
                )
            )
            and not broken_nmap_ids
            and not broken_clip_ids
        ):
            cc_type = "Pose"
            print(
                "POSE CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                pose_nmap_count,
                "| CLIP=",
                pose_clip_count,
                "| OneFrameCLIP=",
                pose_one_frame_clip_count,
                "| Result= Pose"
            )

        # -------------------------
        # ANIMATION CLASSIFICATION PERMANENT
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, package contains only NMAP + CLIP
        # - at least one healthy NMAP and one healthy CLIP are present
        # - every CLIP exposes framecount > 1
        # This deliberately does not broaden the one-frame Pose rule.
        if (
            cc_type == "Unknown"
            and pose_candidate
            and pose_nmap_count > 0
            and pose_clip_count > 0
            and pose_multi_frame_clip_count == pose_clip_count
            and not broken_nmap_ids
            and not broken_clip_ids
        ):
            cc_type = "Animation"
            print(
                "ANIMATION CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                pose_nmap_count,
                "| CLIP=",
                pose_clip_count,
                "| MultiFrameCLIP=",
                pose_multi_frame_clip_count,
                "| Result= Animation"
            )

        # -------------------------
        # STRING TABLE MOD CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, only optional NMAP + STBL exist
        # - at least one STBL exists
        # - all STBL/NMAP resources passed their permanent health checks
        if (
            cc_type == "Unknown"
            and string_table_mod_candidate
            and string_table_mod_stbl_count > 0
            and not broken_stbl_ids
            and not broken_nmap_ids
        ):
            cc_type = "String Table Mod"
            print(
                "STRING TABLE MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| NMAP=",
                string_table_mod_nmap_count,
                "| STBL=",
                string_table_mod_stbl_count,
                "| Result= String Table Mod"
            )

        # -------------------------
        # FOOTPRINT MOD CLASSIFICATION TEST
        # -------------------------
        # Narrow structural rule:
        # - package is still Unknown after stronger classifiers
        # - after ignoring MANIFEST, every substantive resource is FTPT
        # - at least one FTPT exists
        # - all FTPT resources passed the permanent health checker
        if (
            cc_type == "Unknown"
            and footprint_mod_candidate
            and footprint_mod_ftpt_count > 0
            and not broken_ftpt_ids
        ):
            cc_type = "Footprint Mod"
            print(
                "FOOTPRINT MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| FTPT=",
                footprint_mod_ftpt_count,
                "| Result= Footprint Mod"
            )

        # -------------------------
        # INI MOD CLASSIFICATION TEST
        # -------------------------
        # _INI (0x1F886EAD) is the Sims 3 configuration resource type used
        # by camera, lighting, weather, water/world and other INI-based mods.
        # Keep it low priority so established structural classifications win.
        if (
            cc_type == "Unknown"
            and ini_mod_marker
        ):
            cc_type = "INI Mod"
            print(
                "INI MOD CLASSIFICATION:",
                os.path.basename(filename),
                "| Result= INI Mod"
            )

        # -------------------------
        # FINAL SKIN TYPE
        # -------------------------

        if cc_type == "Skin" and not skin_type:

            raw_final_skin_name = os.path.splitext(
                os.path.basename(filename)
            )[0]

            filename_lower = raw_final_skin_name.lower()

            final_skin_parts = []

            for index, char in enumerate(raw_final_skin_name):

                if (
                    index > 0
                    and char.isupper()
                    and raw_final_skin_name[index - 1].islower()
                ):
                    final_skin_parts.append(" ")

                final_skin_parts.append(char)

            final_skin_name = "".join(
                char.lower() if char.isalnum() else " "
                for char in "".join(final_skin_parts)
            )

            final_skin_tokens = set(
                final_skin_name.split()
            )

            if (
                "nondefault" in filename_lower
                or "non-default" in filename_lower
                or "nd" in final_skin_tokens
                or "_nd" in filename_lower
                or "-nd" in filename_lower
            ):

                skin_type = "Non-Default"


            elif any(word in filename_lower for word in [
                "defaultskin",
                "default",
                "defskin",
                "def_skin",
                "def_",
                "_def",
                "-def",
                "replacement",
                "_dr_",
                "-dr-",
                "dr_",
                "_dr"
            ]):

                skin_type = "Default Replacement"


            else:

                skin_type = "Unknown"


        # Recalculate using the FINAL CC type
        # Confirmed Custom Sims keep their technical measurements, but do not
        # receive aggregate HQ/poly performance warnings. Structural health
        # status remains independent and is calculated below.
        if cc_type == "Custom Sim":
            hq_warning = "Normal"
            poly_warning = "Normal"

        elif cc_type == "Mixed CAS" and package_casp_types:

            warning_priority = {
                "Normal": 0,
                "HEAVY": 1,
                "VERY HEAVY": 2
            }

            mixed_warnings = [
                get_poly_warning(
                    highest_polys,
                    casp_type
                )
                for casp_type in package_casp_types
            ]

            poly_warning = max(
                mixed_warnings,
                key=lambda warning: warning_priority.get(
                    warning,
                    0
                )
            )

        else:
            poly_warning = get_poly_warning(
                highest_polys,
                cc_type
            )


        if DEBUG:
            print(
                "POLY CHECK:",
                cc_type,
                highest_polys,
                poly_warning
            )

        # -------------------------
        # UNKNOWN CLASS TRACKING
        # -------------------------

        if cc_type == "Unknown":

            if casp == 0 and geom == 0 and vpxy == 0 and img_count > 0:
                unknown_reason = "IMG only"

            elif casp == 0 and geom == 0 and img_count == 0:
                unknown_reason = "No recognised resources"

            else:
                unknown_reason = "Mixed resources"
              # -------------------------
        # UNKNOWN CLASS DEBUG
        # -------------------------

        if DEBUG and cc_type == "Unknown":

            print(
                "UNKNOWN:",
                os.path.basename(filename),
                "|",
                unknown_reason,
                "| CASP:",
                casp,
                "GEOM:",
                geom,
                "VPXY:",
                vpxy,
                "IMG:",
                img_count,
                "SLIDERS:",
                slider_count
            )  


        # -------------------------
        # CONTENT TYPE
        # -------------------------

        total_resources = resources.Count

        # Package MANIFEST (0x73E93EEB) is provenance/metadata, not payload
        # evidence for EA-derived content.  The same generic TGI is reused by
        # ordinary custom packages, so counting it as an EA hit creates false
        # EA Candidate classifications.  Keep hashing/checking MANIFEST for
        # structural health, but exclude it from content-origin semantics.
        ea_manifest_exact = sum(
            1 for resource_type, _, _ in ea_exact_details
            if resource_type == MANIFEST_TYPE
        )
        ea_manifest_modified = sum(
            1 for resource_type, _, _ in ea_modified_details
            if resource_type == MANIFEST_TYPE
        )
        ea_manifest_unverifiable = sum(
            1 for resource_type, _, _ in ea_unverifiable_details
            if resource_type == MANIFEST_TYPE
        )
        ea_manifest_tgi = (
            ea_manifest_exact
            + ea_manifest_modified
            + ea_manifest_unverifiable
        )
        non_ea_manifest = ea_non_ea_type_counts.get(MANIFEST_TYPE, 0)

        effective_total_resources = (
            total_resources
            - ea_manifest_tgi
            - non_ea_manifest
        )
        effective_ea_matches = ea_matches - ea_manifest_exact
        effective_ea_tgi_matches = ea_tgi_matches - ea_manifest_tgi
        effective_ea_modified_matches = (
            ea_modified_matches - ea_manifest_modified
        )
        effective_ea_unverifiable_matches = (
            ea_unverifiable_matches - ea_manifest_unverifiable
        )
        effective_non_ea_type_counts = {
            resource_type: count
            for resource_type, count in ea_non_ea_type_counts.items()
            if resource_type != MANIFEST_TYPE
        }

        if (
            effective_total_resources > 0
            and effective_ea_matches == effective_total_resources
        ):

            content_type = "EA Match"


        elif (
            effective_total_resources > 0
            and effective_ea_modified_matches > 0
            and (
                effective_ea_tgi_matches
                + sum(
                    count
                    for resource_type, count in effective_non_ea_type_counts.items()
                    if (
                        resource_type == KEY_TYPE
                        or resource_type in THUMBNAIL_PREVIEW_PRIORITY
                    )
                )
            ) == effective_total_resources
            and all(
                (
                    resource_type == KEY_TYPE
                    or resource_type in THUMBNAIL_PREVIEW_PRIORITY
                )
                for resource_type in effective_non_ea_type_counts
            )
        ):

            # NMAP / Name Map and dedicated preview/icon/thumbnail resources
            # are auxiliary package metadata. They must not hide a proven EA
            # payload override when at least one substantive EA TGI is actually
            # modified. This fixes real default replacements whose creator-added
            # preview images use custom TGIs, while ordinary custom CAS packages
            # with no EA TGI overlap remain Custom Content.
            content_type = "EA Override"


        elif effective_ea_tgi_matches > 0:

            # At least one substantive TGI belongs to EA, but the package is
            # mixed or an exact-vs-modified decision could not be completed
            # safely.  MANIFEST-only overlap is deliberately ignored.
            content_type = "EA Candidate"


        else:

            content_type = "Custom Content"


        ignored_manifest_count = ea_manifest_tgi + non_ea_manifest
        if ignored_manifest_count > 0:
            print(
                "EA AUXILIARY MANIFEST IGNORED:",
                os.path.basename(filename),
                "| MANIFEST=", ignored_manifest_count
            )

        if (
            content_type == "EA Override"
            and effective_non_ea_type_counts.get(KEY_TYPE, 0) > 0
        ):
            print(
                "EA AUXILIARY NMAP EXEMPTION:",
                os.path.basename(filename),
                "| NMAP=", effective_non_ea_type_counts.get(KEY_TYPE, 0)
            )

        print(
            "EA CONTENT CHECK:",
            os.path.basename(filename),
            "| Total=", total_resources,
            "| EffectiveTotal=", effective_total_resources,
            "| EA-TGI=", effective_ea_tgi_matches,
            "| Exact=", effective_ea_matches,
            "| Modified=", effective_ea_modified_matches,
            "| Unverifiable=", effective_ea_unverifiable_matches,
            "| Result=", content_type
        )

        # Temporary V4.2 diagnostics.  Print only for packages where the EA
        # result is mixed/ambiguous so we can identify concrete resource types
        # before changing classification semantics.
        if content_type == "EA Candidate":
            for detail_status, details in (
                ("EXACT", ea_exact_details),
                ("MODIFIED", ea_modified_details),
                ("UNVERIFIABLE", ea_unverifiable_details),
            ):
                for dt, dg, di in details:
                    if dt == MANIFEST_TYPE:
                        continue
                    print(
                        "EA RESOURCE DETAIL:",
                        detail_status,
                        "| Type=0x%08X" % (dt & 0xFFFFFFFF),
                        "| Group=0x%08X" % (dg & 0xFFFFFFFF),
                        "| Instance=0x%016X" % (di & 0xFFFFFFFFFFFFFFFF)
                    )

            if effective_non_ea_type_counts:
                print(
                    "EA NON-EA TYPE SUMMARY:",
                    ", ".join(
                        "0x%08X x%d" % (rtype_id, count)
                        for rtype_id, count in sorted(effective_non_ea_type_counts.items())
                    )
                )


        if DEBUG and unknown_resources:

            print("")
            print("UNKNOWN RESOURCE TYPES:")

            for r in sorted(set(unknown_resources)):
                print(r)

            print("")

        # -------------------------
        # RETURN DATA
        # -------------------------

        if DEBUG:

            print("IMG SIZES:")
            print(img_sizes)

        # -------------------------
        # BUILD BUY DEBUG
        # -------------------------

        if DEBUG_BUILD:

            print("BUILD BUY DEBUG:")
            print("OBJD:", objd)
            print("OBJK:", objk)
            print("MODL:", modl)
            print("MLOD:", mlod)
            print("CASP:", casp)
            print("")
            print("UNKNOWN RESOURCE TYPES:")



        # -------------------------
        # HEALTH STATUS
        # -------------------------

        if (
            broken_geom_ids
            or broken_casp_ids
            or broken_img_ids
            or broken_vpxy_ids
            or broken_bone_ids
            or broken_bond_ids
            or broken_morph_ids
            or broken_build_buy_ids
            or broken_preset_ids
            or broken_tone_ids
            or broken_imag_ids
            or broken_txtc_ids
            or broken_ptrn_ids
            or broken_simo_ids
            or broken_txtf_ids
            or broken_famd_ids
            or broken_xml_ids
            or broken_stbl_ids
            or broken_nmap_ids
            or broken_rig_ids
            or broken_jazz_ids
            or broken_s3sa_ids
            or broken_lite_ids
            or broken_clip_ids
            or broken_ngmp_ids
            or broken_layo_ids
            or broken_hinf_ids
            or broken_2ary_ids
            or broken_complate_ids
            or broken_63a33ea7_ids
            or broken_fireplace_ids
            or broken_thum_ids
            or broken_626f60ce_ids
            or broken_icon_ids
            or broken_twni_ids
            or broken_manifest_ids
            or broken_8f78f74b_ids
            or broken_0580a2b5_ids
            or broken_0580a2b4_ids
            or broken_0580a2b6_ids
            or broken_0580a2cd_ids
            or broken_0580a2cf_ids
            or broken_package_storage_ids
        ):
            health_status = "BROKEN"
        else:
            health_status = "OK"

        # Health Issues is intentionally concise for spreadsheet review.
        # Full raw parser exceptions are already printed by the health checks
        # and remain available in Health_Scan_Log.txt. Exact IDs also remain in
        # the far-right Broken <TYPE> IDs columns.
        health_issue_lines = _build_condensed_health_issues([
            ("GEOM", broken_geom_errors),
            ("CASP", broken_casp_errors),
            ("IMG", broken_img_errors),
            ("VPXY", broken_vpxy_errors),
            ("BONE", broken_bone_errors),
            ("BOND", broken_bond_errors),
            ("MORPH", broken_morph_errors),
            ("BUILD/BUY", broken_build_buy_errors),
            ("PRESET", broken_preset_errors),
            ("TONE", broken_tone_errors),
            ("IMAG", broken_imag_errors),
            ("TXTC", broken_txtc_errors),
            ("PTRN", broken_ptrn_errors),
            ("SIMO", broken_simo_errors),
            ("TXTF", broken_txtf_errors),
            ("FAMD", broken_famd_errors),
            ("XML", broken_xml_errors),
            ("STBL", broken_stbl_errors),
            ("NMAP", broken_nmap_errors),
            ("RIG", broken_rig_errors),
            ("JAZZ", broken_jazz_errors),
            ("S3SA", broken_s3sa_errors),
            ("LITE", broken_lite_errors),
            ("CLIP", broken_clip_errors),
            ("NGMP", broken_ngmp_errors),
            ("LAYO", broken_layo_errors),
            ("HINF", broken_hinf_errors),
            ("2ARY", broken_2ary_errors),
            ("COMPLATE", broken_complate_errors),
            ("63A33EA7", broken_63a33ea7_errors),
            ("FIREPLACE", broken_fireplace_errors),
            ("THUM", broken_thum_errors),
            ("626F60CE", broken_626f60ce_errors),
            ("ICON", broken_icon_errors),
            ("TWNI", broken_twni_errors),
            ("MANIFEST", broken_manifest_errors),
            ("8F78F74B", broken_8f78f74b_errors),
            ("THUM 0580A2B5", broken_0580a2b5_errors),
            ("THUM 0580A2B4", broken_0580a2b4_errors),
            ("THUM 0580A2B6", broken_0580a2b6_errors),
            ("SNAP 0580A2CD", broken_0580a2cd_errors),
            ("SNAP 0580A2CF", broken_0580a2cf_errors),
            ("PACKAGE STORAGE", broken_package_storage_errors),
        ])

        # -------------------------
        # RETURN DATA
        # -------------------------
        return [
            biggest_size,
            img_count,
            "; ".join(
                f"{k} ({v})"
                for k, v in resolutions.items()
            ),
            fourk,
            hq_warning,
            casp,
            geom,
            slider_count,
            total_vertices,
            highest_polys,
            geom_total,
            geom_largest,
            highest_geom_id,
            poly_warning,
            cc_type,
            content_type,
            ea_matches,
            ea_casp_matches,
            ea_geom_matches,
            ea_img_matches,
            vpxy,
            bones,
            presets,
            preset_type,
            possible_skin,
            skin_type,
            "\n".join(casp_instances),
            "\n".join(geom_instances),
            "\n".join(img_instances),
            "\n".join(vpxy_instances),
            "\n".join(stbl_instances),
            "\n".join(bgeo_instances),
            "\n".join(face_instances),
            "\n".join(fbln_instances),
            "\n".join(key_instances),
            health_status,                         # 35 Health Status
            len(broken_geom_ids),                  # 36 Broken GEOM Count
            "\n".join(broken_geom_ids),          # 37 Broken GEOM IDs
            "\n".join(health_issue_lines),       # 38 Health Issues
            len(broken_casp_ids),                  # 39 Broken CASP Count
            "\n".join(broken_casp_ids),          # 40 Broken CASP IDs
            len(broken_img_ids),                   # 41 Broken IMG Count
            "\n".join(broken_img_ids),           # 42 Broken IMG IDs
            len(broken_vpxy_ids),                  # 43 Broken VPXY Count
            "\n".join(broken_vpxy_ids),          # 44 Broken VPXY IDs
            len(broken_bone_ids),                   # 45 Broken BONE Count
            "\n".join(broken_bone_ids),          # 46 Broken BONE IDs
            len(broken_bond_ids),                   # 47 Broken BOND Count
            "\n".join(broken_bond_ids),          # 48 Broken BOND IDs
            len(broken_morph_ids),                  # 49 Broken Morph Count
            "\n".join(broken_morph_ids),         # 50 Broken Morph IDs
            len(broken_build_buy_ids),              # 51 Broken Build/Buy Count
            "\n".join(broken_build_buy_ids),     # 52 Broken Build/Buy IDs
            len(broken_preset_ids),                   # 53 Broken PRESET Count
            "\n".join(broken_preset_ids),        # 54 Broken PRESET IDs
            len(broken_tone_ids),                     # 55 Broken TONE Count
            "\n".join(broken_tone_ids),          # 56 Broken TONE IDs
            len(broken_imag_ids),                     # 57 Broken IMAG Count
            "\n".join(broken_imag_ids),          # 58 Broken IMAG IDs
            len(broken_txtc_ids),                     # 59 Broken TXTC Count
            "\n".join(broken_txtc_ids),          # 60 Broken TXTC IDs
            len(broken_ptrn_ids),                     # 61 Broken PTRN Count
            "\n".join(broken_ptrn_ids),          # 62 Broken PTRN IDs
            len(broken_simo_ids),                     # 63 Broken SIMO Count
            "\n".join(broken_simo_ids),          # 64 Broken SIMO IDs
            len(broken_txtf_ids),                     # 65 Broken TXTF Count
            "\n".join(broken_txtf_ids),          # 66 Broken TXTF IDs
            len(broken_famd_ids),                     # 67 Broken FAMD Count
            "\n".join(broken_famd_ids),          # 68 Broken FAMD IDs
            len(broken_objd_ids),                     # 69 Broken OBJD Count
            "\n".join(broken_objd_ids),          # 70 Broken OBJD IDs
            len(broken_objk_ids),                     # 71 Broken OBJK Count
            "\n".join(broken_objk_ids),          # 72 Broken OBJK IDs
            len(broken_xml_ids),                      # 73 Broken XML Count
            "\n".join(broken_xml_ids),           # 74 Broken XML IDs
            len(broken_stbl_ids),                     # 75 Broken STBL Count
            "\n".join(broken_stbl_ids),          # 76 Broken STBL IDs
            len(broken_nmap_ids),                     # 77 Broken NMAP Count
            "\n".join(broken_nmap_ids),          # 78 Broken NMAP IDs
            len(broken_rslt_ids),                     # 79 Broken RSLT Count
            "\n".join(broken_rslt_ids),          # 80 Broken RSLT IDs
            len(broken_ftpt_ids),                     # 81 Broken FTPT Count
            "\n".join(broken_ftpt_ids),          # 82 Broken FTPT IDs
            len(broken_rig_ids),                      # 83 Broken RIG Count
            "\n".join(broken_rig_ids),           # 84 Broken RIG IDs
            len(broken_jazz_ids),                     # 85 Broken JAZZ Count
            "\n".join(broken_jazz_ids),          # 86 Broken JAZZ IDs
            len(broken_s3sa_ids),                     # 87 Broken S3SA Count
            "\n".join(broken_s3sa_ids),          # 88 Broken S3SA IDs
            len(broken_lite_ids),                     # 89 Broken LITE Count
            "\n".join(broken_lite_ids),          # 90 Broken LITE IDs
            len(broken_clip_ids),                     # 91 Broken CLIP Count
            "\n".join(broken_clip_ids),          # 92 Broken CLIP IDs
            len(broken_matd_ids),                     # 93 Broken MATD Count
            "\n".join(broken_matd_ids),          # 94 Broken MATD IDs
            len(broken_mtst_ids),                     # 95 Broken MTST Count
            "\n".join(broken_mtst_ids),          # 96 Broken MTST IDs
            len(broken_mdlr_ids),                  # 97 Broken MDLR Count
            "\n".join(broken_mdlr_ids),       # 98 Broken MDLR IDs
            len(broken_cwal_ids),                  # 99 Broken CWAL Count
            "\n".join(broken_cwal_ids),       # 100 Broken CWAL IDs
            len(broken_cfen_ids),                  # 101 Broken CFEN Count
            "\n".join(broken_cfen_ids),       # 102 Broken CFEN IDs
            len(broken_ctpt_ids),                  # 103 Broken CTPT Count
            "\n".join(broken_ctpt_ids),       # 104 Broken CTPT IDs
            len(broken_ngmp_ids),                  # 105 Broken NGMP Count
            "\n".join(broken_ngmp_ids),       # 106 Broken NGMP IDs
            len(broken_crmt_ids),                  # 107 Broken CRMT Count
            "\n".join(broken_crmt_ids),       # 108 Broken CRMT IDs
            len(broken_spt2_ids),                  # 109 Broken SPT2 Count
            "\n".join(broken_spt2_ids),       # 110 Broken SPT2 IDs
            len(broken_spt_ids),                  # 111 Broken _SPT Count
            "\n".join(broken_spt_ids),       # 112 Broken _SPT IDs
            len(broken_cral_ids),                  # 113 Broken CRAL Count
            "\n".join(broken_cral_ids),       # 114 Broken CRAL IDs
            len(broken_layo_ids),                  # 115 Broken LAYO Count
            "\n".join(broken_layo_ids),       # 116 Broken LAYO IDs
            len(broken_hinf_ids),                  # 117 Broken HINF Count
            "\n".join(broken_hinf_ids),       # 118 Broken HINF IDs
            len(broken_2ary_ids),                  # 119 Broken 2ARY Count
            "\n".join(broken_2ary_ids),       # 120 Broken 2ARY IDs
            len(broken_complate_ids),                  # 121 Broken Complate Count
            "\n".join(broken_complate_ids),       # 122 Broken Complate IDs
            len(broken_63a33ea7_ids),                  # 123 Broken 63A33EA7 Count
            "\n".join(broken_63a33ea7_ids),       # 124 Broken 63A33EA7 IDs
            len(broken_package_storage_ids),              # 125 Broken Package Storage Count
            "\n".join(broken_package_storage_ids),       # 126 Broken Package Storage IDs
            len(broken_fireplace_ids),                      # 127 Broken Fireplace Count
            "\n".join(broken_fireplace_ids),             # 128 Broken Fireplace IDs
            len(broken_thum_ids),                           # 129 Broken THUM Count
            "\n".join(broken_thum_ids),                  # 130 Broken THUM IDs
            len(broken_626f60ce_ids),                       # 131 Broken 626F60CE Count
            "\n".join(broken_626f60ce_ids),              # 132 Broken 626F60CE IDs
            len(broken_icon_ids),                           # 133 Broken ICON Count
            "\n".join(broken_icon_ids),                  # 134 Broken ICON IDs
            len(broken_0580a2b5_ids),                       # 135 Broken THUM 0580A2B5 Count
            "\n".join(broken_0580a2b5_ids),              # 136 Broken THUM 0580A2B5 IDs
            len(broken_0580a2b4_ids),                       # 137 Broken THUM 0580A2B4 Count
            "\n".join(broken_0580a2b4_ids),              # 138 Broken THUM 0580A2B4 IDs
            len(broken_0580a2b6_ids),                       # 139 Broken THUM 0580A2B6 Count
            "\n".join(broken_0580a2b6_ids),              # 140 Broken THUM 0580A2B6 IDs
            len(broken_0580a2cd_ids),                       # 141 Broken SNAP 0580A2CD Count
            "\n".join(broken_0580a2cd_ids),              # 142 Broken SNAP 0580A2CD IDs
            len(broken_0580a2cf_ids),                       # 143 Broken SNAP 0580A2CF Count
            "\n".join(broken_0580a2cf_ids),              # 144 Broken SNAP 0580A2CF IDs
            len(broken_twni_ids),                         # 145 Broken TWNI Count
            "\n".join(broken_twni_ids),                  # 146 Broken TWNI IDs
            len(broken_manifest_ids),                     # 147 Broken MANIFEST Count
            "\n".join(broken_manifest_ids),              # 148 Broken MANIFEST IDs
            len(broken_8f78f74b_ids),                     # 149 Broken 8F78F74B Count
            "\n".join(broken_8f78f74b_ids),              # 150 Broken 8F78F74B IDs
            scripted_object_subtype,                       # 151 Scripted Object Subtype
        ], casp_records
    finally:


        try:
            # release all wrapped resource streams
            for r in resource_objects:

                try:
                    if r is not None and hasattr(r, "Stream"):

                        stream = r.Stream

                        if stream is not None:
                            stream.Dispose()
                            stream.Close()

                except:
                    pass

            resource_objects.clear()

        except Exception as e:
            print("RESOURCE CLEANUP ERROR:", e)


        # release resource list
        resources = None


        # explicitly close package file handle
        try:
            if package is not None:
                Package.ClosePackage(0, package)

        except Exception as e:
            print("PACKAGE CLOSE ERROR:", e)

        finally:
            package = None


INDIVIDUAL_SCAN_FOLDER = r"D:\Documents\CCMagicCheck\Sims3_CC_for_Review"
THUMBNAIL_BACKFILL_FOLDER = r"D:\Documents\CCMagicCheck\Thumbnail_Backfill"
BATCH_SCAN_FOLDER = r"D:\Documents\CCMagicCheck\Big_Review"

_gui_source_override = os.environ.get("S3CC_SOURCE_FOLDER", "").strip()
_gui_mode_override = os.environ.get("S3CC_SCAN_MODE", "").strip().upper()

if _gui_source_override:
    folder = os.path.abspath(_gui_source_override)
    ACTIVE_PACKAGE_BATCH_LIMIT = None
    if _gui_mode_override in ("INDIVIDUAL", "THUMBNAIL_BACKFILL", "BIG_REVIEW_BATCH"):
        SCAN_SOURCE_MODE = _gui_mode_override
    else:
        SCAN_SOURCE_MODE = "INDIVIDUAL"
    print("GUI SOURCE OVERRIDE:", folder)
    print("GUI MODE OVERRIDE:", SCAN_SOURCE_MODE)
    print("GUI MOVE FILES:", MOVE_FILES)
    print("GUI ROUTING PREVIEW:", ROUTING_PREVIEW)
    if ROUTING_ROOT:
        print("GUI ROUTING ROOT:", ROUTING_ROOT)
    if SKIP_MDB_UPDATE:
        print("GUI MASTER DB WRITE: SKIPPED (dry run)")
else:
    os.makedirs(THUMBNAIL_BACKFILL_FOLDER, exist_ok=True)

    # Source priority stays predictable:
    #   1) Sims3_CC_for_Review for normal individual scans
    #   2) Thumbnail_Backfill for fast image-only refreshes of packages already in MDB
    #   3) Big_Review for the normal unlimited bulk scanner
    _individual_scan_items = []

    for _root, _dirs, _files in os.walk(INDIVIDUAL_SCAN_FOLDER):
        for _file in _files:
            if _file.lower().endswith((".package", ".sim", ".sims3pack", ".zip", ".rar")):
                _individual_scan_items.append(os.path.join(_root, _file))

    _thumbnail_backfill_items = []
    for _root, _dirs, _files in os.walk(THUMBNAIL_BACKFILL_FOLDER):
        for _file in _files:
            if _file.lower().endswith(".package"):
                _thumbnail_backfill_items.append(os.path.join(_root, _file))

    if _individual_scan_items:
        folder = INDIVIDUAL_SCAN_FOLDER
        ACTIVE_PACKAGE_BATCH_LIMIT = None
        SCAN_SOURCE_MODE = "INDIVIDUAL"
    elif _thumbnail_backfill_items:
        folder = THUMBNAIL_BACKFILL_FOLDER
        ACTIVE_PACKAGE_BATCH_LIMIT = None
        SCAN_SOURCE_MODE = "THUMBNAIL_BACKFILL"
    else:
        folder = BATCH_SCAN_FOLDER
        ACTIVE_PACKAGE_BATCH_LIMIT = PACKAGE_BATCH_LIMIT
        SCAN_SOURCE_MODE = "BIG_REVIEW_BATCH"

# -------------------------
# ORIGINAL SOURCE PATH MANIFEST SUPPORT
# -------------------------
# Random validation batches contain Validation_Batch_Manifest.txt files with
# paired "Original:" and "Copy:" lines.  Load them before scanning so the
# temporary copied package can carry its real CC Magic source path into the DB.
# Exact copy-path matching is preferred.  A same-folder + basename fallback is
# used only when a batch folder was renamed after the manifest was created.

def _source_path_key(path):
    try:
        return os.path.normcase(os.path.abspath(path))
    except Exception:
        return os.path.normcase(str(path))


def load_original_source_manifest_maps(scan_root):
    exact_map = {}
    folder_name_map = {}
    manifests_found = 0
    mappings_loaded = 0
    conflicts = 0

    def add_mapping(mapping, key, original):
        nonlocal conflicts
        if not key or not original:
            return False
        existing = mapping.get(key)
        if existing is None and key in mapping:
            return False
        if existing and existing != original:
            # Conservative provenance rule: conflicting evidence means do not
            # store a source path for this key.
            mapping[key] = None
            conflicts += 1
            return False
        mapping[key] = original
        return True

    for manifest_root, dirs, files in os.walk(scan_root):
        for manifest_name in files:
            if manifest_name.lower() != "validation_batch_manifest.txt":
                continue

            manifest_path = os.path.join(manifest_root, manifest_name)
            manifests_found += 1
            pending_original = ""

            try:
                with open(manifest_path, "r", encoding="utf-8-sig", errors="replace") as handle:
                    for raw_line in handle:
                        line = raw_line.rstrip("\r\n")

                        if line.startswith("Original:"):
                            pending_original = line.split(":", 1)[1].strip()
                            continue

                        if line.startswith("Copy:") and pending_original:
                            copied_path = line.split(":", 1)[1].strip()
                            if copied_path:
                                exact_key = _source_path_key(copied_path)
                                fallback_key = (
                                    _source_path_key(manifest_root),
                                    os.path.normcase(os.path.basename(copied_path))
                                )
                                exact_added = add_mapping(exact_map, exact_key, pending_original)
                                fallback_added = add_mapping(folder_name_map, fallback_key, pending_original)
                                if exact_added or fallback_added:
                                    mappings_loaded += 1
                            pending_original = ""

            except Exception as exc:
                print(
                    "SOURCE PATH MANIFEST WARNING:",
                    manifest_path,
                    "|",
                    type(exc).__name__ + ":",
                    exc
                )

    return exact_map, folder_name_map, manifests_found, mappings_loaded, conflicts


def original_source_path_for(scanned_path):
    exact_key = _source_path_key(scanned_path)
    exact_value = ORIGINAL_SOURCE_EXACT.get(exact_key)
    if exact_value:
        return exact_value

    fallback_key = (
        _source_path_key(os.path.dirname(scanned_path)),
        os.path.normcase(os.path.basename(scanned_path))
    )
    fallback_value = ORIGINAL_SOURCE_FOLDER_NAME.get(fallback_key)
    if fallback_value:
        return fallback_value

    return ""


def append_unique_source_path(existing_value, new_path):
    existing_lines = [
        line.strip()
        for line in str(existing_value or "").split("\n")
        if line.strip()
    ]
    if new_path and new_path not in existing_lines:
        existing_lines.append(new_path)
    return "\n".join(existing_lines)


def load_existing_scanned_package_hashes():
    """Load exact SHA-256 package identities already present in the Master DB.

    Duplicate routing is intentionally based only on the stored full-file hash.
    CASP overlap, filenames, normalized names, and other heuristics are not used.
    """
    hashes = set()

    if not os.path.exists(MASTER_DATABASE):
        return hashes

    wb = None

    try:
        wb = openpyxl.load_workbook(
            MASTER_DATABASE,
            read_only=True,
            data_only=True
        )

        for sheet_name in (
            "Packages",
            "CC Master Database"
        ):
            if sheet_name not in wb.sheetnames:
                continue

            ws_hash = wb[sheet_name]
            header_map = {}

            for cell in ws_hash[1]:
                if cell.value:
                    header_map[str(cell.value)] = cell.column

            hash_col = header_map.get("Package Hash")

            if not hash_col:
                continue

            for row in ws_hash.iter_rows(
                min_row=2,
                min_col=hash_col,
                max_col=hash_col,
                values_only=True
            ):
                package_hash = str(row[0] or "").strip().lower()

                if len(package_hash) == 64:
                    hashes.add(package_hash)

    except Exception as exc:
        print(
            "DUPLICATE HASH INDEX WARNING:",
            type(exc).__name__ + ":",
            exc
        )

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass

    return hashes


# -------------------------
# FOLDER SAFETY CHECK
# -------------------------

folder_lower = folder.lower()

danger_words = [
    "sims3_sorted_cc",
    "health_scan_result",
    "health_scan_report"
]

for word in danger_words:

    if word in folder_lower:

        print("")
        print("==========================================")
        print("WARNING:")
        print("This folder appears to contain processed CC.")
        print("")
        print("Folder:")
        print(folder)
        print("")
        
        if os.environ.get("S3CC_GUI_MODE", "0") == "1":
            print("GUI TEST SAFETY: refusing processed/output-looking source folder without an interactive override.")
            cleanup_lock_file()
            sys.exit(12)

        answer = input("Continue anyway? (Y/N): ")

        if answer.lower() != "y":
            print("")
            print("Scan cancelled.")
            sys.exit()


print(
    "SCAN SOURCE:",
    SCAN_SOURCE_MODE,
    "|",
    folder
)

ORIGINAL_SOURCE_EXACT, ORIGINAL_SOURCE_FOLDER_NAME, _source_manifest_count, _source_mapping_count, _source_conflict_count = load_original_source_manifest_maps(folder)
print(
    f"SOURCE PATH MANIFEST: manifests={_source_manifest_count:,} | "
    f"mappings={_source_mapping_count:,} | conflicts={_source_conflict_count:,}"
)

if SCAN_SOURCE_MODE == "THUMBNAIL_BACKFILL":
    run_thumbnail_backfill(folder)
    print("Console log saved:", LOG_OUTPUT)
    cleanup_lock_file()
    sys.exit(0)

# Exact duplicate routing is deliberately limited to bulk Big_Review scans.
# Duplicate detection is scoped to the CURRENT source batch only.  Do not seed
# it from Package Hash values already stored in the MDB: the normal workflow is
# Run Scan -> Preview Routing -> Run Routing, so a freshly scanned package is
# already present in the MDB and would otherwise be misclassified as a duplicate
# of itself.  The first byte-for-byte copy in this source batch keeps its normal
# route; later copies in the same batch are routed as duplicates.
if SCAN_SOURCE_MODE == "BIG_REVIEW_BATCH":
    seen_exact_package_hashes = set()
    print("DUPLICATE HASH ROUTING: CURRENT SOURCE BATCH ONLY")
else:
    seen_exact_package_hashes = set()
    print("DUPLICATE HASH ROUTING: OFF FOR INDIVIDUAL SCAN")

rows = []
tech_rows = []
all_casp_records = []
files_to_move = []
duplicates = []
exact_duplicate_paths = set()
# GUI dry-run preview uses its own route records so duplicate detection never
# replaces/suppresses the normal CC category shown to the user.  Like real
# routing, duplicate status is scoped to the current source batch; MDB history
# must not make a package duplicate itself immediately after Run Scan.
routing_preview_records = []
preview_seen_package_hashes = set()
if ROUTING_PREVIEW:
    print("ROUTING PREVIEW DUPLICATE INDEX: current source batch only")
import time
print("")
print("SCAN SESSION:")
print(scan_session_folder)
print("")
scan_start = time.time()


# Discover package paths once and choose a deterministic batch.
# Sorting by normalized full path makes batch membership stable across runs.
all_package_paths = []

for root, dirs, files in os.walk(folder):
    for file in files:
        if file.lower().endswith(".package"):
            all_package_paths.append(
                os.path.join(root, file)
            )

all_package_paths.sort(
    key=lambda p: os.path.normcase(os.path.abspath(p))
)

package_available_total = len(all_package_paths)

if ACTIVE_PACKAGE_BATCH_LIMIT is None:
    selected_package_paths = all_package_paths
else:
    selected_package_paths = all_package_paths[:ACTIVE_PACKAGE_BATCH_LIMIT]

selected_package_path_set = set(selected_package_paths)
package_total = len(selected_package_paths)


print("")
print("TOTAL PACKAGES AVAILABLE:", package_available_total)
if ACTIVE_PACKAGE_BATCH_LIMIT is None:
    if SCAN_SOURCE_MODE == "BIG_REVIEW_BATCH":
        print(
            "BATCH MODE:",
            "UNLIMITED - all available packages selected",
            "| Packages selected=",
            package_total
        )
    else:
        print(
            "BATCH MODE:",
            "OFF - individual scan folder has priority",
            "| Packages selected=",
            package_total
        )
else:
    print(
        "BATCH MODE:",
        package_total,
        "package(s) selected",
        "| Limit=",
        ACTIVE_PACKAGE_BATCH_LIMIT
    )
print("TOTAL PACKAGES TO SCAN:", package_total)
print("")


scan_number = 0
sims2_unsupported_count = 0
sims4_unsupported_count = 0
thumbnail_paths_by_hash = {}
thumbnail_extracted_count = 0
thumbnail_reused_count = 0
thumbnail_missing_count = 0
thumbnail_error_count = 0
sims3pack_passthrough_count = 0
archive_passthrough_count = 0
image_passthrough_count = 0


for root, dirs, files in os.walk(folder):

    for file in files:

        file_lower = file.lower()

        # -------------------------
        # NON-PACKAGE PASSTHROUGH FILES
        # -------------------------
        # Never open or extract these. Move them intact after the package scan.
        if file_lower.endswith(".sims3pack"):
            path = os.path.join(root, file)
            files_to_move.append((path, "SIMS3PACK"))
            if ROUTING_PREVIEW:
                routing_preview_records.append((path, "SIMS3PACK", "No"))
            sims3pack_passthrough_count += 1
            continue

        if file_lower.endswith((".zip", ".rar")):
            path = os.path.join(root, file)
            files_to_move.append((path, "ARCHIVES"))
            if ROUTING_PREVIEW:
                routing_preview_records.append((path, "ARCHIVES", "No"))
            archive_passthrough_count += 1
            continue

        if file_lower.endswith((".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp")):
            path = os.path.join(root, file)
            files_to_move.append((path, "IMAGES"))
            if ROUTING_PREVIEW:
                routing_preview_records.append((path, "IMAGES", "No"))
            image_passthrough_count += 1
            continue

        # -------------------------
        # EXPORTED SIM FILES
        # -------------------------

        if file_lower.endswith(".sim"):

            path = os.path.join(root, file)

            if DEBUG:
                print(
                    f"Found Sim file: {file}"
                )

            rows.append([
                file,              # 0 Package
                "",                # 1 CASP IDs
                "",                # 2 GEOM IDs
                "",                # 3 IMG IDs
                "",                # 4 VPXY IDs
                "",                # 5 Duplicate Files
                "Custom Sim",      # 6 CC Type
                "Custom Content",  # 7 Content Type
                "Safe",            # 8 Risk
                "0/10",            # 9 Risk Score
                "Normal",          # 10 HQ Warning
                "Normal",          # 11 Poly Warning
                0,                 # 12 Highest Polys
                "",                # 13 Biggest Texture
                "",                # 14 Skin Type
                "",                # 15 STBL IDs
                "",                # 16 BGEO IDs
                "",                # 17 FACE IDs
                "",                # 18 FBLN IDs
                "",                # 19 KEY IDs
                0,                 # 20 BONE Resource Count
                "",                # 21 Package Hash
                "OK",              # 22 Health Status
                0,                 # 23 Broken GEOM Count
                "",                # 24 Broken GEOM IDs
                "",                # 25 Health Issues
                0,                 # 26 Broken CASP Count
                "",                # 27 Broken CASP IDs
                0,                 # 28 Broken IMG Count
                "",                # 29 Broken IMG IDs
                0,                 # 30 Broken VPXY Count
                "",                # 31 Broken VPXY IDs
                0,                 # 32 Broken BONE Count
                "",                # 33 Broken BONE IDs
                0,                 # 34 Broken BOND Count
                "",                # 35 Broken BOND IDs
                0,                 # 36 Broken Morph Count
                "",                # 37 Broken Morph IDs
                0,                 # 38 Broken Build/Buy Count
                "",                # 39 Broken Build/Buy IDs
                0,                 # 40 Broken PRESET Count
                "",                # 41 Broken PRESET IDs
                0,                 # 42 Broken TONE Count
                "",                # 43 Broken TONE IDs
                0,                 # 44 Broken IMAG Count
                "",                # 45 Broken IMAG IDs
                0,                 # 46 Broken TXTC Count
                "",                # 47 Broken TXTC IDs
                0,                 # 48 Broken PTRN Count
                "",                # 49 Broken PTRN IDs
                0,                 # 50 Broken SIMO Count
                "",                # 51 Broken SIMO IDs
                0,                 # 52 Broken TXTF Count
                "",                # 53 Broken TXTF IDs
                0,                 # 54 Broken FAMD Count
                "",                # 55 Broken FAMD IDs
                0,                 # 56 Broken OBJD Count
                "",                # 57 Broken OBJD IDs
                0,                 # 58 Broken OBJK Count
                "",                # 59 Broken OBJK IDs
                0,                 # 60 Broken XML Count
                "",                # 61 Broken XML IDs
                0,                 # 62 Broken STBL Count
                "",                # 63 Broken STBL IDs
                0,                 # 64 Broken NMAP Count
                "",                # 65 Broken NMAP IDs
                0,                 # 66 Broken RSLT Count
                "",                # 67 Broken RSLT IDs
                0,                 # 68 Broken FTPT Count
                "",                # 69 Broken FTPT IDs
                0,                 # 70 Broken RIG Count
                "",                # 71 Broken RIG IDs
                0,                 # 72 Broken JAZZ Count
                "",                # 73 Broken JAZZ IDs
                0,                 # 74 Broken S3SA Count
                "",                # 75 Broken S3SA IDs
                0,                 # 76 Broken LITE Count
                "",                # 77 Broken LITE IDs
                0,                 # 78 Broken CLIP Count
                "",                # 79 Broken CLIP IDs
                0,                 # 80 Broken MATD Count
                "",                # 81 Broken MATD IDs
                0,                 # 82 Broken MTST Count
                "",                # 83 Broken MTST IDs
                0,                 # 84 Broken MDLR Count
                "",                # 85 Broken MDLR IDs
                0,                 # 86 Broken CWAL Count
                "",                # 87 Broken CWAL IDs
                0,                 # 88 Broken CFEN Count
                "",                # 89 Broken CFEN IDs
                0,                 # 90 Broken CTPT Count
                "",                # 91 Broken CTPT IDs
                0,                 # 92 Broken NGMP Count
                "",                # 93 Broken NGMP IDs
                0,                 # 94 Broken CRMT Count
                "",                # 95 Broken CRMT IDs
                0,                 # 96 Broken SPT2 Count
                "",                # 97 Broken SPT2 IDs
                0,                 # 98 Broken _SPT Count
                "",                # 99 Broken _SPT IDs
                0,                 # 100 Broken CRAL Count
                "",                # 101 Broken CRAL IDs
                0,                 # 102 Broken LAYO Count
                "",                # 103 Broken LAYO IDs
                0,                 # 104 Broken HINF Count
                "",                # 105 Broken HINF IDs
                0,                 # 106 Broken 2ARY Count
                "",                # 107 Broken 2ARY IDs
                0,                 # 108 Broken Complate Count
                "",                # 109 Broken Complate IDs
                0,                 # 110 Broken 63A33EA7 Count
                "",                # 111 Broken 63A33EA7 IDs
                0,                 # 112 Broken Package Storage Count
                "",                # 113 Broken Package Storage IDs
                0,                 # 114 Broken Fireplace Count
                "",                # 115 Broken Fireplace IDs
                0,                 # 116 Broken THUM Count
                "",                # 117 Broken THUM IDs
                0,                 # 118 Broken 626F60CE Count
                "",                # 119 Broken 626F60CE IDs
                0,                 # 120 Broken ICON Count
                "",                # 121 Broken ICON IDs
                0,                 # 122 Broken THUM 0580A2B5 Count
                "",                # 123 Broken THUM 0580A2B5 IDs
                0,                 # 124 Broken THUM 0580A2B4 Count
                "",                # 125 Broken THUM 0580A2B4 IDs
                0,                 # 126 Broken THUM 0580A2B6 Count
                "",                # 127 Broken THUM 0580A2B6 IDs
                0,                 # 128 Broken SNAP 0580A2CD Count
                "",                # 129 Broken SNAP 0580A2CD IDs
                0,                 # 130 Broken SNAP 0580A2CF Count
                "",                # 131 Broken SNAP 0580A2CF IDs
                0,                 # 132 Broken TWNI Count
                "",                # 133 Broken TWNI IDs
                0,                 # 134 Broken MANIFEST Count
                "",                # 135 Broken MANIFEST IDs
                0,                 # 136 Broken 8F78F74B Count
                "",                # 137 Broken 8F78F74B IDs
                ""                 # 138 Original Source Path
            ])

            files_to_move.append(
                (
                    path,
                    "Custom_Sim"
                )
            )
            if ROUTING_PREVIEW:
                routing_preview_records.append((path, "Custom_Sim", "No"))

            continue


        if file.lower().endswith(".package"):

            path = os.path.join(root, file)

            if path not in selected_package_path_set:
                continue

            scan_number += 1

            original_source_path = original_source_path_for(path)

            print(
                f"Scanning {scan_number}/{package_total}: {file}"
            )

            elapsed = int(time.time() - scan_start)

            minutes = elapsed // 60
            seconds = elapsed % 60

            print(
                f"Elapsed time: {minutes}m {seconds}s"
            )

            try:

                if os.path.getsize(path) == 0:
                    raise ValueError("Empty package file (0 bytes)")

                foreign_game, dbpf_major, dbpf_minor = _known_foreign_sims_game(path)
                if foreign_game:
                    raise ValueError(
                        f"{FOREIGN_GAME_MARKER}{foreign_game}; "
                        f"DBPF={dbpf_major}.{dbpf_minor}"
                    )

                data, casp_records = scan_package(path)

                package_hash = calculate_file_hash(path)
                preview_duplicate_flag = "No"
                if ROUTING_PREVIEW and package_hash:
                    package_hash_key_for_preview = package_hash.lower()
                    if package_hash_key_for_preview in preview_seen_package_hashes:
                        preview_duplicate_flag = "Yes"
                    preview_seen_package_hashes.add(package_hash_key_for_preview)

                thumbnail_path, thumbnail_status = extract_best_package_thumbnail(
                    path,
                    package_hash
                )
                if thumbnail_path and package_hash:
                    thumbnail_paths_by_hash[package_hash.lower()] = thumbnail_path

                if thumbnail_status == "extracted":
                    thumbnail_extracted_count += 1
                elif thumbnail_status == "existing":
                    thumbnail_reused_count += 1
                elif thumbnail_status == "none":
                    thumbnail_missing_count += 1
                elif thumbnail_status == "error":
                    thumbnail_error_count += 1

                is_exact_duplicate = False

                if (
                    SCAN_SOURCE_MODE == "BIG_REVIEW_BATCH"
                    and package_hash
                ):
                    package_hash_key = package_hash.lower()

                    if package_hash_key in seen_exact_package_hashes:
                        is_exact_duplicate = True
                        exact_duplicate_paths.add(path)
                        print(
                            "EXACT SHA256 DUPLICATE:",
                            file,
                            "|",
                            package_hash
                        )
                    else:
                        # The first copy seen in this run keeps normal routing.
                        # Later byte-for-byte copies in the same batch are duplicates.
                        seen_exact_package_hashes.add(package_hash_key)
                
                all_casp_records.extend(casp_records)

                if DEBUG:
                    print("DEBUG DATA:")
                    print(data)

                possible_skin = data[24]
                skin_type = data[25]
                hq_warning = data[4]
                fourk = data[3]
                poly_warning = data[13]
                geom_bytes = data[10]
                cc_type = data[14]
                vpxy = data[20]
                health_status = data[35]
                broken_geom_count = data[36]
                broken_geom_ids = data[37]
                health_issues = data[38]
                broken_casp_count = data[39]
                broken_casp_ids = data[40]
                broken_img_count = data[41]
                broken_img_ids = data[42]
                broken_vpxy_count = data[43]
                broken_vpxy_ids = data[44]
                broken_bone_count = data[45]
                broken_bone_ids = data[46]
                broken_bond_count = data[47]
                broken_bond_ids = data[48]
                broken_morph_count = data[49]
                broken_morph_ids = data[50]
                broken_build_buy_count = data[51]
                broken_build_buy_ids = data[52]
                broken_preset_count = data[53]
                broken_preset_ids = data[54]
                broken_tone_count = data[55]
                broken_tone_ids = data[56]
                broken_imag_count = data[57]
                broken_imag_ids = data[58]
                broken_txtc_count = data[59]
                broken_txtc_ids = data[60]
                broken_ptrn_count = data[61]
                broken_ptrn_ids = data[62]
                broken_simo_count = data[63]
                broken_simo_ids = data[64]
                broken_txtf_count = data[65]
                broken_txtf_ids = data[66]
                broken_famd_count = data[67]
                broken_famd_ids = data[68]
                broken_objd_count = data[69]
                broken_objd_ids = data[70]
                broken_objk_count = data[71]
                broken_objk_ids = data[72]
                broken_xml_count = data[73]
                broken_xml_ids = data[74]
                broken_stbl_count = data[75]
                broken_stbl_ids = data[76]
                broken_nmap_count = data[77]
                broken_nmap_ids = data[78]
                broken_rslt_count = data[79]
                broken_rslt_ids = data[80]
                broken_ftpt_count = data[81]
                broken_ftpt_ids = data[82]
                broken_rig_count = data[83]
                broken_rig_ids = data[84]
                broken_jazz_count = data[85]
                broken_jazz_ids = data[86]
                broken_s3sa_count = data[87]
                broken_s3sa_ids = data[88]
                broken_lite_count = data[89]
                broken_lite_ids = data[90]
                broken_clip_count = data[91]
                broken_clip_ids = data[92]
                broken_matd_count = data[93]
                broken_matd_ids = data[94]
                broken_mtst_count = data[95]
                broken_mtst_ids = data[96]
                broken_mdlr_count = data[97]
                broken_mdlr_ids = data[98]
                broken_cwal_count = data[99]
                broken_cwal_ids = data[100]
                broken_cfen_count = data[101]
                broken_cfen_ids = data[102]
                broken_ctpt_count = data[103]
                broken_ctpt_ids = data[104]
                broken_ngmp_count = data[105]
                broken_ngmp_ids = data[106]
                broken_crmt_count = data[107]
                broken_crmt_ids = data[108]
                broken_spt2_count = data[109]
                broken_spt2_ids = data[110]
                broken_spt_count = data[111]
                broken_spt_ids = data[112]
                broken_cral_count = data[113]
                broken_cral_ids = data[114]
                broken_layo_count = data[115]
                broken_layo_ids = data[116]
                broken_hinf_count = data[117]
                broken_hinf_ids = data[118]
                broken_2ary_count = data[119]
                broken_2ary_ids = data[120]
                broken_complate_count = data[121]
                broken_complate_ids = data[122]
                broken_63a33ea7_count = data[123]
                broken_63a33ea7_ids = data[124]
                broken_package_storage_count = data[125]
                broken_package_storage_ids = data[126]
                broken_fireplace_count = data[127]
                broken_fireplace_ids = data[128]
                broken_thum_count = data[129]
                broken_thum_ids = data[130]
                broken_626f60ce_count = data[131]
                broken_626f60ce_ids = data[132]
                broken_icon_count = data[133]
                broken_icon_ids = data[134]
                broken_0580a2b5_count = data[135]
                broken_0580a2b5_ids = data[136]
                broken_0580a2b4_count = data[137]
                broken_0580a2b4_ids = data[138]
                broken_0580a2b6_count = data[139]
                broken_0580a2b6_ids = data[140]
                broken_0580a2cd_count = data[141]
                broken_0580a2cd_ids = data[142]
                broken_0580a2cf_count = data[143]
                broken_0580a2cf_ids = data[144]
                broken_twni_count = data[145]
                broken_twni_ids = data[146]
                broken_manifest_count = data[147]
                broken_manifest_ids = data[148]
                broken_8f78f74b_count = data[149]
                broken_8f78f74b_ids = data[150]


                risk_score, overall_risk = calculate_risk(
                    hq_warning,
                    poly_warning,
                    geom_bytes,
                    fourk,
                    vpxy,
                    cc_type
                )

                risk_reasons = get_risk_reasons(
                    data[4],    # HQ Warning
                    data[13],   # Poly Warning
                    data[10],   # GEOM bytes
                    data[3],    # 4K count
                    data[9],    # Highest poly count
                    data[14],   # CC Type
                    data[20]    # VPXY count
                )


                # -------------------------
                # FILE MOVE CATEGORY

                content_type = data[15]


                if ROUTING_ACTIVE:
                    move_category = None


                    # -------------------------
                    # EXACT SHA256 DUPLICATE ROUTING
                    # -------------------------
                    # Physical routing only: the package is still fully scanned and
                    # its health/risk/classification data is retained in the report/DB.
                    if is_exact_duplicate and not ROUTING_PREVIEW:

                        move_category = "DUPLICATES"

                    # -------------------------
                    # BROKEN RESOURCE HEALTH OVERRIDE
                    # -------------------------

                    elif health_status == "BROKEN":

                        if broken_package_storage_count:
                            broken_type = "PACKAGE_STORAGE"

                        else:
                            broken_health_types = sum(
                                1
                                for count in (
                                    broken_geom_count,
                                    broken_casp_count,
                                    broken_img_count,
                                    broken_vpxy_count,
                                    broken_bone_count,
                                    broken_bond_count,
                                    broken_morph_count,
                                    broken_build_buy_count,
                                    broken_preset_count,
                                    broken_tone_count,
                                    broken_imag_count,
                                    broken_txtc_count,
                                    broken_ptrn_count,
                                    broken_simo_count,
                                    broken_txtf_count,
                                    broken_famd_count,
                                    broken_xml_count,
                                    broken_stbl_count,
                                    broken_nmap_count,
                                    broken_rig_count,
                                    broken_jazz_count,
                                    broken_s3sa_count,
                                    broken_lite_count,
                                    broken_clip_count,
                                    broken_ngmp_count,
                                    broken_layo_count,
                                    broken_hinf_count,
                                    broken_2ary_count,
                                    broken_complate_count,
                                    broken_63a33ea7_count,
                                    broken_fireplace_count,
                                    broken_thum_count,
                                    broken_626f60ce_count,
                                    broken_icon_count,
                                    broken_twni_count,
                                    broken_manifest_count,
                                    broken_8f78f74b_count,
                                    broken_0580a2b5_count,
                                    broken_0580a2b4_count,
                                    broken_0580a2b6_count,
                                    broken_0580a2cd_count,
                                    broken_0580a2cf_count
                                )
                                if count
                            )

                            if broken_health_types > 1:
                                broken_type = "MULTIPLE"
                            elif broken_0580a2cf_count:
                                broken_type = "SNAP"
                            elif broken_0580a2cd_count:
                                broken_type = "SNAP"
                            elif broken_0580a2b6_count:
                                broken_type = "THUM"
                            elif broken_0580a2b4_count:
                                broken_type = "THUM"
                            elif broken_0580a2b5_count:
                                broken_type = "THUM"
                            elif broken_8f78f74b_count:
                                broken_type = "8F78F74B"
                            elif broken_manifest_count:
                                broken_type = "MANIFEST"
                            elif broken_twni_count:
                                broken_type = "TWNI"
                            elif broken_icon_count:
                                broken_type = "ICON"
                            elif broken_626f60ce_count:
                                broken_type = "THUM"
                            elif broken_thum_count:
                                broken_type = "THUM"
                            elif broken_fireplace_count:
                                broken_type = "BUILD_BUY"
                            elif broken_63a33ea7_count:
                                broken_type = "BUILD_BUY"
                            elif broken_complate_count:
                                broken_type = "COMPLATE"
                            elif broken_2ary_count:
                                broken_type = "2ARY"
                            elif broken_hinf_count:
                                broken_type = "HINF"
                            elif broken_layo_count:
                                broken_type = "LAYO"
                            elif broken_ngmp_count:
                                broken_type = "NGMP"
                            elif broken_clip_count:
                                broken_type = "CLIP"
                            elif broken_lite_count:
                                broken_type = "LITE"
                            elif broken_s3sa_count:
                                broken_type = "S3SA"
                            elif broken_jazz_count:
                                broken_type = "JAZZ"
                            elif broken_rig_count:
                                broken_type = "RIG"
                            elif broken_nmap_count:
                                broken_type = "NMAP"
                            elif broken_stbl_count:
                                broken_type = "STBL"
                            elif broken_xml_count:
                                broken_type = "XML"
                            elif broken_famd_count:
                                broken_type = "FAMD"
                            elif broken_txtf_count:
                                broken_type = "TXTF"
                            elif broken_simo_count:
                                broken_type = "SIMO"
                            elif broken_ptrn_count:
                                broken_type = "PTRN"
                            elif broken_txtc_count:
                                broken_type = "TXTC"
                            elif broken_imag_count:
                                broken_type = "IMAG"
                            elif broken_tone_count:
                                broken_type = "SKIN"
                            elif broken_preset_count:
                                broken_type = "PRESET"
                            elif broken_build_buy_count:
                                broken_type = "BUILD_BUY"
                            elif broken_morph_count:
                                broken_type = "MORPH"
                            elif broken_bond_count:
                                broken_type = "BOND"
                            elif broken_bone_count:
                                broken_type = "BONE"
                            elif broken_vpxy_count:
                                broken_type = "VPXY"
                            elif broken_img_count:
                                broken_type = "IMG"
                            elif broken_casp_count:
                                broken_type = "CASP"
                            elif broken_geom_count:
                                broken_type = "GEOM"
                            else:
                                broken_type = "UNKNOWN"


                        move_category = os.path.join(
                            "BROKEN_CC",
                            broken_type
                        )

                    # -------------------------
                    # UNIDENTIFIED CC
                    # -------------------------

                    elif (
                        content_type == "Custom Content"
                        and cc_type == "Unknown"
                        and not possible_skin
                    ):

                        move_category = "UNIDENTIFIED_CC"


                    # -------------------------
                    # EA MATCH
                    # -------------------------

                    elif content_type == "EA Match":

                        move_category = "EA_MATCH"


                    # -------------------------
                    # EA OVERRIDE
                    # -------------------------

                    elif content_type == "EA Override":

                        move_category = "EA_OVERRIDE"


                    # -------------------------
                    # EA CANDIDATE
                    # -------------------------

                    elif (
                        content_type == "EA Candidate"
                        and not possible_skin
                    ):

                        move_category = "EA_CANDIDATE"


                    # -------------------------
                    # POSSIBLE SKIN
                    # -------------------------

                    if DEBUG:
                        print(
                            "SKIN CHECK:",
                            os.path.basename(path),
                            "CC TYPE:",
                            cc_type,
                            "POSSIBLE SKIN:",
                            possible_skin
                        )

                    if (
                        health_status != "BROKEN"
                        and possible_skin
                        and cc_type == "Unknown"
                        and content_type not in ["EA Match", "EA Override"]
                    ):

                        if skin_type not in [
                            "Non-Default",
                            "Default Replacement"
                        ]:
                            skin_type = "Unknown"

                        move_category = os.path.join(
                            "POSSIBLE_SKIN",
                            skin_type
                        )

                    # -------------------------
                    # NORMAL NON-BROKEN CUSTOM CONTENT
                    # -------------------------
                    # Risk Score / textual Risk are review information only.
                    # They must never decide the physical routing destination.
                    # Packages with direct CASP identity are grouped under
                    # CASP_OBSERVED; other identified custom content is grouped
                    # under CUSTOM_CONTENT. Unknown custom content was handled
                    # above by UNIDENTIFIED_CC.

                    elif (
                        move_category is None
                        and content_type == "Custom Content"
                        and cc_type
                        and cc_type != "Unknown"
                    ):

                        casp_observed = bool(str(data[26] or "").strip())
                        route_root = "CASP_OBSERVED" if casp_observed else "CUSTOM_CONTENT"

                        move_category = os.path.join(
                            route_root,
                            "Custom_Sim" if cc_type == "Custom Sim" else cc_type
                        )


                    # -------------------------
                    # MOVE QUEUE
                    # -------------------------

                    if move_category:

                        files_to_move.append(
                            (
                                path,
                                move_category
                            )
                        )

                    if ROUTING_PREVIEW:
                        routing_preview_records.append(
                            (
                                path,
                                move_category or "UNIDENTIFIED_CC",
                                preview_duplicate_flag
                            )
                        )
                


                # Priority label for Technical Details sheet

                if health_status == "BROKEN":

                    priority = "BROKEN"

                elif overall_risk == "Extreme":

                    priority = "URGENT"

                elif overall_risk == "High":

                    priority = "HIGH"

                elif overall_risk == "Moderate":

                    priority = "MODERATE"

                else:

                    priority = "NORMAL"


                # -------------------------
                # Summary Sheet Data
                # -------------------------
                rows.append([
                    file,
                    data[26],      # CASP IDs
                    data[27],      # GEOM IDs
                    data[28],      # IMG IDs
                    data[29],      # VPXY IDs
                    "",
                    data[14],      # CC Type
                    data[15],      # Content Type
                    overall_risk,
                    risk_score,
                    data[4],       # HQ Warning
                    data[13],      # Poly Warning
                    data[9],       # Highest Poly Count
                    data[0],       # Biggest Size
                    data[25],      # Skin Type

                    data[30],      # STBL
                    data[31],      # BGEO
                    data[32],      # FACE
                    data[33],      # FBLN
                    data[34],      # KEY
                    data[21],      # BONE Resource Count
                    package_hash,  # Package Hash
                    health_status,       # 22 Health Status
                    broken_geom_count,   # 23 Broken GEOM Count
                    broken_geom_ids,     # 24 Broken GEOM IDs
                    health_issues,       # 25 Health Issues
                    broken_casp_count,   # 26 Broken CASP Count
                    broken_casp_ids,     # 27 Broken CASP IDs
                    broken_img_count,    # 28 Broken IMG Count
                    broken_img_ids,      # 29 Broken IMG IDs
                    broken_vpxy_count,   # 30 Broken VPXY Count
                    broken_vpxy_ids,     # 31 Broken VPXY IDs
                    broken_bone_count,   # 32 Broken BONE Count
                    broken_bone_ids,     # 33 Broken BONE IDs
                    broken_bond_count,   # 34 Broken BOND Count
                    broken_bond_ids,     # 35 Broken BOND IDs
                    broken_morph_count,      # 36 Broken Morph Count
                    broken_morph_ids,        # 37 Broken Morph IDs
                    broken_build_buy_count,  # 38 Broken Build/Buy Count
                    broken_build_buy_ids,    # 39 Broken Build/Buy IDs
                    broken_preset_count,     # 40 Broken PRESET Count
                    broken_preset_ids,       # 41 Broken PRESET IDs
                    broken_tone_count,       # 42 Broken TONE Count
                    broken_tone_ids,         # 43 Broken TONE IDs
                    broken_imag_count,       # 44 Broken IMAG Count
                    broken_imag_ids,         # 45 Broken IMAG IDs
                    broken_txtc_count,       # 46 Broken TXTC Count
                    broken_txtc_ids,         # 47 Broken TXTC IDs
                    broken_ptrn_count,       # 48 Broken PTRN Count
                    broken_ptrn_ids,         # 49 Broken PTRN IDs
                    broken_simo_count,       # 50 Broken SIMO Count
                    broken_simo_ids,         # 51 Broken SIMO IDs
                    broken_txtf_count,       # 52 Broken TXTF Count
                    broken_txtf_ids,         # 53 Broken TXTF IDs
                    broken_famd_count,       # 54 Broken FAMD Count
                    broken_famd_ids,         # 55 Broken FAMD IDs
                    broken_objd_count,       # 56 Broken OBJD Count
                    broken_objd_ids,         # 57 Broken OBJD IDs
                    broken_objk_count,       # 58 Broken OBJK Count
                    broken_objk_ids,         # 59 Broken OBJK IDs
                    broken_xml_count,        # 60 Broken XML Count
                    broken_xml_ids,          # 61 Broken XML IDs
                    broken_stbl_count,       # 62 Broken STBL Count
                    broken_stbl_ids,         # 63 Broken STBL IDs
                    broken_nmap_count,       # 64 Broken NMAP Count
                    broken_nmap_ids,         # 65 Broken NMAP IDs
                    broken_rslt_count,       # 66 Broken RSLT Count
                    broken_rslt_ids,         # 67 Broken RSLT IDs
                    broken_ftpt_count,       # 68 Broken FTPT Count
                    broken_ftpt_ids,         # 69 Broken FTPT IDs
                    broken_rig_count,        # 70 Broken RIG Count
                    broken_rig_ids,          # 71 Broken RIG IDs
                    broken_jazz_count,       # 72 Broken JAZZ Count
                    broken_jazz_ids,         # 73 Broken JAZZ IDs
                    broken_s3sa_count,       # 74 Broken S3SA Count
                    broken_s3sa_ids,         # 75 Broken S3SA IDs
                    broken_lite_count,       # 76 Broken LITE Count
                    broken_lite_ids,         # 77 Broken LITE IDs
                    broken_clip_count,       # 78 Broken CLIP Count
                    broken_clip_ids,         # 79 Broken CLIP IDs
                    broken_matd_count,       # 80 Broken MATD Count
                    broken_matd_ids,         # 81 Broken MATD IDs
                    broken_mtst_count,       # 82 Broken MTST Count
                    broken_mtst_ids,         # 83 Broken MTST IDs
                    broken_mdlr_count,    # 84 Broken MDLR Count
                    broken_mdlr_ids,      # 85 Broken MDLR IDs
                    broken_cwal_count,    # 86 Broken CWAL Count
                    broken_cwal_ids,      # 87 Broken CWAL IDs
                    broken_cfen_count,    # 88 Broken CFEN Count
                    broken_cfen_ids,      # 89 Broken CFEN IDs
                    broken_ctpt_count,    # 90 Broken CTPT Count
                    broken_ctpt_ids,      # 91 Broken CTPT IDs
                    broken_ngmp_count,    # 92 Broken NGMP Count
                    broken_ngmp_ids,      # 93 Broken NGMP IDs
                    broken_crmt_count,    # 94 Broken CRMT Count
                    broken_crmt_ids,      # 95 Broken CRMT IDs
                    broken_spt2_count,    # 96 Broken SPT2 Count
                    broken_spt2_ids,      # 97 Broken SPT2 IDs
                    broken_spt_count,     # 98 Broken _SPT Count
                    broken_spt_ids,       # 99 Broken _SPT IDs
                    broken_cral_count,    # 100 Broken CRAL Count
                    broken_cral_ids,      # 101 Broken CRAL IDs
                    broken_layo_count,    # 102 Broken LAYO Count
                    broken_layo_ids,      # 103 Broken LAYO IDs
                    broken_hinf_count,    # 104 Broken HINF Count
                    broken_hinf_ids,      # 105 Broken HINF IDs
                    broken_2ary_count,    # 106 Broken 2ARY Count
                    broken_2ary_ids,      # 107 Broken 2ARY IDs
                    broken_complate_count,    # 108 Broken Complate Count
                    broken_complate_ids,      # 109 Broken Complate IDs
                    broken_63a33ea7_count,    # 110 Broken 63A33EA7 Count
                    broken_63a33ea7_ids,      # 111 Broken 63A33EA7 IDs
                    broken_package_storage_count,    # 112 Broken Package Storage Count
                    broken_package_storage_ids,      # 113 Broken Package Storage IDs
                    broken_fireplace_count,          # 114 Broken Fireplace Count
                    broken_fireplace_ids,            # 115 Broken Fireplace IDs
                    broken_thum_count,               # 116 Broken THUM Count
                    broken_thum_ids,                 # 117 Broken THUM IDs
                    broken_626f60ce_count,           # 118 Broken 626F60CE Count
                    broken_626f60ce_ids,             # 119 Broken 626F60CE IDs
                    broken_icon_count,               # 120 Broken ICON Count
                    broken_icon_ids,                 # 121 Broken ICON IDs
                    broken_0580a2b5_count,           # 122 Broken THUM 0580A2B5 Count
                    broken_0580a2b5_ids,             # 123 Broken THUM 0580A2B5 IDs
                    broken_0580a2b4_count,           # 124 Broken THUM 0580A2B4 Count
                    broken_0580a2b4_ids,             # 125 Broken THUM 0580A2B4 IDs
                    broken_0580a2b6_count,           # 126 Broken THUM 0580A2B6 Count
                    broken_0580a2b6_ids,             # 127 Broken THUM 0580A2B6 IDs
                    broken_0580a2cd_count,           # 128 Broken SNAP 0580A2CD Count
                    broken_0580a2cd_ids,             # 129 Broken SNAP 0580A2CD IDs
                    broken_0580a2cf_count,           # 130 Broken SNAP 0580A2CF Count
                    broken_0580a2cf_ids,             # 131 Broken SNAP 0580A2CF IDs
                    broken_twni_count,                # 132 Broken TWNI Count
                    broken_twni_ids,                  # 133 Broken TWNI IDs
                    broken_manifest_count,            # 134 Broken MANIFEST Count
                    broken_manifest_ids,              # 135 Broken MANIFEST IDs
                    broken_8f78f74b_count,            # 136 Broken 8F78F74B Count
                    broken_8f78f74b_ids,              # 137 Broken 8F78F74B IDs
                    original_source_path,              # 138 Original Source Path
                    ])


                # -------------------------
                # Technical Details Data
                # -------------------------

                if data[15] == "EA Match":
                    ea_status = "EA Match"

                elif data[15] == "EA Override":
                    ea_status = "EA Override"

                elif data[15] == "EA Candidate":
                    ea_status = "EA Candidate"

                elif data[15] == "Custom Content":
                    ea_status = "Custom"

                else:
                    ea_status = "Unknown"

                # Compact report-only fields retained from the former
                # Technical Details sheet. These are appended after the stable
                # MDB fields so all existing database indices remain unchanged.
                rows[-1].extend([
                    priority,
                    risk_reasons,
                    ea_status,
                    data[151] if len(data) > 151 else "",  # 142 structural subtype
                ])

                tech_rows.append([
                    file,
                    priority,
                    risk_reasons,
                    ea_status,

                    data[16],   # EA Matches
                    data[17],   # EA CASP Matches
                    data[18],   # EA GEOM Matches
                    data[19],   # EA IMG Matches

                    data[2],    # All Resolutions

                    data[5],              # CASP Count
                    data[6],              # GEOM Count
                    data[7],              # Slider Count

                    format_size(data[10]), # GEOM Total Size
                    format_size(data[11]), # Largest GEOM
                    data[10],              # GEOM Total Bytes
                    data[12],              # Highest GEOM ID

                    data[3],               # 4K Count

                    data[23],   # Preset Type
                    data[22],   # Preset Count

                    data[25],   # Skin Type

                    data[26],   # CASP Instance IDs

                    data[30],   # STBL IDs
                    data[31],   # BGEO IDs
                    data[32],   # FACE IDs
                    data[33],   # FBLN IDs
                    data[34],   # KEY IDs

                    health_status,       # Health Status
                    broken_geom_count,   # Broken GEOM Count
                    broken_geom_ids,     # Broken GEOM IDs
                    broken_casp_count,   # Broken CASP Count
                    broken_casp_ids,     # Broken CASP IDs
                    broken_img_count,    # Broken IMG Count
                    broken_img_ids,      # Broken IMG IDs
                    health_issues,       # Health Issues
                    broken_vpxy_count,   # Broken VPXY Count
                    broken_vpxy_ids,     # Broken VPXY IDs
                    broken_bone_count,   # Broken BONE Count
                    broken_bone_ids,     # Broken BONE IDs
                    broken_bond_count,   # Broken BOND Count
                    broken_bond_ids,     # Broken BOND IDs
                    broken_morph_count,      # Broken Morph Count
                    broken_morph_ids,        # Broken Morph IDs
                    broken_build_buy_count,  # Broken Build/Buy Count
                    broken_build_buy_ids,    # Broken Build/Buy IDs
                    broken_preset_count,     # Broken PRESET Count
                    broken_preset_ids,       # Broken PRESET IDs
                    broken_tone_count,       # Broken TONE Count
                    broken_tone_ids,         # Broken TONE IDs
                    broken_imag_count,       # Broken IMAG Count
                    broken_imag_ids,         # Broken IMAG IDs
                    broken_txtc_count,       # Broken TXTC Count
                    broken_txtc_ids,         # Broken TXTC IDs
                    broken_ptrn_count,       # Broken PTRN Count
                    broken_ptrn_ids,         # Broken PTRN IDs
                    broken_simo_count,       # Broken SIMO Count
                    broken_simo_ids,         # Broken SIMO IDs
                    broken_txtf_count,       # Broken TXTF Count
                    broken_txtf_ids,         # Broken TXTF IDs
                    broken_famd_count,       # Broken FAMD Count
                    broken_famd_ids,         # Broken FAMD IDs
                    broken_objd_count,       # Broken OBJD Count
                    broken_objd_ids,         # Broken OBJD IDs
                    broken_objk_count,       # Broken OBJK Count
                    broken_objk_ids,         # Broken OBJK IDs
                    broken_xml_count,        # Broken XML Count
                    broken_xml_ids,          # Broken XML IDs
                    broken_stbl_count,       # Broken STBL Count
                    broken_stbl_ids,         # Broken STBL IDs
                    broken_nmap_count,       # Broken NMAP Count
                    broken_nmap_ids,         # Broken NMAP IDs
                    broken_rslt_count,       # Broken RSLT Count
                    broken_rslt_ids,         # Broken RSLT IDs
                    broken_ftpt_count,       # Broken FTPT Count
                    broken_ftpt_ids,         # Broken FTPT IDs
                    broken_rig_count,        # Broken RIG Count
                    broken_rig_ids,          # Broken RIG IDs
                    broken_jazz_count,       # Broken JAZZ Count
                    broken_jazz_ids,         # Broken JAZZ IDs
                    broken_s3sa_count,       # Broken S3SA Count
                    broken_s3sa_ids,         # Broken S3SA IDs
                    broken_lite_count,       # Broken LITE Count
                    broken_lite_ids,         # Broken LITE IDs
                    broken_clip_count,       # Broken CLIP Count
                    broken_clip_ids,         # Broken CLIP IDs
                    broken_matd_count,       # Broken MATD Count
                    broken_matd_ids,         # Broken MATD IDs
                    broken_mtst_count,       # Broken MTST Count
                    broken_mtst_ids,         # Broken MTST IDs
                    broken_mdlr_count,    # Broken MDLR Count
                    broken_mdlr_ids,      # Broken MDLR IDs
                    broken_cwal_count,    # Broken CWAL Count
                    broken_cwal_ids,      # Broken CWAL IDs
                    broken_cfen_count,    # Broken CFEN Count
                    broken_cfen_ids,      # Broken CFEN IDs
                    broken_ctpt_count,    # Broken CTPT Count
                    broken_ctpt_ids,      # Broken CTPT IDs
                    broken_ngmp_count,    # Broken NGMP Count
                    broken_ngmp_ids,      # Broken NGMP IDs
                    broken_crmt_count,    # Broken CRMT Count
                    broken_crmt_ids,      # Broken CRMT IDs
                    broken_spt2_count,    # Broken SPT2 Count
                    broken_spt2_ids,      # Broken SPT2 IDs
                    broken_spt_count,     # Broken _SPT Count
                    broken_spt_ids,       # Broken _SPT IDs
                    broken_cral_count,    # Broken CRAL Count
                    broken_cral_ids,      # Broken CRAL IDs
                    broken_layo_count,    # Broken LAYO Count
                    broken_layo_ids,      # Broken LAYO IDs
                    broken_hinf_count,    # Broken HINF Count
                    broken_hinf_ids,      # Broken HINF IDs
                    broken_2ary_count,    # Broken 2ARY Count
                    broken_2ary_ids,      # Broken 2ARY IDs
                    broken_complate_count,    # Broken Complate Count
                    broken_complate_ids,      # Broken Complate IDs
                    broken_63a33ea7_count,    # Broken 63A33EA7 Count
                    broken_63a33ea7_ids,      # Broken 63A33EA7 IDs
                    broken_package_storage_count,    # Broken Package Storage Count
                    broken_package_storage_ids,      # Broken Package Storage IDs
                    broken_fireplace_count,          # Broken Fireplace Count
                    broken_fireplace_ids,            # Broken Fireplace IDs
                    broken_thum_count,               # Broken THUM Count
                    broken_thum_ids,                 # Broken THUM IDs
                    broken_626f60ce_count,           # Broken 626F60CE Count
                    broken_626f60ce_ids,             # Broken 626F60CE IDs
                    broken_icon_count,               # Broken ICON Count
                    broken_icon_ids,                 # Broken ICON IDs
                    broken_0580a2b5_count,           # Broken THUM 0580A2B5 Count
                    broken_0580a2b5_ids,             # Broken THUM 0580A2B5 IDs
                    broken_0580a2b4_count,           # Broken THUM 0580A2B4 Count
                    broken_0580a2b4_ids,             # Broken THUM 0580A2B4 IDs
                    broken_0580a2b6_count,           # Broken THUM 0580A2B6 Count
                    broken_0580a2b6_ids,             # Broken THUM 0580A2B6 IDs
                    broken_0580a2cd_count,           # Broken SNAP 0580A2CD Count
                    broken_0580a2cd_ids,             # Broken SNAP 0580A2CD IDs
                    broken_0580a2cf_count,           # Broken SNAP 0580A2CF Count
                    broken_0580a2cf_ids,             # Broken SNAP 0580A2CF IDs
                    broken_twni_count,                # Broken TWNI Count
                    broken_twni_ids,                  # Broken TWNI IDs
                    broken_manifest_count,            # Broken MANIFEST Count
                    broken_manifest_ids,              # Broken MANIFEST IDs
                    broken_8f78f74b_count,            # Broken 8F78F74B Count
                    broken_8f78f74b_ids,              # Broken 8F78F74B IDs
                ])


            except Exception as e:

                empty_package_file = False
                try:
                    empty_package_file = os.path.getsize(path) == 0
                except Exception:
                    pass

                error_text = (
                    "Empty package file (0 bytes)"
                    if empty_package_file
                    else str(e)
                )
                foreign_game = (
                    ""
                    if empty_package_file
                    else _foreign_game_from_error_text(error_text)
                )
                sims2_major1 = (
                    (not empty_package_file)
                    and (
                        foreign_game == "SIMS_2"
                        or _is_sims2_major1_error_text(error_text)
                    )
                )
                sims4_dbpf21 = (foreign_game == "SIMS_4")
                error_health_status = (
                    "BROKEN"
                    if empty_package_file
                    else (
                        "UNSUPPORTED GAME"
                        if (sims2_major1 or sims4_dbpf21)
                        else "SCAN ERROR"
                    )
                )
                error_package_hash = (
                    calculate_file_hash(path)
                    if empty_package_file
                    else ""
                )

                if empty_package_file:
                    print("EMPTY PACKAGE HEALTH FAILURE:", file, "| 0 bytes")
                elif sims2_major1:
                    sims2_unsupported_count += 1
                    print("UNSUPPORTED GAME: Sims 2 / DBPF 1.x:", file)
                elif sims4_dbpf21:
                    sims4_unsupported_count += 1
                    print("UNSUPPORTED GAME: Sims 4 / DBPF 2.1:", file)
                else:
                    print("ERROR:", e)

                if DEBUG:
                    import traceback
                    traceback.print_exc()

                rows.append([
                    file,              # 0 Package
                    "",                # 1 CASP IDs
                    "",                # 2 GEOM IDs
                    "",                # 3 IMG IDs
                    "",                # 4 VPXY IDs
                    "",                # 5 Duplicate Files

                    "UNKNOWN",         # 6 CC Type
                    "UNSUPPORTED",     # 7 Content Type
                    "UNKNOWN",         # 8 Risk
                    "N/A",             # 9 Risk Score

                    "",                # 10 HQ Warning
                    "",                # 11 Poly Warning
                    0,                 # 12 Highest Polys
                    "",                # 13 Biggest Texture
                    "",                # 14 Skin Type

                    "",                # 15 STBL IDs
                    "",                # 16 BGEO IDs
                    "",                # 17 FACE IDs
                    "",                # 18 FBLN IDs
                    "",                # 19 KEY IDs
                    0,                 # 20 BONE Resource Count
                    error_package_hash, # 21 Package Hash
                    error_health_status, # 22 Health Status
                    0,                 # 23 Broken GEOM Count
                    "",                # 24 Broken GEOM IDs
                    _condense_package_health_issue(error_text), # 25 Health Issues
                    0,                  # 26 Broken CASP Count
                    "",                 # 27 Broken CASP IDs
                    0,                  # 28 Broken IMG Count
                    "",                 # 29 Broken IMG IDs
                    0,                  # 30 Broken VPXY Count
                    "",                 # 31 Broken VPXY IDs
                    0,                  # 32 Broken BONE Count
                    "",                 # 33 Broken BONE IDs
                    0,                  # 34 Broken BOND Count
                    "",                 # 35 Broken BOND IDs
                    0,                  # 36 Broken Morph Count
                    "",                 # 37 Broken Morph IDs
                    0,                  # 38 Broken Build/Buy Count
                    "",                 # 39 Broken Build/Buy IDs
                    0,                  # 40 Broken PRESET Count
                    "",                # 41 Broken PRESET IDs
                    0,                 # 42 Broken TONE Count
                    "",                # 43 Broken TONE IDs
                    0,                 # 44 Broken IMAG Count
                    "",                # 45 Broken IMAG IDs
                    0,                 # 46 Broken TXTC Count
                    "",                # 47 Broken TXTC IDs
                    0,                 # 48 Broken PTRN Count
                    "",                # 49 Broken PTRN IDs
                    0,                 # 50 Broken SIMO Count
                    "",                # 51 Broken SIMO IDs
                    0,                 # 52 Broken TXTF Count
                    "",                # 53 Broken TXTF IDs
                    0,                 # 54 Broken FAMD Count
                    "",                # 55 Broken FAMD IDs
                    0,                 # 56 Broken OBJD Count
                    "",                # 57 Broken OBJD IDs
                    0,                 # 58 Broken OBJK Count
                    "",                # 59 Broken OBJK IDs
                    0,                 # 60 Broken XML Count
                    "",                # 61 Broken XML IDs
                    0,                 # 62 Broken STBL Count
                    "",                # 63 Broken STBL IDs
                    0,                 # 64 Broken NMAP Count
                    "",                # 65 Broken NMAP IDs
                    0,                 # 66 Broken RSLT Count
                    "",                # 67 Broken RSLT IDs
                    0,                 # 68 Broken FTPT Count
                    "",                # 69 Broken FTPT IDs
                    0,                 # 70 Broken RIG Count
                    "",                # 71 Broken RIG IDs
                    0,                 # 72 Broken JAZZ Count
                    "",                # 73 Broken JAZZ IDs
                    0,                 # 74 Broken S3SA Count
                    "",                # 75 Broken S3SA IDs
                    0,                 # 76 Broken LITE Count
                    "",                # 77 Broken LITE IDs
                    0,                 # 78 Broken CLIP Count
                    "",                # 79 Broken CLIP IDs
                    0,                 # 80 Broken MATD Count
                    "",                # 81 Broken MATD IDs
                    0,                 # 82 Broken MTST Count
                    "",                # 83 Broken MTST IDs
                    0,                 # 84 Broken MDLR Count
                    "",                # 85 Broken MDLR IDs
                    0,                 # 86 Broken CWAL Count
                    "",                # 87 Broken CWAL IDs
                    0,                 # 88 Broken CFEN Count
                    "",                # 89 Broken CFEN IDs
                    0,                 # 90 Broken CTPT Count
                    "",                # 91 Broken CTPT IDs
                    0,                 # 92 Broken NGMP Count
                    "",                # 93 Broken NGMP IDs
                    0,                 # 94 Broken CRMT Count
                    "",                # 95 Broken CRMT IDs
                    0,                 # 96 Broken SPT2 Count
                    "",                # 97 Broken SPT2 IDs
                    0,                 # 98 Broken _SPT Count
                    "",                # 99 Broken _SPT IDs
                    0,                 # 100 Broken CRAL Count
                    "",                # 101 Broken CRAL IDs
                    0,                 # 102 Broken LAYO Count
                    "",                # 103 Broken LAYO IDs
                    0,                 # 104 Broken HINF Count
                    "",                # 105 Broken HINF IDs
                    0,                 # 106 Broken 2ARY Count
                    "",                # 107 Broken 2ARY IDs
                    0,                 # 108 Broken Complate Count
                    "",                # 109 Broken Complate IDs
                    0,                 # 110 Broken 63A33EA7 Count
                    "",                # 111 Broken 63A33EA7 IDs
                    0,                 # 112 Broken Package Storage Count
                    "",                # 113 Broken Package Storage IDs
                    0,                 # 114 Broken Fireplace Count
                    "",                # 115 Broken Fireplace IDs
                    0,                 # 116 Broken THUM Count
                    "",                # 117 Broken THUM IDs
                    0,                 # 118 Broken 626F60CE Count
                    "",                # 119 Broken 626F60CE IDs
                    0,                 # 120 Broken ICON Count
                    "",                # 121 Broken ICON IDs
                    0,                 # 122 Broken THUM 0580A2B5 Count
                    "",                # 123 Broken THUM 0580A2B5 IDs
                    0,                 # 124 Broken THUM 0580A2B4 Count
                    "",                # 125 Broken THUM 0580A2B4 IDs
                    0,                 # 126 Broken THUM 0580A2B6 Count
                    "",                # 127 Broken THUM 0580A2B6 IDs
                    0,                 # 128 Broken SNAP 0580A2CD Count
                    "",                # 129 Broken SNAP 0580A2CD IDs
                    0,                 # 130 Broken SNAP 0580A2CF Count
                    "",                # 131 Broken SNAP 0580A2CF IDs
                    0,                 # 132 Broken TWNI Count
                    "",                # 133 Broken TWNI IDs
                    0,                 # 134 Broken MANIFEST Count
                    "",                # 135 Broken MANIFEST IDs
                    0,                 # 136 Broken 8F78F74B Count
                    "",                # 137 Broken 8F78F74B IDs
                    original_source_path  # 138 Original Source Path
                    ])

                rows[-1].extend([
                    ("BROKEN" if empty_package_file else "ERROR"),
                    _condense_package_health_issue(error_text),
                    "ERROR",
                ])
                                
                tech_rows.append([
                    file,        # 0 Package
                    ("BROKEN" if empty_package_file else "ERROR"), # 1 Priority
                    error_text,  # 2 Risk Reasons
                    "ERROR",     # 3 EA Status
                    0,           # 4 EA Matches
                    0,           # 5 EA CASP Matches
                    0,           # 6 EA GEOM Matches
                    0,           # 7 EA IMG Matches
                    "",          # 8 All Resolutions
                    0,           # 9 CASP Count
                    0,           # 10 GEOM Count
                    0,           # 11 Slider Count
                    "",          # 12 GEOM Total Size
                    "",          # 13 Largest GEOM
                    0,           # 14 GEOM Total Bytes
                    "",          # 15 Highest Poly GEOM ID
                    0,           # 16 4K Count
                    "",          # 17 Preset Type
                    0,           # 18 Preset Count
                    "",          # 19 Skin Type
                    "",          # 20 CASP Instance IDs
                    "",          # 21 STBL IDs
                    "",          # 22 BGEO IDs
                    "",          # 23 FACE IDs
                    "",          # 24 FBLN IDs
                    "",          # 25 KEY IDs
                    error_health_status, # 26 Health Status
                    0,           # 27 Broken GEOM Count
                    "",          # 28 Broken GEOM IDs
                    0,           # 29 Broken CASP Count
                    "",          # 30 Broken CASP IDs
                    0,           # 31 Broken IMG Count
                    "",          # 32 Broken IMG IDs
                    _condense_package_health_issue(error_text), # 33 Health Issues
                    0,           # 34 Broken VPXY Count
                    "",          # 35 Broken VPXY IDs
                    0,           # 36 Broken BONE Count
                    "",          # 37 Broken BONE IDs
                    0,           # 38 Broken BOND Count
                    "",          # 39 Broken BOND IDs
                    0,           # 40 Broken Morph Count
                    "",          # 41 Broken Morph IDs
                    0,           # 42 Broken Build/Buy Count
                    "",          # 43 Broken Build/Buy IDs
                    0,           # 44 Broken PRESET Count
                    "",          # 45 Broken PRESET IDs
                    0,           # 46 Broken TONE Count
                    "",          # 47 Broken TONE IDs
                    0,           # 48 Broken IMAG Count
                    "",          # 49 Broken IMAG IDs
                    0,           # 50 Broken TXTC Count
                    "",          # 51 Broken TXTC IDs
                    0,           # 52 Broken PTRN Count
                    "",          # 53 Broken PTRN IDs
                    0,           # 54 Broken SIMO Count
                    "",          # 55 Broken SIMO IDs
                    0,           # 56 Broken TXTF Count
                    "",          # 57 Broken TXTF IDs
                    0,           # 58 Broken FAMD Count
                    "",          # 59 Broken FAMD IDs
                    0,           # 60 Broken OBJD Count
                    "",          # 61 Broken OBJD IDs
                    0,           # 62 Broken OBJK Count
                    "",          # 63 Broken OBJK IDs
                    0,           # 64 Broken XML Count
                    "",          # 65 Broken XML IDs
                    0,           # 66 Broken STBL Count
                    "",          # 67 Broken STBL IDs
                    0,           # 68 Broken NMAP Count
                    "",          # 69 Broken NMAP IDs
                    0,           # 70 Broken RSLT Count
                    "",          # 71 Broken RSLT IDs
                    0,           # 72 Broken FTPT Count
                    "",          # 73 Broken FTPT IDs
                    0,           # 74 Broken RIG Count
                    "",          # 75 Broken RIG IDs
                    0,           # 76 Broken JAZZ Count
                    "",          # 77 Broken JAZZ IDs
                    0,           # 78 Broken S3SA Count
                    "",          # 79 Broken S3SA IDs
                    0,           # 80 Broken LITE Count
                    "",          # 81 Broken LITE IDs
                    0,           # 82 Broken CLIP Count
                    "",          # 83 Broken CLIP IDs
                    0,           # 84 Broken MATD Count
                    "",          # 85 Broken MATD IDs
                    0,           # 86 Broken MTST Count
                    "",          # 87 Broken MTST IDs
                    0,           # 88 Broken MDLR Count
                    "",          # 89 Broken MDLR IDs
                    0,           # 90 Broken CWAL Count
                    "",          # 91 Broken CWAL IDs
                    0,           # 92 Broken CFEN Count
                    "",          # 93 Broken CFEN IDs
                    0,           # 94 Broken CTPT Count
                    "",          # 95 Broken CTPT IDs
                    0,           # 96 Broken NGMP Count
                    "",          # 97 Broken NGMP IDs
                    0,           # 98 Broken CRMT Count
                    "",          # 99 Broken CRMT IDs
                    0,           # 100 Broken SPT2 Count
                    "",          # 101 Broken SPT2 IDs
                    0,           # 102 Broken _SPT Count
                    "",          # 103 Broken _SPT IDs
                    0,           # 104 Broken CRAL Count
                    "",          # 105 Broken CRAL IDs
                    0,           # 106 Broken LAYO Count
                    "",          # 107 Broken LAYO IDs
                    0,           # 108 Broken HINF Count
                    "",          # 109 Broken HINF IDs
                    0,           # 110 Broken 2ARY Count
                    "",          # 111 Broken 2ARY IDs
                    0,           # 112 Broken Complate Count
                    "",          # 113 Broken Complate IDs
                    0,           # 114 Broken 63A33EA7 Count
                    "",          # 115 Broken 63A33EA7 IDs
                    0,           # 116 Broken Package Storage Count
                    "",          # 117 Broken Package Storage IDs
                    0,           # 118 Broken Fireplace Count
                    "",          # 119 Broken Fireplace IDs
                    0,           # 120 Broken THUM Count
                    "",          # 121 Broken THUM IDs
                    0,           # 122 Broken 626F60CE Count
                    "",          # 123 Broken 626F60CE IDs
                    0,           # 124 Broken ICON Count
                    "",          # 125 Broken ICON IDs
                    0,           # 126 Broken THUM 0580A2B5 Count
                    "",          # 127 Broken THUM 0580A2B5 IDs
                    0,           # 128 Broken THUM 0580A2B4 Count
                    "",          # 129 Broken THUM 0580A2B4 IDs
                    0,           # 130 Broken THUM 0580A2B6 Count
                    "",          # 131 Broken THUM 0580A2B6 IDs
                    0,           # 132 Broken SNAP 0580A2CD Count
                    "",          # 133 Broken SNAP 0580A2CD IDs
                    0,           # 134 Broken SNAP 0580A2CF Count
                    "",          # 135 Broken SNAP 0580A2CF IDs
                    0,           # 136 Broken TWNI Count
                    "",          # 137 Broken TWNI IDs
                    0,           # 138 Broken MANIFEST Count
                    "",          # 139 Broken MANIFEST IDs
                    0,           # 140 Broken 8F78F74B Count
                    ""           # 141 Broken 8F78F74B IDs
                ])

                error_route_category = (
                    os.path.join("BROKEN_CC", "PACKAGE")
                    if empty_package_file
                    else (
                        os.path.join("UNSUPPORTED", "SIMS_2")
                        if sims2_major1
                        else (
                            os.path.join("UNSUPPORTED", "SIMS_4")
                            if sims4_dbpf21
                            else os.path.join("UNSUPPORTED", "OTHER_OR_UNKNOWN")
                        )
                    )
                )
                files_to_move.append((path, error_route_category))
                if ROUTING_PREVIEW:
                    routing_preview_records.append((path, error_route_category, "No"))

# NO INDENT HERE - scan loop finished


# Preserve the selected source-folder name for image passthrough routing.
# This avoids flattening loose images from individually selected creator/set folders
# into the shared IMAGES root while still preserving any inner relative folders.
image_source_folder_name = os.path.basename(os.path.normpath(folder)) or "SOURCE"

# -------------------------
# ROUTING PREVIEW (DRY RUN)
# -------------------------

if ROUTING_PREVIEW:
    preview_output = os.path.join(scan_session_folder, "Routing_Preview.txt")
    preview_lines = []
    print("")
    print("ROUTING PREVIEW: DRY RUN - NO SOURCE FILES WILL BE MOVED")
    print("ROUTING PREVIEW ROOT:", MOVE_FOLDER)

    print("ROUTING PREVIEW SOURCE RECORDS:", len(routing_preview_records))
    for path, move_category, duplicate_flag in routing_preview_records:
        try:
            # Big Review real routing diverts exact package duplicates into
            # DUPLICATES\<scan session> while preserving the source-relative
            # path. The dry run must preview that *actual* destination rather
            # than showing the package's normal category destination.
            preview_duplicate_route = (
                SCAN_SOURCE_MODE == "BIG_REVIEW_BATCH"
                and duplicate_flag == "Yes"
                and str(path).lower().endswith(".package")
            )

            if preview_duplicate_route:
                duplicate_relative_path = os.path.relpath(path, folder)
                destination = os.path.join(
                    DUPLICATES_FOLDER,
                    scan_folder_name,
                    duplicate_relative_path
                )
            elif move_category in ("SIMS3PACK", "ARCHIVES", "IMAGES"):
                passthrough_relative_path = os.path.relpath(path, folder)
                if move_category == "SIMS3PACK":
                    passthrough_root = SIMS3PACK_FOLDER
                elif move_category == "ARCHIVES":
                    passthrough_root = ARCHIVES_FOLDER
                else:
                    passthrough_root = IMAGES_FOLDER
                if move_category == "IMAGES":
                    destination = os.path.join(
                        passthrough_root,
                        image_source_folder_name,
                        passthrough_relative_path
                    )
                else:
                    destination = os.path.join(
                        passthrough_root,
                        passthrough_relative_path
                    )
            else:
                destination = os.path.join(
                    MOVE_FOLDER,
                    move_category,
                    os.path.basename(path)
                )

            if os.path.exists(destination):
                note = "DESTINATION EXISTS (DUPLICATE ROUTE)" if preview_duplicate_route else "DESTINATION EXISTS"
            else:
                note = "READY (DUPLICATE ROUTE)" if preview_duplicate_route else "READY"
            print(
                "ROUTING PREVIEW ITEM\t{}\t{}\t{}\t{}\t{}".format(
                    str(move_category).replace("\t", " "),
                    duplicate_flag,
                    str(path).replace("\t", " "),
                    str(destination).replace("\t", " "),
                    note,
                )
            )
            preview_lines.append(
                "{}\t{}\t{}\t{}\t{}".format(move_category, duplicate_flag, path, destination, note)
            )
        except Exception as e:
            print("ROUTING PREVIEW ERROR:", path, e)

    try:
        with open(preview_output, "w", encoding="utf-8") as preview_file:
            preview_file.write("Category\tDuplicate\tSource\tDestination\tStatus\n")
            for preview_line in preview_lines:
                preview_file.write(preview_line + "\n")
        print("Routing preview saved:", preview_output)
        print("ROUTING PREVIEW COUNT:", len(preview_lines))
    except Exception as e:
        print("ROUTING PREVIEW SAVE ERROR:", e)

# -------------------------
# MOVE FILES AFTER SCANNING
# -------------------------

move_manifest_records = []

if MOVE_FILES:

    import clr
    from System import GC

    GC.Collect()
    GC.WaitForPendingFinalizers()
    GC.Collect()

    print("MOVING FILES...")

    for path, move_category in files_to_move:

        try:

            if move_category == "DUPLICATES":
                duplicate_relative_path = os.path.relpath(
                    path,
                    folder
                )
                destination = os.path.join(
                    DUPLICATES_FOLDER,
                    scan_folder_name,
                    duplicate_relative_path
                )
            elif move_category in ("SIMS3PACK", "ARCHIVES", "IMAGES"):
                passthrough_relative_path = os.path.relpath(path, folder)
                if move_category == "SIMS3PACK":
                    passthrough_root = SIMS3PACK_FOLDER
                elif move_category == "ARCHIVES":
                    passthrough_root = ARCHIVES_FOLDER
                else:
                    passthrough_root = IMAGES_FOLDER
                if move_category == "IMAGES":
                    destination = os.path.join(
                        passthrough_root,
                        image_source_folder_name,
                        passthrough_relative_path
                    )
                else:
                    destination = os.path.join(
                        passthrough_root,
                        passthrough_relative_path
                    )
            else:
                destination = os.path.join(
                    MOVE_FOLDER,
                    move_category,
                    os.path.basename(path)
                )

            if os.path.exists(destination):

                print(
                    "SKIPPED DUPLICATE:",
                    os.path.basename(path)
                )

                duplicates.append(
                    (
                        path,
                        destination
                    )
                )
                move_manifest_records.append((path, destination, move_category, "SKIPPED_DESTINATION_EXISTS"))

                continue

            os.makedirs(
                os.path.dirname(destination),
                exist_ok=True
            )

            shutil.move(
                path,
                destination
            )

            if move_category == "DUPLICATES":
                duplicates.append(
                    (
                        path,
                        destination
                    )
                )

            print("MOVED:", destination)
            move_manifest_records.append((path, destination, move_category, "MOVED"))

        except PermissionError:

            print("FILE LOCKED - RETRYING:", path)

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

            time.sleep(0.25)

            try:
                shutil.move(
                    path,
                    destination
                )

                if move_category == "DUPLICATES":
                    duplicates.append(
                        (
                            path,
                            destination
                        )
                    )

                print("MOVED AFTER RETRY:", destination)
                move_manifest_records.append((path, destination, move_category, "MOVED_AFTER_RETRY"))

            except PermissionError:
                print("STILL LOCKED:", path)
                move_manifest_records.append((path, destination, move_category, "FAILED_LOCKED"))

            except Exception as e:
                print("MOVE RETRY FAILED:", path, e)
                move_manifest_records.append((path, destination, move_category, "FAILED_RETRY: {}".format(e)))

        except Exception as e:

            print("MOVE FAILED:", path, e)
            try:
                move_manifest_records.append((path, destination, move_category, "FAILED: {}".format(e)))
            except Exception:
                move_manifest_records.append((path, "", move_category, "FAILED: {}".format(e)))

    move_manifest_path = os.path.join(scan_session_folder, "Move_Manifest.txt")
    try:
        with open(move_manifest_path, "w", encoding="utf-8") as manifest_file:
            manifest_file.write("Source\tDestination\tCategory\tStatus\n")
            for source_path, destination_path, category, status in move_manifest_records:
                manifest_file.write(
                    "{}\t{}\t{}\t{}\n".format(
                        str(source_path).replace("\t", " "),
                        str(destination_path).replace("\t", " "),
                        str(category).replace("\t", " "),
                        str(status).replace("\t", " "),
                    )
                )
        print("Move manifest saved:", move_manifest_path)
        print("MOVE MANIFEST COUNT:", len(move_manifest_records))
    except Exception as e:
        print("MOVE MANIFEST SAVE ERROR:", e)

os.makedirs(
    os.path.dirname(OUTPUT),
    exist_ok=True
)



# Create workbook


wb = Workbook()

summary = wb.active
summary.title = "Scan Summary"

ws = wb.create_sheet("CC Summary")

tech = wb.create_sheet("Technical Details")

duplicates_sheet = wb.create_sheet("Duplicate Filenames")


# -------------------------
# Scan Summary Sheet
# -------------------------

summary["A1"] = "SIMS 3 CC SCAN SUMMARY"
summary["A1"].font = Font(bold=True, size=14)

summary_data = [
    ("Files Scanned", len(rows)),
    ("Unknown / Unidentified CC", sum(1 for _path, category in files_to_move if category == "UNIDENTIFIED_CC")),
    ("Exact SHA256 Duplicates", len(exact_duplicate_paths)),
    ("", ""),

    ("Extreme", sum(1 for r in rows if str(r[8]).upper() == "EXTREME")),
    ("High", sum(1 for r in rows if str(r[8]).upper() == "HIGH")),
    ("Moderate", sum(1 for r in rows if str(r[8]).upper() == "MODERATE")),
    ("Safe", sum(1 for r in rows if str(r[8]).upper() == "SAFE")),

    ("", ""),

    ("HEALTH", ""),
    ("Malformed Packages", sum(1 for r in rows if len(r) > 22 and r[22] == "BROKEN")),
    ("Malformed GEOM Packages", sum(1 for r in rows if len(r) > 23 and int(r[23] or 0) > 0)),
    ("Malformed CASP Packages", sum(1 for r in rows if len(r) > 26 and int(r[26] or 0) > 0)),
    ("Malformed IMG Packages", sum(1 for r in rows if len(r) > 28 and int(r[28] or 0) > 0)),
    ("Malformed VPXY Packages", sum(1 for r in rows if len(r) > 30 and int(r[30] or 0) > 0)),
    ("Malformed BONE Packages", sum(1 for r in rows if len(r) > 32 and int(r[32] or 0) > 0)),
    ("Malformed BOND Packages", sum(1 for r in rows if len(r) > 34 and int(r[34] or 0) > 0)),
    ("Malformed Morph Packages", sum(1 for r in rows if len(r) > 36 and int(r[36] or 0) > 0)),
    ("Malformed Build/Buy Packages", sum(1 for r in rows if len(r) > 38 and int(r[38] or 0) > 0)),
    ("Malformed PRESET Packages", sum(1 for r in rows if len(r) > 40 and int(r[40] or 0) > 0)),
    ("Malformed TONE Packages", sum(1 for r in rows if len(r) > 42 and int(r[42] or 0) > 0)),
    ("Malformed IMAG Packages", sum(1 for r in rows if len(r) > 44 and int(r[44] or 0) > 0)),
    ("Malformed TXTC Packages", sum(1 for r in rows if len(r) > 46 and int(r[46] or 0) > 0)),
    ("Malformed PTRN Packages", sum(1 for r in rows if len(r) > 48 and int(r[48] or 0) > 0)),
    ("Malformed SIMO Packages", sum(1 for r in rows if len(r) > 50 and int(r[50] or 0) > 0)),
    ("Malformed TXTF Packages", sum(1 for r in rows if len(r) > 52 and int(r[52] or 0) > 0)),
    ("Malformed FAMD Packages", sum(1 for r in rows if len(r) > 54 and int(r[54] or 0) > 0)),
    ("Malformed OBJD Packages", sum(1 for r in rows if len(r) > 56 and int(r[56] or 0) > 0)),
    ("Malformed OBJK Packages", sum(1 for r in rows if len(r) > 58 and int(r[58] or 0) > 0)),
    ("Malformed XML Packages", sum(1 for r in rows if len(r) > 60 and int(r[60] or 0) > 0)),
    ("Malformed STBL Packages", sum(1 for r in rows if len(r) > 62 and int(r[62] or 0) > 0)),
    ("Malformed NMAP Packages", sum(1 for r in rows if len(r) > 64 and int(r[64] or 0) > 0)),
    ("Malformed RSLT Packages", sum(1 for r in rows if len(r) > 66 and int(r[66] or 0) > 0)),
    ("Malformed FTPT Packages", sum(1 for r in rows if len(r) > 68 and int(r[68] or 0) > 0)),
    ("Malformed RIG Packages", sum(1 for r in rows if len(r) > 70 and int(r[70] or 0) > 0)),
    ("Malformed JAZZ Packages", sum(1 for r in rows if len(r) > 72 and int(r[72] or 0) > 0)),
    ("Malformed S3SA Packages", sum(1 for r in rows if len(r) > 74 and int(r[74] or 0) > 0)),
    ("Malformed LITE Packages", sum(1 for r in rows if len(r) > 76 and int(r[76] or 0) > 0)),
    ("Malformed CLIP Packages", sum(1 for r in rows if len(r) > 78 and int(r[78] or 0) > 0)),
    ("Malformed MATD Packages", sum(1 for r in rows if len(r) > 80 and int(r[80] or 0) > 0)),
    ("Malformed MTST Packages", sum(1 for r in rows if len(r) > 82 and int(r[82] or 0) > 0)),
    ("Malformed MDLR Packages", sum(1 for r in rows if len(r) > 84 and int(r[84] or 0) > 0)),
    ("Malformed CWAL Packages", sum(1 for r in rows if len(r) > 86 and int(r[86] or 0) > 0)),
    ("Malformed CFEN Packages", sum(1 for r in rows if len(r) > 88 and int(r[88] or 0) > 0)),
    ("Malformed CTPT Packages", sum(1 for r in rows if len(r) > 90 and int(r[90] or 0) > 0)),
    ("Malformed NGMP Packages", sum(1 for r in rows if len(r) > 92 and int(r[92] or 0) > 0)),
    ("Malformed CRMT Packages", sum(1 for r in rows if len(r) > 94 and int(r[94] or 0) > 0)),
    ("Malformed SPT2 Packages", sum(1 for r in rows if len(r) > 96 and int(r[96] or 0) > 0)),
    ("Malformed _SPT Packages", sum(1 for r in rows if len(r) > 98 and int(r[98] or 0) > 0)),
    ("Malformed CRAL Packages", sum(1 for r in rows if len(r) > 100 and int(r[100] or 0) > 0)),
    ("Malformed LAYO Packages", sum(1 for r in rows if len(r) > 102 and int(r[102] or 0) > 0)),
    ("Malformed HINF Packages", sum(1 for r in rows if len(r) > 104 and int(r[104] or 0) > 0)),
    ("Malformed 2ARY Packages", sum(1 for r in rows if len(r) > 106 and int(r[106] or 0) > 0)),
    ("Malformed Complate Packages", sum(1 for r in rows if len(r) > 108 and int(r[108] or 0) > 0)),
    ("Malformed 63A33EA7 Packages", sum(1 for r in rows if len(r) > 110 and int(r[110] or 0) > 0)),
    ("Malformed Package Storage Packages", sum(1 for r in rows if len(r) > 112 and int(r[112] or 0) > 0)),
    ("Malformed Fireplace Packages", sum(1 for r in rows if len(r) > 114 and int(r[114] or 0) > 0)),
    ("Malformed THUM Packages", sum(1 for r in rows if len(r) > 116 and int(r[116] or 0) > 0)),
    ("Malformed 626F60CE Packages", sum(1 for r in rows if len(r) > 118 and int(r[118] or 0) > 0)),
    ("Malformed ICON Packages", sum(1 for r in rows if len(r) > 120 and int(r[120] or 0) > 0)),
    ("Malformed THUM 0580A2B5 Packages", sum(1 for r in rows if len(r) > 122 and int(r[122] or 0) > 0)),
    ("Malformed THUM 0580A2B4 Packages", sum(1 for r in rows if len(r) > 124 and int(r[124] or 0) > 0)),
    ("Malformed THUM 0580A2B6 Packages", sum(1 for r in rows if len(r) > 126 and int(r[126] or 0) > 0)),
    ("Malformed SNAP 0580A2CD Packages", sum(1 for r in rows if len(r) > 128 and int(r[128] or 0) > 0)),
    ("Malformed SNAP 0580A2CF Packages", sum(1 for r in rows if len(r) > 130 and int(r[130] or 0) > 0)),
    ("Malformed TWNI Packages", sum(1 for r in rows if len(r) > 132 and int(r[132] or 0) > 0)),
    ("Malformed MANIFEST Packages", sum(1 for r in rows if len(r) > 134 and int(r[134] or 0) > 0)),
    ("Malformed 8F78F74B Packages", sum(1 for r in rows if len(r) > 136 and int(r[136] or 0) > 0)),

    ("", ""),

    ("CONTENT TYPE", ""),
    ("EA Override", sum(1 for r in rows if r[7] == "EA Override")),
    ("EA Candidate", sum(1 for r in rows if r[7] == "EA Candidate")),
    ("EA Match", sum(1 for r in rows if r[7] == "EA Match")),
    ("Custom", sum(1 for r in rows if r[7] == "Custom Content")),

    ("", ""),

    ("HQ WARNINGS", ""),
    ("4K Textures", sum(1 for r in rows if r[10] == "4K TEXTURE")),
    ("2048 Textures", sum(1 for r in rows if r[10] == "2048 HIGH")),

    ("Heavy Mesh", sum(1 for r in rows if r[11] == "HEAVY")),
    ("Very Heavy Mesh", sum(1 for r in rows if r[11] == "VERY HEAVY")),
]
for item in summary_data:
    summary.append(item)

summary.column_dimensions["A"].width = 35
# -------------------------
# Style Scan Summary
# -------------------------

for row in summary.iter_rows():
    for cell in row:
        if cell.value in ["OVERALL RISK", "HEALTH", "CONTENT TYPE", "HQ WARNINGS"]:
            cell.font = Font(bold=True)


# Colour fills

red = PatternFill("solid", fgColor=MASTER_DB_FAINT_RED)
orange = PatternFill("solid", fgColor=MASTER_DB_FAINT_ORANGE)
yellow = PatternFill("solid", fgColor=MASTER_DB_FAINT_YELLOW)


# Risk colours

for row in summary.iter_rows():

    label = row[0].value
    cell = row[1]

    if label == "Extreme":
        cell.fill = red

    elif label == "High":
        cell.fill = orange

    elif label == "Moderate":
        cell.fill = yellow

    elif label in [
        "Malformed Packages",
        "Malformed GEOM Packages",
        "Malformed CASP Packages",
        "Malformed IMG Packages",
        "Malformed VPXY Packages",
        "Malformed BONE Packages",
        "Malformed BOND Packages",
        "Malformed Morph Packages",
        "Malformed Build/Buy Packages"
    ] and cell.value:
        cell.fill = red


headers = [
    "Package",
    "CASP Instance IDs",
    "CC Type",
    "Content Type",
    "Health Status",
    "Overall Risk",
    "Risk Score",
    "Priority",
    "Risk Reasons",
    "EA Status",
    "Highest Poly Count",
    "Poly Warning",
    "Biggest Texture",
    "HQ Warning",
    "Skin Type"
]


tech_headers = [
    "Package",
    "Priority",
    "Risk Reasons",
    "EA Status",
    "EA Matches",
    "EA CASP Matches",
    "EA GEOM Matches",
    "EA IMG Matches",
    "All Resolutions",
    "CASP Count",
    "GEOM Count",
    "Slider Count",
    "GEOM Total Size",
    "Largest GEOM",
    "GEOM Total Bytes",
    "Highest Poly GEOM ID",
    "4K Count",
    "Preset Type",
    "Preset Count",
    "Skin Type",
    "CASP Instance IDs",
    "STBL IDs",
    "BGEO IDs",
    "FACE IDs",
    "FBLN IDs",
    "KEY IDs",
    "Health Status",
    "Broken GEOM Count",
    "Broken GEOM IDs",
    "Broken CASP Count",
    "Broken CASP IDs",
    "Broken IMG Count",
    "Broken IMG IDs",
    "Health Issues",
    "Broken VPXY Count",
    "Broken VPXY IDs",
    "Broken BONE Count",
    "Broken BONE IDs",
    "Broken BOND Count",
    "Broken BOND IDs",
    "Broken Morph Count",
    "Broken Morph IDs",
    "Broken Build/Buy Count",
    "Broken Build/Buy IDs",
    "Broken PRESET Count",
    "Broken PRESET IDs",
    "Broken TONE Count",
    "Broken TONE IDs",
    "Broken IMAG Count",
    "Broken IMAG IDs",
    "Broken TXTC Count",
    "Broken TXTC IDs",
    "Broken PTRN Count",
    "Broken PTRN IDs",
    "Broken SIMO Count",
    "Broken SIMO IDs",
    "Broken TXTF Count",
    "Broken TXTF IDs",
    "Broken FAMD Count",
    "Broken FAMD IDs",
    "Broken OBJD Count",
    "Broken OBJD IDs",
    "Broken OBJK Count",
    "Broken OBJK IDs",
    "Broken XML Count",
    "Broken XML IDs",
    "Broken STBL Count",
    "Broken STBL IDs",
    "Broken NMAP Count",
    "Broken NMAP IDs",
    "Broken RSLT Count",
    "Broken RSLT IDs",
    "Broken FTPT Count",
    "Broken FTPT IDs",
    "Broken RIG Count",
    "Broken RIG IDs",
    "Broken JAZZ Count",
    "Broken JAZZ IDs",
    "Broken S3SA Count",
    "Broken S3SA IDs",
    "Broken LITE Count",
    "Broken LITE IDs",
    "Broken CLIP Count",
    "Broken CLIP IDs",
    "Broken MATD Count",
    "Broken MATD IDs",
    "Broken MTST Count",
    "Broken MTST IDs",
    "Broken MDLR Count",
    "Broken MDLR IDs",
    "Broken CWAL Count",
    "Broken CWAL IDs",
    "Broken CFEN Count",
    "Broken CFEN IDs",
    "Broken CTPT Count",
    "Broken CTPT IDs",
    "Broken NGMP Count",
    "Broken NGMP IDs",
    "Broken CRMT Count",
    "Broken CRMT IDs",
    "Broken SPT2 Count",
    "Broken SPT2 IDs",
    "Broken _SPT Count",
    "Broken _SPT IDs",
    "Broken CRAL Count",
    "Broken CRAL IDs",
    "Broken LAYO Count",
    "Broken LAYO IDs",
    "Broken HINF Count",
    "Broken HINF IDs",
    "Broken 2ARY Count",
    "Broken 2ARY IDs",
    "Broken Complate Count",
    "Broken Complate IDs",
    "Broken 63A33EA7 Count",
    "Broken 63A33EA7 IDs",
    "Broken Package Storage Count",
    "Broken Package Storage IDs",
    "Broken Fireplace Count",
    "Broken Fireplace IDs",
    "Broken THUM Count",
    "Broken THUM IDs",
    "Broken 626F60CE Count",
    "Broken 626F60CE IDs",
    "Broken ICON Count",
    "Broken ICON IDs",
    "Broken THUM 0580A2B5 Count",
    "Broken THUM 0580A2B5 IDs",
    "Broken THUM 0580A2B4 Count",
    "Broken THUM 0580A2B4 IDs",
    "Broken THUM 0580A2B6 Count",
    "Broken THUM 0580A2B6 IDs",
    "Broken SNAP 0580A2CD Count",
    "Broken SNAP 0580A2CD IDs",
    "Broken SNAP 0580A2CF Count",
    "Broken SNAP 0580A2CF IDs",
    "Broken TWNI Count",
    "Broken TWNI IDs",
    "Broken MANIFEST Count",
    "Broken MANIFEST IDs",
    "Broken 8F78F74B Count",
    "Broken 8F78F74B IDs",
]


duplicates_sheet.append([
    "Package",
    "Original Location",
    "Existing Destination"
])

# -------------------------
# SORT SUMMARY DATA
# -------------------------

risk_order = {
    "EXTREME": 0,
    "HIGH": 1,
    "MODERATE": 2,
    "SAFE": 3
}

cc_type_order = {
    "SKIN": 0,
    "HAIR": 1,
    "TOP": 2,
    "BOTTOM": 3,
    "BODY": 4,
    "SHOES": 5,
    "ACCESSORY": 6,
    "MAKEUP": 7,
    "OBJECT": 8,
    "BUILD/BUY": 9,
    "UNKNOWN": 10
}

rows.sort(
    key=lambda r: (
        0 if len(r) > 22 and r[22] == "BROKEN" else 1,  # Broken health first
        risk_order.get(str(r[8]).upper(), 99),           # Overall Risk
        cc_type_order.get(str(r[6]).upper(), 99),        # CC Type
        -int(r[12] or 0),                                # Highest Poly Count
        r[0].lower()
    )
)

# Add headers FIRST

ws.append(headers)
tech.append(tech_headers)


# Add data SECOND
# Display only selected columns.
# Keep full rows for Master Database.

for row in rows:

    ws.append([
        row[0],   # Package
        row[1],   # CASP Instance IDs
        _display_cc_type_with_subtype(
            row[6],
            detect_subtype(
                row[0],
                row[6],
                detect_package_format(row[0]),
                row[142] if len(row) > 142 else ""
            )
        ),       # CC Type + optional Subtype
        _display_content_type(row[7]),   # Content Type (display)
        _display_health_status(row[22]),  # Health Status (display)
        row[8],   # Overall Risk
        row[9],   # Risk Score
        _display_health_status(row[139]) if len(row) > 139 else "",  # Priority
        row[140] if len(row) > 140 else "",  # Risk Reasons
        row[141] if len(row) > 141 else _display_content_type(row[7]),  # EA Status
        row[12],  # Highest Poly Count
        row[11],  # Poly Warning
        row[13],  # Biggest Texture
        row[10],  # HQ Warning
        row[14],  # Skin Type
    ])


for row in tech_rows:
    display_row = list(row)
    if len(display_row) > 26:
        display_row[26] = _display_health_status(display_row[26])
    tech.append(display_row)
for original, destination in duplicates:

    duplicates_sheet.append([
        os.path.basename(original),
        original,
        destination
    ])
      
# Header style - report tables

_apply_table_header_divider(ws)
_apply_table_header_divider(tech)
_apply_table_header_divider(duplicates_sheet)

# LibreOffice does not reliably choose useful widths for this report sheet.
# Keep filenames readable and give both filesystem paths enough room.
duplicates_sheet.column_dimensions["A"].width = 48
duplicates_sheet.column_dimensions["B"].width = 80
duplicates_sheet.column_dimensions["C"].width = 72
for duplicate_row in duplicates_sheet.iter_rows():
    for duplicate_cell in duplicate_row:
        duplicate_cell.alignment = Alignment(
            horizontal="left", vertical="top", wrap_text=True
        )


# Align all cells left - both sheets

for sheet in [ws, tech]:

    for row in sheet.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                horizontal="left",
                vertical="top",
                wrap_text=True
            )

_apply_table_header_divider(ws)
_apply_table_header_divider(tech)
_apply_table_header_divider(duplicates_sheet)

# Colours

red = PatternFill("solid", fgColor=MASTER_DB_FAINT_RED)
orange = PatternFill("solid", fgColor=MASTER_DB_FAINT_ORANGE)
yellow = PatternFill("solid", fgColor=MASTER_DB_FAINT_YELLOW)

# Priority colours - Technical Details

for row in tech.iter_rows(min_row=2):

    if row[1].value == "URGENT":
        row[1].fill = red

    elif row[1].value == "HIGH":
        row[1].fill = orange

    elif row[1].value == "MODERATE":
        row[1].fill = yellow
        
# Risk Reasons colours - Technical Details

for row in tech.iter_rows(min_row=2):

    reason = str(row[2].value).strip().upper()

    if reason not in ("", "NONE"):
        row[2].fill = orange
        
   # EA Status colours - Technical Details

for row in tech.iter_rows(min_row=2):

    if row[3].value in ("EA Override", "EA Candidate"):
        row[3].fill = orange

    elif row[3].value == "EA Match":
        row[3].fill = yellow     

for row in ws.iter_rows(min_row=2):

    # Health Status (Column E)
    health = str(row[4].value or "").strip().upper()
    if health in ("BROKEN", "MALFORMED"):
        row[4].fill = red

    # Overall Risk (Column F)
    risk = str(row[5].value or "").strip().upper()
    if risk == "EXTREME":
        row[5].fill = red
    elif risk == "HIGH":
        row[5].fill = orange
    elif risk == "MODERATE":
        row[5].fill = yellow

    # Priority (Column H)
    priority_value = str(row[7].value or "").strip().upper()
    if priority_value in ("BROKEN", "MALFORMED", "URGENT"):
        row[7].fill = red
    elif priority_value == "HIGH":
        row[7].fill = orange
    elif priority_value == "MODERATE":
        row[7].fill = yellow

    # Risk Reasons (Column I)
    reason = str(row[8].value or "").strip().upper()
    if reason not in ("", "NONE"):
        row[8].fill = orange

    # EA Status (Column J)
    ea_status_value = str(row[9].value or "").strip()
    if ea_status_value in ("EA Override", "EA Candidate"):
        row[9].fill = orange
    elif ea_status_value == "EA Match":
        row[9].fill = yellow

    # Poly Warning (Column L)
    poly = str(row[11].value or "").strip().upper()
    if poly == "VERY HEAVY":
        row[11].fill = red
    elif poly == "HEAVY":
        row[11].fill = orange

    # HQ Warning (Column N)
    hq = str(row[13].value or "").strip().upper()
    if hq == "4K TEXTURE":
        row[13].fill = red
    elif hq == "2048 HIGH":
        row[13].fill = orange

    # LOW and SAFE stay white
   
    # ADD PRESET TYPE COLOURS HERE

for row in tech.iter_rows(min_row=2):

    preset = str(row[17].value).strip().upper()

    if preset in [
        "EYE PRESET",
        "NOSE PRESET",
        "LIP PRESET",
        "FACE PRESET"
    ]:
        row[17].fill = yellow

# Add filters to both sheets

ws.auto_filter.ref = ws.dimensions
tech.auto_filter.ref = tech.dimensions

# Column widths - Summary sheet

summary_sizes = {
    "A": 55,   # Package
    "B": 24,   # CASP Instance IDs
    "C": 20,   # CC Type + optional Subtype
    "D": 13,   # Content Type
    "E": 14,   # Health Status
    "F": 14,   # Overall Risk
    "G": 12,   # Risk Score
    "H": 12,   # Priority
    "I": 46,   # Risk Reasons
    "J": 14,   # EA Status
    "K": 18,   # Highest Poly Count
    "L": 15,   # Poly Warning
    "M": 18,   # Biggest Texture
    "N": 14,   # HQ Warning
    "O": 18    # Skin Type
}


for col, width in summary_sizes.items():
    ws.column_dimensions[col].width = width



# Column widths - Technical Details sheet

tech_sizes = {
    "L": 12,   # Slider Count
    "M": 14,   # GEOM Total Size
    "N": 14,   # Largest GEOM
    "O": 15,   # GEOM Total Bytes
    "P": 22,   # Highest Poly GEOM ID
    "Q": 8,    # 4K Count
    "R": 18,   # Preset Type
    "S": 12,   # Preset Count
    "T": 20,   # Skin Type
    "U": 28,   # CASP Instance IDs
    "V": 45,   # STBL IDs
    "W": 45,   # BGEO IDs
    "X": 45,   # FACE IDs
    "Y": 45,   # FBLN IDs
    "Z": 45    # KEY IDs
}

for col, width in tech_sizes.items():
    tech.column_dimensions[col].width = width

for column_cells in tech.columns:
    max_length = 0
    column_letter = get_column_letter(column_cells[0].column)

    if column_letter in tech_sizes:
        continue

    for cell in column_cells:
        if cell.value:
            length = len(str(cell.value))

            if length > max_length:
                max_length = length

    tech.column_dimensions[column_letter].width = min(max_length + 2, 50)

# -------------------------
# FREEZE FIRST COLUMN AND HEADER
# -------------------------

ws.freeze_panes = "A2"
tech.freeze_panes = "A2"

# -------------------------
# CASP ID COLUMN FORMATTING
# -------------------------
for cell in ws["B"]:
    cell.alignment = Alignment(
        wrap_text=True,
        vertical="top"
    )

ws.column_dimensions["B"].width = 22

for r in rows:
    print(
        "FINAL SUMMARY:",
        r[0],   # filename
        _display_health_status(r[22]),  # health status (display)
        r[8],   # risk level
        r[9]    # risk score
    )

# The Master DB is the permanent detailed record. Keep the per-run workbook
# intentionally compact: Scan Summary, CC Summary and Duplicate Filenames only.
if "Technical Details" in wb.sheetnames:
    wb.remove(wb["Technical Details"])

wb.save(OUTPUT)
# -------------------------
# UPDATE MASTER DATABASE
# -------------------------
print("CASP RECORDS FOUND:", len(all_casp_records))
print("EXACT SHA256 DUPLICATES ROUTED:", len(exact_duplicate_paths))
print(
    "THUMBNAIL SUMMARY:",
    f"extracted={thumbnail_extracted_count:,}",
    f"reused={thumbnail_reused_count:,}",
    f"no dedicated preview={thumbnail_missing_count:,}",
    f"errors={thumbnail_error_count:,}",
    f"folder={THUMBNAILS_FOLDER}"
)

if (
    sims2_unsupported_count
    or sims4_unsupported_count
    or exact_duplicate_paths
    or sims3pack_passthrough_count
    or archive_passthrough_count
    or image_passthrough_count
):
    print("")
    print("ROUTING SUMMARY:")
    if sims2_unsupported_count:
        print(
            f"  Sims 2 packages found: {sims2_unsupported_count:,} "
            f"-> sent to UNSUPPORTED\\SIMS_2"
        )
    if sims4_unsupported_count:
        print(
            f"  Sims 4 packages found: {sims4_unsupported_count:,} "
            f"-> sent to UNSUPPORTED\\SIMS_4"
        )
    if exact_duplicate_paths:
        print(
            f"  Sims 3 exact SHA256 duplicates found: {len(exact_duplicate_paths):,} "
            f"-> sent to DUPLICATES\\<scan session>"
        )
    if sims3pack_passthrough_count:
        print(
            f"  Sims3Pack files found: {sims3pack_passthrough_count:,} "
            f"-> moved intact to SIMS3PACK"
        )
    if archive_passthrough_count:
        print(
            f"  Archive files found: {archive_passthrough_count:,} "
            f"-> moved intact to ARCHIVES"
        )
    if image_passthrough_count:
        print(
            f"  Image files found: {image_passthrough_count:,} "
            f"-> moved intact to IMAGES"
        )
    if sims2_unsupported_count or sims4_unsupported_count:
        print("  Sims 2 / Sims 4 packages were not added to the Sims 3 Master Database.")
    if sims3pack_passthrough_count or archive_passthrough_count or image_passthrough_count:
        print("  Non-package files were moved intact; no archives were extracted.")
    print("")

if DEBUG:
    print(all_casp_records[:3])
    
    
if DEBUG:
    print("========== MASTER DATABASE DEBUG ==========")
    print("ROWS TO UPDATE:", len(rows))

    for r in rows:
        print(
            "ADDING:",
            r[0]
        )

    print("==========================================")
    


master_db_rows = [
    row for row in rows
    if not _row_is_foreign_game_unsupported(row)
]
excluded_foreign_rows = len(rows) - len(master_db_rows)
if excluded_foreign_rows:
    print(
        f"MASTER DB: excluding {excluded_foreign_rows:,} known Sims 2/Sims 4 "
        f"package(s) from the Sims 3 database."
    )

if SKIP_MDB_UPDATE:
    print(
        f"MASTER DB: DRY RUN - update skipped for {len(master_db_rows):,} packages / "
        f"{len(all_casp_records):,} CASP records."
    )
else:
    update_master_database(
        master_db_rows,
        all_casp_records
    )


print("Finished! Scan summary saved:", OUTPUT)
print("Console log saved:", LOG_OUTPUT)


# -------------------------
# SCAN TIMER
# -------------------------

elapsed = int(time.time() - scan_start)

minutes = elapsed // 60
seconds = elapsed % 60

print(
    f"Scan time: {minutes} minutes {seconds} seconds"
)


# -------------------------
# REMOVE SINGLE INSTANCE LOCK

cleanup_lock_file()
