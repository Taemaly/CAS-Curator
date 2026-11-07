from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple, Union
import hashlib
import json
import os
import shutil
import sqlite3
import tempfile
import time
import zipfile

from mdb_reader import (
    MDBData,
    MDBWorkbook,
    CURRENT_PERFORMANCE_SCORE_VERSION,
    PERFORMANCE_SCORE_VERSION_METADATA_KEY,
    _aggregate_casp_age_gender,
    _aggregate_casp_outfit_categories,
    _load_shared_strings,
    _read_sheet_rows,
    _sheet_paths,
)

SCHEMA_VERSION = 1
DEFAULT_DB_NAME = "CAS_Curator.db"


@dataclass
class SQLiteMigrationResult:
    db_path: Path
    backup_path: Optional[Path]
    source_path: Path
    source_sha256: str
    cc_master_rows: int
    packages_rows: int
    casp_rows: int
    library_profiles: int
    merge_profiles: int
    file_locations: int
    profile_layouts: int
    elapsed_seconds: float
    orphan_casp_skipped: int = 0


@dataclass
class SQLiteCreateResult:
    db_path: Path
    backup_path: Optional[Path]
    elapsed_seconds: float


FRESH_LIBRARY_HEADERS = [
    "Package", "Status", "Creator", "Converter", "Health Status", "Health Issues",
    "CC Type", "Content Type", "Subtype", "Package Format", "CASP Instance IDs",
    "GEOM Instance IDs", "IMG Instance IDs", "VPXY Instance IDs", "STBL IDs", "BGEO IDs",
    "FACE IDs", "FBLN IDs", "KEY IDs", "Risk", "Risk Score", "HQ Warning",
    "Poly Warning", "Highest Polys", "Biggest Texture", "Age", "Gender",
    "Outfit Category", "Skin Type", "EA Resource Status", "File Location",
    "CC Magic Profiles", "Merge Profiles", "Notes", "Duplicate Files", "Last Scanned",
    "Package Hash", "Original Source Path", "Texture Resolutions", "GEOM Total Bytes",
    "4K Texture Count", "Performance Data Version",
]


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _clean_json_dict(record: Dict[str, object]) -> str:
    return json.dumps(
        {str(key): "" if value is None else str(value) for key, value in record.items()},
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def _rows_for_sheet(path: Path, sheet_name: str) -> Tuple[List[str], List[Tuple[int, Dict[str, str]]]]:
    """Return header names plus non-empty workbook rows from one sheet.

    This intentionally reads OOXML directly, matching the existing fast MDB reader,
    so migration does not depend on openpyxl and does not modify the workbook.
    """
    with zipfile.ZipFile(path, "r") as zf:
        shared = _load_shared_strings(zf)
        sheets = _sheet_paths(zf)
        sheet_path = sheets.get(sheet_name)
        if not sheet_path:
            return [], []

        headers_by_col: Dict[int, str] = {}
        headers: List[str] = []
        rows: List[Tuple[int, Dict[str, str]]] = []
        for row_num, values in _read_sheet_rows(zf, sheet_path, shared):
            if row_num == 1:
                headers_by_col = {col: str(value or "").strip() for col, value in values.items() if str(value or "").strip()}
                headers = [headers_by_col[col] for col in sorted(headers_by_col)]
                continue
            if not headers_by_col:
                continue
            record: Dict[str, str] = {}
            for col, header in headers_by_col.items():
                value = values.get(col, "")
                if value not in (None, ""):
                    record[header] = str(value)
            if not record:
                continue
            rows.append((row_num, record))
        return headers, rows


def _first(record: Dict[str, str], *names: str) -> str:
    for name in names:
        value = str(record.get(name, "") or "").strip()
        if value:
            return value
    return ""


def _create_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        PRAGMA foreign_keys = ON;

        CREATE TABLE meta (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );

        CREATE TABLE cc_master (
            id INTEGER PRIMARY KEY,
            excel_row INTEGER NOT NULL,
            package TEXT NOT NULL,
            package_hash TEXT,
            source_path TEXT,
            status TEXT,
            cc_type TEXT,
            health_status TEXT,
            risk TEXT,
            risk_score TEXT,
            library_profiles TEXT,
            merge_profiles TEXT,
            notes TEXT,
            thumbnail_target TEXT,
            record_json TEXT NOT NULL
        );
        CREATE INDEX idx_cc_master_package ON cc_master(package COLLATE NOCASE);
        CREATE INDEX idx_cc_master_hash ON cc_master(package_hash);
        CREATE INDEX idx_cc_master_source_path ON cc_master(source_path COLLATE NOCASE);
        CREATE INDEX idx_cc_master_cc_type ON cc_master(cc_type COLLATE NOCASE);

        CREATE TABLE packages (
            id INTEGER PRIMARY KEY,
            excel_row INTEGER NOT NULL,
            package TEXT,
            package_hash TEXT,
            source_path TEXT,
            record_json TEXT NOT NULL
        );
        CREATE INDEX idx_packages_package ON packages(package COLLATE NOCASE);
        CREATE INDEX idx_packages_hash ON packages(package_hash);
        CREATE INDEX idx_packages_source_path ON packages(source_path COLLATE NOCASE);

        CREATE TABLE casp_master (
            id INTEGER PRIMARY KEY,
            excel_row INTEGER NOT NULL,
            package TEXT,
            casp_instance TEXT,
            cas_type TEXT,
            record_json TEXT NOT NULL
        );
        CREATE INDEX idx_casp_package ON casp_master(package COLLATE NOCASE);
        CREATE INDEX idx_casp_instance ON casp_master(casp_instance COLLATE NOCASE);
        CREATE INDEX idx_casp_type ON casp_master(cas_type COLLATE NOCASE);

        CREATE TABLE library_profiles (
            name TEXT PRIMARY KEY COLLATE NOCASE
        );
        CREATE TABLE merge_profiles (
            name TEXT PRIMARY KEY COLLATE NOCASE
        );
        CREATE TABLE file_locations (
            name TEXT PRIMARY KEY COLLATE NOCASE,
            folder_path TEXT NOT NULL
        );
        CREATE TABLE profile_layouts (
            id INTEGER PRIMARY KEY,
            profile TEXT NOT NULL,
            package_hash TEXT,
            package_name TEXT,
            file_location TEXT NOT NULL
        );
        CREATE INDEX idx_profile_layout_profile ON profile_layouts(profile COLLATE NOCASE);

        CREATE TABLE workbook_settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        """
    )


def _set_meta(conn: sqlite3.Connection, key: str, value: object) -> None:
    conn.execute("INSERT OR REPLACE INTO meta(key, value) VALUES (?, ?)", (str(key), "" if value is None else str(value)))


def create_fresh_sqlite_library(
    db_path: Union[str, Path],
    headers: Optional[Iterable[str]] = None,
) -> SQLiteCreateResult:
    """Create a brand-new empty CAS Curator SQLite Library safely."""
    started = time.perf_counter()
    target = Path(db_path).resolve()
    if target.suffix.lower() not in {".db", ".sqlite", ".sqlite3"}:
        target = target.with_suffix(".db")
    target.parent.mkdir(parents=True, exist_ok=True)

    registry = []
    for field in (headers or FRESH_LIBRARY_HEADERS):
        clean = str(field or "").strip()
        if clean and clean not in registry:
            registry.append(clean)
    if "Package" not in registry:
        registry.insert(0, "Package")

    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{target.stem}_creating_", suffix=target.suffix, dir=str(target.parent)
    )
    os.close(fd)
    tmp = Path(tmp_name)
    backup_path: Optional[Path] = None
    try:
        if tmp.exists():
            tmp.unlink()
        conn = sqlite3.connect(str(tmp))
        try:
            conn.execute("PRAGMA journal_mode=DELETE")
            conn.execute("PRAGMA synchronous=FULL")
            _create_schema(conn)
            _ensure_primary_library_tables(conn)
            with conn:
                _set_meta(conn, "schema_version", SCHEMA_VERSION)
                _set_meta(conn, "created_utc", datetime.utcnow().replace(microsecond=0).isoformat() + "Z")
                _set_meta(conn, "storage_mode", "sqlite_primary_fresh")
                _set_meta(conn, "source_workbook", "")
                _set_meta(conn, "source_sha256", "")
                _set_meta(conn, "source_size", 0)
                _set_meta(conn, "source_mtime_ns", 0)
                _set_meta(conn, "cc_master_headers_json", json.dumps(registry, ensure_ascii=False))
                _set_meta(conn, "packages_headers_json", "[]")
                _set_meta(conn, "casp_headers_json", "[]")
                # A fresh SQLite Library has no legacy rows. Every score it later
                # receives is produced by this release's current score model, so
                # do not prompt the user to refresh immediately after their first scan.
                _set_meta(conn, PERFORMANCE_SCORE_VERSION_METADATA_KEY, CURRENT_PERFORMANCE_SCORE_VERSION)

            integrity = conn.execute("PRAGMA integrity_check").fetchone()
            if not integrity or str(integrity[0]).casefold() != "ok":
                raise RuntimeError(
                    "SQLite integrity check failed: " + (str(integrity[0]) if integrity else "no result")
                )
            for table in ("cc_master", "packages", "casp_master", "manual_package_fields", "deleted_library_records"):
                count = int(conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])
                if count != 0:
                    raise RuntimeError(f"Fresh database validation failed: {table} contains {count} row(s).")
        finally:
            conn.close()

        if target.exists():
            backup_dir = target.parent / "SQLite_Backups"
            backup_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
            backup_path = backup_dir / f"{target.stem}_before_create_{stamp}{target.suffix}"
            shutil.copy2(target, backup_path)

        os.replace(str(tmp), str(target))
        return SQLiteCreateResult(
            db_path=target,
            backup_path=backup_path,
            elapsed_seconds=time.perf_counter() - started,
        )
    except Exception:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        raise


def migrate_mdb_to_sqlite(
    mdb_path: Union[str, Path],
    db_path: Optional[Union[str, Path]] = None,
) -> SQLiteMigrationResult:
    """Create a validated SQLite snapshot from the current MDB.

    TEST_013 intentionally does *not* switch the scanner or Library to SQLite.
    The source workbook is opened read-only and its SHA-256 is checked again after
    migration.  The SQLite file is built off to the side and atomically installed
    only after validation succeeds.
    """
    started = time.perf_counter()
    source = Path(mdb_path).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if source.suffix.lower() != ".xlsx":
        raise ValueError("SQLite migration currently expects the .xlsx Master Database.")

    target = Path(db_path).resolve() if db_path else source.with_name(DEFAULT_DB_NAME)
    if target.suffix.lower() not in {".db", ".sqlite", ".sqlite3"}:
        target = target.with_suffix(".db")
    target.parent.mkdir(parents=True, exist_ok=True)

    source_hash_before = _sha256_file(source)
    source_stat = source.stat()
    mdb = MDBWorkbook().load(source)

    package_headers, package_rows = _rows_for_sheet(source, MDBWorkbook.PACKAGES_SHEET)
    casp_headers, casp_rows = _rows_for_sheet(source, MDBWorkbook.CASP_SHEET)

    # Keep only real package/CASP records; formatted empty workbook rows are not data.
    package_rows = [item for item in package_rows if _first(item[1], "Package", "Filename", "File Name")]

    # CASP rows are child metadata. Never migrate an orphan CASP that cannot be
    # tied to an actual package represented by CC Master or Packages. This prevents
    # an "empty" old MDB from creating a new SQLite Library with CASP rows but 0 files.
    valid_package_names = {
        _first(rec, "Package", "Filename", "File Name").casefold()
        for rec in mdb.records
        if _first(rec, "Package", "Filename", "File Name")
    }
    valid_package_names.update(
        _first(rec, "Package", "Filename", "File Name").casefold()
        for _, rec in package_rows
        if _first(rec, "Package", "Filename", "File Name")
    )
    raw_casp_rows = [
        item for item in casp_rows
        if _first(item[1], "Package") or _first(item[1], "CASP Instance")
    ]
    casp_rows = [
        item for item in raw_casp_rows
        if _first(item[1], "Package").casefold() in valid_package_names
    ]
    orphan_casp_skipped = len(raw_casp_rows) - len(casp_rows)

    fd, tmp_name = tempfile.mkstemp(prefix=f".{target.stem}_migrating_", suffix=target.suffix, dir=str(target.parent))
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        if tmp.exists():
            tmp.unlink()
        conn = sqlite3.connect(str(tmp))
        try:
            conn.execute("PRAGMA journal_mode=DELETE")
            conn.execute("PRAGMA synchronous=FULL")
            _create_schema(conn)
            with conn:
                _set_meta(conn, "schema_version", SCHEMA_VERSION)
                _set_meta(conn, "created_utc", datetime.utcnow().replace(microsecond=0).isoformat() + "Z")
                _set_meta(conn, "source_workbook", str(source))
                _set_meta(conn, "source_sha256", source_hash_before)
                _set_meta(conn, "source_size", source_stat.st_size)
                _set_meta(conn, "source_mtime_ns", source_stat.st_mtime_ns)
                _set_meta(conn, "storage_mode", "sqlite_migration_test_013")
                _set_meta(conn, "migration_orphan_casp_skipped", orphan_casp_skipped)
                # Preserve the complete workbook column registries, including columns
                # that are currently blank in every record. Reconstructing headers
                # from record_json alone drops all-blank fields such as Status.
                _set_meta(conn, "cc_master_headers_json", json.dumps(list(mdb.headers), ensure_ascii=False))
                _set_meta(conn, "packages_headers_json", json.dumps(list(package_headers), ensure_ascii=False))
                _set_meta(conn, "casp_headers_json", json.dumps(list(casp_headers), ensure_ascii=False))

                for rec in mdb.records:
                    conn.execute(
                        """INSERT INTO cc_master(
                            excel_row, package, package_hash, source_path, status, cc_type,
                            health_status, risk, risk_score, library_profiles, merge_profiles,
                            notes, thumbnail_target, record_json
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (
                            int(rec.get("_ExcelRow", "0") or 0),
                            _first(rec, "Package"),
                            _first(rec, "Package Hash", "SHA256", "SHA-256"),
                            _first(rec, "Source Path", "Exact Path", "Current Path", "Path"),
                            _first(rec, "Status"),
                            _first(rec, "CC Type"),
                            _first(rec, "Health Status", "Health"),
                            _first(rec, "Risk"),
                            _first(rec, "Risk Score", "Performance Risk Score", "Performance Score"),
                            _first(rec, "CC Magic Profiles"),
                            _first(rec, "Merge Profiles"),
                            _first(rec, "Notes"),
                            _first(rec, "_ThumbnailTarget"),
                            _clean_json_dict(rec),
                        ),
                    )

                for row_num, rec in package_rows:
                    conn.execute(
                        "INSERT INTO packages(excel_row, package, package_hash, source_path, record_json) VALUES (?, ?, ?, ?, ?)",
                        (
                            row_num,
                            _first(rec, "Package", "Filename", "File Name"),
                            _first(rec, "Package Hash", "SHA256", "SHA-256"),
                            _first(rec, "Source Path", "Exact Path", "Current Path", "Path"),
                            _clean_json_dict(rec),
                        ),
                    )

                for row_num, rec in casp_rows:
                    conn.execute(
                        "INSERT INTO casp_master(excel_row, package, casp_instance, cas_type, record_json) VALUES (?, ?, ?, ?, ?)",
                        (
                            row_num,
                            _first(rec, "Package"),
                            _first(rec, "CASP Instance", "Instance"),
                            _first(rec, "CAS Type", "CC Type", "Type"),
                            _clean_json_dict(rec),
                        ),
                    )

                for name in mdb.profile_options:
                    clean = str(name or "").strip()
                    if clean:
                        conn.execute("INSERT OR IGNORE INTO library_profiles(name) VALUES (?)", (clean,))
                for name in mdb.merge_profile_options:
                    clean = str(name or "").strip()
                    if clean:
                        conn.execute("INSERT OR IGNORE INTO merge_profiles(name) VALUES (?)", (clean,))
                for name, folder_path in mdb.file_locations:
                    name = str(name or "").strip()
                    folder_path = str(folder_path or "").strip()
                    if name and folder_path:
                        conn.execute(
                            "INSERT OR REPLACE INTO file_locations(name, folder_path) VALUES (?, ?)",
                            (name, folder_path),
                        )
                for profile, package_hash, package_name, location_name in mdb.profile_layouts:
                    conn.execute(
                        "INSERT INTO profile_layouts(profile, package_hash, package_name, file_location) VALUES (?, ?, ?, ?)",
                        (profile, package_hash, package_name, location_name),
                    )
                for key, value in mdb.metadata.items():
                    conn.execute(
                        "INSERT OR REPLACE INTO workbook_settings(key, value) VALUES (?, ?)",
                        (str(key), "" if value is None else str(value)),
                    )

                # If migration produced a truly empty Library (for example an old
                # blank MDB that only contained orphan CASP rows), there are no old
                # package scores to refresh. Mark the empty SQLite Library current.
                if not mdb.records and not package_rows and not casp_rows:
                    _set_meta(conn, PERFORMANCE_SCORE_VERSION_METADATA_KEY, CURRENT_PERFORMANCE_SCORE_VERSION)

            # Validate the just-written database before it can replace a prior test DB.
            integrity = conn.execute("PRAGMA integrity_check").fetchone()
            if not integrity or str(integrity[0]).lower() != "ok":
                raise RuntimeError(f"SQLite integrity check failed: {integrity[0] if integrity else 'no result'}")

            actual_cc = int(conn.execute("SELECT COUNT(*) FROM cc_master").fetchone()[0])
            actual_packages = int(conn.execute("SELECT COUNT(*) FROM packages").fetchone()[0])
            actual_casp = int(conn.execute("SELECT COUNT(*) FROM casp_master").fetchone()[0])
            if actual_cc != len(mdb.records):
                raise RuntimeError(f"CC Master validation failed: expected {len(mdb.records)}, SQLite has {actual_cc}.")
            if actual_packages != len(package_rows):
                raise RuntimeError(f"Packages validation failed: expected {len(package_rows)}, SQLite has {actual_packages}.")
            if actual_casp != len(casp_rows):
                raise RuntimeError(f"CASP Master validation failed: expected {len(casp_rows)}, SQLite has {actual_casp}.")
        finally:
            conn.close()

        source_hash_after = _sha256_file(source)
        if source_hash_after != source_hash_before:
            raise RuntimeError("Migration stopped because the Master Database changed while it was being read.")

        backup_path: Optional[Path] = None
        if target.exists():
            backup_dir = target.parent / "SQLite_Backups"
            backup_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
            backup_path = backup_dir / f"{target.stem}_backup_{stamp}{target.suffix}"
            shutil.copy2(target, backup_path)

        os.replace(str(tmp), str(target))

        return SQLiteMigrationResult(
            db_path=target,
            backup_path=backup_path,
            source_path=source,
            source_sha256=source_hash_before,
            cc_master_rows=len(mdb.records),
            packages_rows=len(package_rows),
            casp_rows=len(casp_rows),
            library_profiles=len(mdb.profile_options),
            merge_profiles=len(mdb.merge_profile_options),
            file_locations=len(mdb.file_locations),
            profile_layouts=len(mdb.profile_layouts),
            elapsed_seconds=time.perf_counter() - started,
            orphan_casp_skipped=orphan_casp_skipped,
        )
    except Exception:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        raise


def validate_sqlite_library(db_path: Union[str, Path]) -> Dict[str, object]:
    path = Path(db_path)
    if not path.is_file():
        raise FileNotFoundError(path)
    conn = sqlite3.connect(str(path))
    try:
        integrity = conn.execute("PRAGMA integrity_check").fetchone()
        result: Dict[str, object] = {
            "integrity": integrity[0] if integrity else "no result",
            "cc_master_rows": int(conn.execute("SELECT COUNT(*) FROM cc_master").fetchone()[0]),
            "packages_rows": int(conn.execute("SELECT COUNT(*) FROM packages").fetchone()[0]),
            "casp_rows": int(conn.execute("SELECT COUNT(*) FROM casp_master").fetchone()[0]),
            "library_profiles": int(conn.execute("SELECT COUNT(*) FROM library_profiles").fetchone()[0]),
            "merge_profiles": int(conn.execute("SELECT COUNT(*) FROM merge_profiles").fetchone()[0]),
            "file_locations": int(conn.execute("SELECT COUNT(*) FROM file_locations").fetchone()[0]),
            "profile_layouts": int(conn.execute("SELECT COUNT(*) FROM profile_layouts").fetchone()[0]),
        }
        meta = dict(conn.execute("SELECT key, value FROM meta"))
        result.update({f"meta_{key}": value for key, value in meta.items()})
        return result
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# TEST_015: live SQLite -> Library overlay
# ---------------------------------------------------------------------------

_BROKEN_RESOURCE_LABELS = [
    "GEOM", "CASP", "IMG", "VPXY", "BONE", "BOND", "Morph", "Build/Buy",
    "PRESET", "TONE", "IMAG", "TXTC", "PTRN", "SIMO", "TXTF", "FAMD",
    "OBJD", "OBJK", "XML", "STBL", "NMAP", "RSLT", "FTPT", "RIG", "JAZZ",
    "S3SA", "LITE", "CLIP", "MATD", "MTST", "MDLR", "CWAL", "CFEN", "CTPT",
    "NGMP", "CRMT", "SPT2", "_SPT", "CRAL", "LAYO", "HINF", "2ARY", "Complate",
    "63A33EA7", "Package Storage", "Fireplace", "THUM", "626F60CE", "ICON",
    "THUM 0580A2B5", "THUM 0580A2B4", "THUM 0580A2B6", "SNAP 0580A2CD",
    "SNAP 0580A2CF", "TWNI", "MANIFEST", "8F78F74B",
]


def _scanner_row_library_record(scanner_row: List[object], live_row: Dict[str, object]) -> Dict[str, str]:
    """Convert the scanner's stable row layout into the Library's named fields.

    TEST_014 stores the complete scanner row in JSON. TEST_015 deliberately reads
    that stored row rather than re-opening the Excel workbook after every scan.
    """
    row = list(scanner_row or [])

    def value(index: int, default: object = "") -> str:
        if index < 0 or index >= len(row):
            return "" if default is None else str(default)
        raw = row[index]
        return "" if raw is None else str(raw)

    rec: Dict[str, str] = {
        "Package": value(0),
        "CASP Instance IDs": value(1),
        "GEOM Instance IDs": value(2),
        "IMG Instance IDs": value(3),
        "VPXY Instance IDs": value(4),
        "CC Type": value(6),
        "Content Type": value(7),
        "Risk": value(8),
        "Risk Score": value(9),
        "HQ Warning": value(10),
        "Poly Warning": value(11),
        "Highest Polys": value(12),
        "Biggest Texture": value(13),
        "Skin Type": value(14),
        "STBL IDs": value(15),
        "BGEO IDs": value(16),
        "FACE IDs": value(17),
        "FBLN IDs": value(18),
        "KEY IDs": value(19),
        "BONE Resource Count": value(20),
        "Package Hash": value(21),
        "Health Status": value(22),
        "Health Issues": value(25),
        "Original Source Path": value(138),
        "EA Resource Status": value(141),
        "Subtype": value(142),
        "Texture Resolutions": value(143),
        "GEOM Total Bytes": value(144),
        "4K Texture Count": value(145),
        "Performance Data Version": value(146),
        "Last Scanned": str(live_row.get("last_scanned", "") or ""),
    }

    # Broken resource pairs are stable scanner columns. CASP starts after the
    # single Health Issues field at index 25, so map that one explicitly.
    rec["Broken GEOM Count"] = value(23)
    rec["Broken GEOM IDs"] = value(24)
    pair_index = 26
    for label in _BROKEN_RESOURCE_LABELS[1:]:
        rec[f"Broken {label} Count"] = value(pair_index)
        rec[f"Broken {label} IDs"] = value(pair_index + 1)
        pair_index += 2

    package_name = rec.get("Package", "")
    lower_name = package_name.casefold()
    if "merged" in lower_name or "merge" in lower_name:
        rec["Package Format"] = "Merged Collection"
    elif any(token in lower_name for token in ("collection", "bundle", "suite")):
        rec["Package Format"] = "Collection"
    else:
        rec["Package Format"] = "Single Item"

    # Keep the live identity available for TEST_015 diagnostics. It is private
    # metadata and therefore never appears as a user-facing Library column.
    rec["_SQLiteIdentity"] = str(live_row.get("identity_key", "") or "")
    rec["_SQLiteLive"] = "1"
    return rec


def _find_thumbnail_for_hash(mdb_path: Path, package_hash: str) -> str:
    hash_text = str(package_hash or "").strip().lower()
    if len(hash_text) != 64:
        return ""
    folder = Path(mdb_path).parent / "Thumbnails"
    if not folder.is_dir():
        return ""
    for ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"):
        candidate = folder / (hash_text + ext)
        if candidate.is_file():
            return str(candidate.resolve())
    return ""


def sqlite_baseline_matches_mdb(db_path: Union[str, Path], mdb_path: Union[str, Path]) -> Tuple[bool, str]:
    """Return whether a SQLite library was built from the exact saved MDB on disk.

    TEST_015 uses this as a scan preflight.  A mismatch blocks SQLite writes/overlay,
    but must never prevent the MDB itself from opening so the user can rebuild safely.
    """
    db = Path(db_path)
    source = Path(mdb_path)
    if not db.is_file():
        return False, f"SQLite database not found: {db}"
    if not source.is_file():
        return False, f"Master Database not found: {source}"

    current_hash = _sha256_file(source).strip().casefold()
    try:
        conn = sqlite3.connect(str(db))
        try:
            tables = {
                str(row[0]) for row in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )
            }
            if "meta" not in tables:
                return False, "SQLite database has no migration metadata."
            row = conn.execute("SELECT value FROM meta WHERE key='source_sha256'").fetchone()
            stored_hash = str(row[0] or "").strip().casefold() if row else ""
            if not stored_hash:
                return False, "SQLite database has no saved Master Database fingerprint."
            if stored_hash != current_hash:
                return False, (
                    "SQLite Library baseline does not match the open Master Database. "
                    "Rebuild / Validate SQLite Library before scanning or applying the SQLite overlay."
                )
            return True, ""
        finally:
            conn.close()
    except sqlite3.Error as exc:
        return False, f"Could not validate SQLite baseline: {exc}"


def overlay_live_sqlite_library(base: MDBData, db_path: Union[str, Path]) -> MDBData:
    """Return a Library snapshot with TEST_014 live scan rows overlaid from SQLite.

    Existing MDB/manual organisation fields are preserved for matching records;
    newly scanned packages appear immediately even though the MDB stayed read-only.
    The source MDB itself is never modified by this function.
    """
    db = Path(db_path)
    if not db.is_file():
        return base

    conn = sqlite3.connect(str(db))
    conn.row_factory = sqlite3.Row
    try:
        tables = {
            str(row[0]) for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            )
        }
        if "scan_packages_live" not in tables:
            return base

        # Fail closed if the SQLite baseline belongs to a different workbook.
        # TEST_014 guarantees the MDB hash is unchanged by scans, so a mismatch
        # means this database should be rebuilt/validated before it is displayed.
        if "meta" in tables:
            row = conn.execute("SELECT value FROM meta WHERE key='source_sha256'").fetchone()
            stored_source_hash = str(row[0] or "").strip().casefold() if row else ""
            if stored_source_hash and stored_source_hash != str(base.source_sha256 or "").strip().casefold():
                raise RuntimeError(
                    "SQLite Library baseline does not match the open Master Database. "
                    "Run Create / Validate SQLite Library again before using TEST_015."
                )

        live_rows = list(conn.execute(
            "SELECT identity_key, package, package_hash, source_path, last_scanned, record_json "
            "FROM scan_packages_live ORDER BY rowid"
        ))
        if not live_rows:
            return base

        live_casp_by_identity: Dict[str, List[Dict[str, str]]] = {}
        if "scan_casp_live" in tables:
            for row in conn.execute(
                "SELECT package_identity_key, record_json FROM scan_casp_live ORDER BY rowid"
            ):
                try:
                    record = json.loads(row["record_json"] or "{}")
                except (TypeError, ValueError):
                    record = {}
                if isinstance(record, dict):
                    live_casp_by_identity.setdefault(str(row["package_identity_key"]), []).append(
                        {str(k): "" if v is None else str(v) for k, v in record.items()}
                    )

        records = [dict(rec) for rec in base.records]
        by_hash: Dict[str, List[int]] = {}
        by_package: Dict[str, List[int]] = {}
        for index, rec in enumerate(records):
            h = str(rec.get("Package Hash", "") or "").strip().casefold()
            if h:
                by_hash.setdefault(h, []).append(index)
            name = str(rec.get("Package", "") or "").strip().casefold()
            if name:
                by_package.setdefault(name, []).append(index)

        manual_fields = (
            "Image", "Status", "Creator", "Converter", "Library Location", "File Location",
            "CC Magic Profiles", "Merge Profiles", "Notes", "Duplicate Files",
        )
        used_base_indexes = set()
        synthetic_row = max(
            [int(str(rec.get("_ExcelRow", "0") or "0")) for rec in records if str(rec.get("_ExcelRow", "0") or "0").isdigit()] or [1]
        ) + 1

        for row in live_rows:
            try:
                payload = json.loads(row["record_json"] or "{}")
            except (TypeError, ValueError):
                payload = {}
            scanner_row = payload.get("_ScannerRow", []) if isinstance(payload, dict) else []
            live_rec = _scanner_row_library_record(scanner_row, dict(row))
            live_hash = str(row["package_hash"] or live_rec.get("Package Hash", "")).strip().casefold()
            package_key = str(row["package"] or live_rec.get("Package", "")).strip().casefold()

            match_index: Optional[int] = None
            if live_hash and len(by_hash.get(live_hash, [])) == 1:
                match_index = by_hash[live_hash][0]
            elif package_key and len(by_package.get(package_key, [])) == 1:
                match_index = by_package[package_key][0]

            if match_index is not None:
                old = records[match_index]
                for field in manual_fields:
                    if old.get(field, "") not in (None, ""):
                        live_rec[field] = str(old.get(field, ""))
                if old.get("_ExcelRow"):
                    live_rec["_ExcelRow"] = str(old.get("_ExcelRow"))
                if old.get("_ThumbnailTarget"):
                    live_rec["_ThumbnailTarget"] = str(old.get("_ThumbnailTarget"))
                    live_rec["Thumbnail"] = "Yes"
                if old.get("_SQLiteRecordKey"):
                    live_rec["_SQLiteRecordKey"] = str(old.get("_SQLiteRecordKey"))
                records[match_index] = live_rec
                used_base_indexes.add(match_index)
            else:
                live_rec["_ExcelRow"] = str(synthetic_row)
                live_rec["_SQLiteRecordKey"] = "live:" + str(row["identity_key"])
                synthetic_row += 1
                records.append(live_rec)

            if not live_rec.get("_ThumbnailTarget"):
                # SQLite-primary scans store thumbnails beside the active .db.
                # Fall back to the legacy workbook path only for older overlay use.
                thumbnail_base = Path(str((getattr(base, "metadata", {}) or {}).get("_SQLitePrimaryDbPath", "") or base.path))
                thumb = _find_thumbnail_for_hash(thumbnail_base, live_rec.get("Package Hash", ""))
                if thumb:
                    live_rec["_ThumbnailTarget"] = thumb
                    live_rec["Thumbnail"] = "Yes"
                else:
                    live_rec["Thumbnail"] = "No"

            observations = live_casp_by_identity.get(str(row["identity_key"]), [])
            if observations:
                age, gender = _aggregate_casp_age_gender(observations)
                outfit = _aggregate_casp_outfit_categories(observations)
                if age:
                    live_rec["Age"] = age
                if gender:
                    live_rec["Gender"] = gender
                if outfit:
                    live_rec["Outfit Category"] = outfit

        casp_by_package = {key: [dict(item) for item in value] for key, value in base.casp_observations_by_package.items()}
        for row in live_rows:
            identity = str(row["identity_key"])
            observations = live_casp_by_identity.get(identity, [])
            package = str(row["package"] or "").strip()
            if package and observations:
                casp_by_package[package.casefold()] = observations

        headers = list(base.headers)
        preferred_extra = [
            "Creator", "Converter", "CC Magic Profiles", "File Location", "Merge Profiles",
            "Age", "Gender", "Outfit Category", "EA Resource Status", "Subtype",
            "Texture Resolutions", "GEOM Total Bytes", "4K Texture Count", "Performance Data Version",
        ]
        for rec in records:
            for field in rec:
                if field.startswith("_") or field == "Thumbnail":
                    continue
                if field not in headers:
                    headers.append(field)
        for field in preferred_extra:
            if field not in headers:
                headers.append(field)

        thumbnail_count = sum(1 for rec in records if str(rec.get("Thumbnail", "")).casefold() == "yes")
        casp_count = sum(len(items) for items in casp_by_package.values())

        overlay_metadata = dict(base.metadata)
        overlay_metadata["_SQLiteLiveOverlayCount"] = str(len(live_rows))
        overlay_metadata["_SQLiteDbPath"] = str(db.resolve())
        return MDBData(
            path=base.path,
            headers=headers,
            records=records,
            casp_observations_by_package=casp_by_package,
            status_options=list(base.status_options),
            library_options=list(base.library_options),
            profile_options=list(base.profile_options),
            file_locations=list(base.file_locations),
            merge_profile_options=list(base.merge_profile_options),
            profile_layouts=list(base.profile_layouts),
            metadata=overlay_metadata,
            thumbnail_count=thumbnail_count,
            casp_count=casp_count,
            package_history_count=max(base.package_history_count, len(records)),
            source_sha256=base.source_sha256,
            source_size=base.source_size,
            source_mtime_ns=base.source_mtime_ns,
        )
    finally:
        conn.close()

# ---------------------------------------------------------------------------
# TEST_016: SQLite-primary Library + manual-field persistence
# ---------------------------------------------------------------------------

@dataclass
class SQLiteSaveResult:
    backup_path: Path
    changed_rows: int
    changed_fields: int
    deleted_rows: int = 0


def _sqlite_record_key(record: Dict[str, object]) -> str:
    saved = str(record.get("_SQLiteRecordKey", "") or "").strip()
    if saved:
        return saved
    package_hash = str(record.get("Package Hash", "") or "").strip().casefold()
    package = str(record.get("Package", "") or "").strip().casefold()
    source_path = str(
        record.get("Original Source Path", "")
        or record.get("Source Path", "")
        or record.get("Exact Path", "")
        or record.get("Current Path", "")
        or record.get("Path", "")
        or ""
    ).strip()
    if source_path:
        source_path = os.path.normcase(os.path.abspath(source_path)).casefold()
    if package_hash:
        return "hash:" + package_hash + "|package:" + package + "|path:" + source_path
    return "package:" + package + "|path:" + source_path


def _ensure_primary_library_tables(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS manual_package_fields (
            record_key TEXT PRIMARY KEY,
            fields_json TEXT NOT NULL,
            updated_utc TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS deleted_library_records (
            record_key TEXT PRIMARY KEY,
            deleted_utc TEXT NOT NULL
        );
        """
    )


def _table_names(conn: sqlite3.Connection) -> set:
    return {str(row[0]) for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}


def _read_json_record(raw: object) -> Dict[str, str]:
    try:
        value = json.loads(str(raw or "{}"))
    except (TypeError, ValueError, json.JSONDecodeError):
        value = {}
    if not isinstance(value, dict):
        return {}
    return {str(k): "" if v is None else str(v) for k, v in value.items()}


def _json_string_list(raw: object) -> List[str]:
    try:
        value = json.loads(str(raw or "[]"))
    except (TypeError, ValueError, json.JSONDecodeError):
        return []
    if not isinstance(value, list):
        return []
    result: List[str] = []
    for item in value:
        text = str(item or "").strip()
        if text and text not in result:
            result.append(text)
    return result


def _recover_legacy_library_headers(source_path: Path, records: List[Dict[str, str]]) -> List[str]:
    """Recover a complete Library column registry for pre-fix TEST_016 databases.

    Older SQLite snapshots did not store the workbook header list, so a field that
    happened to be blank in every row vanished from Choose Visible Columns. If the
    legacy workbook still exists, read it once to recover the exact registry and
    persist that registry into SQLite. If it is unavailable, retain every observed
    field and guarantee the stable manual Library fields remain available.
    """
    if source_path.is_file() and source_path.suffix.lower() == ".xlsx":
        try:
            return list(MDBWorkbook().load(source_path).headers)
        except Exception:
            pass

    headers: List[str] = []
    for rec in records:
        for field in rec:
            if field.startswith("_") or field == "Thumbnail":
                continue
            if field not in headers:
                headers.append(field)
    # These are stable Library-owned fields and must remain selectable even when
    # every current value is blank.
    for field in (
        "Package", "Status", "Creator", "Converter", "File Location",
        "CC Magic Profiles", "Merge Profiles", "Notes", "Age", "Gender",
        "Outfit Category", "EA Resource Status",
    ):
        if field not in headers:
            headers.append(field)
    if "Package" in headers:
        headers = ["Package"] + [field for field in headers if field != "Package"]
    return headers


def load_sqlite_library(db_path: Union[str, Path]) -> MDBData:
    """Load CAS_Curator.db directly as the TEST_016 primary Library.

    The original XLSX is not opened.  Its saved path/hash remain provenance for the
    transitional scanner/routing code, while normal Library reads come from SQLite.
    """
    db = Path(db_path).resolve()
    if not db.is_file():
        raise FileNotFoundError(db)

    conn = sqlite3.connect(str(db))
    conn.row_factory = sqlite3.Row
    try:
        tables = _table_names(conn)
        required = {"meta", "cc_master", "casp_master", "library_profiles", "merge_profiles", "file_locations", "profile_layouts"}
        missing = sorted(required - tables)
        if missing:
            raise RuntimeError("This SQLite file is not a CAS Curator Library database. Missing: " + ", ".join(missing))
        integrity = conn.execute("PRAGMA integrity_check").fetchone()
        if not integrity or str(integrity[0]).casefold() != "ok":
            raise RuntimeError("SQLite integrity_check failed: " + (str(integrity[0]) if integrity else "no result"))
        _ensure_primary_library_tables(conn)
        conn.commit()

        meta = {str(row[0]): str(row[1] or "") for row in conn.execute("SELECT key, value FROM meta")}
        source_text = str(meta.get("source_workbook", "") or "").strip()
        source_path = Path(source_text).resolve() if source_text else db

        records: List[Dict[str, str]] = []
        for row in conn.execute("SELECT excel_row, record_json FROM cc_master ORDER BY excel_row, id"):
            rec = _read_json_record(row["record_json"])
            if not rec.get("Package"):
                continue
            rec["_ExcelRow"] = str(row["excel_row"] or 0)
            rec["_SQLiteRecordKey"] = _sqlite_record_key(rec)
            records.append(rec)

        casp_by_package: Dict[str, List[Dict[str, str]]] = {}
        for row in conn.execute("SELECT package, record_json FROM casp_master ORDER BY excel_row, id"):
            rec = _read_json_record(row["record_json"])
            package = str(row["package"] or rec.get("Package", "") or "").strip()
            if package:
                casp_by_package.setdefault(package.casefold(), []).append(rec)

        profiles = [str(row[0]) for row in conn.execute("SELECT name FROM library_profiles ORDER BY name COLLATE NOCASE")]
        merge_profiles = [str(row[0]) for row in conn.execute("SELECT name FROM merge_profiles ORDER BY name COLLATE NOCASE")]
        file_locations = [(str(row[0]), str(row[1])) for row in conn.execute("SELECT name, folder_path FROM file_locations ORDER BY name COLLATE NOCASE")]
        profile_layouts = [
            (str(row[0]), str(row[1] or ""), str(row[2] or ""), str(row[3] or ""))
            for row in conn.execute("SELECT profile, package_hash, package_name, file_location FROM profile_layouts ORDER BY id")
        ]
        settings = {}
        if "workbook_settings" in tables:
            settings = {str(row[0]): str(row[1] or "") for row in conn.execute("SELECT key, value FROM workbook_settings")}

        headers = _json_string_list(meta.get("cc_master_headers_json", ""))
        if not headers:
            # Compatibility upgrade for SQLite databases created by the first
            # TEST_013/014/015/016 builds. Recover the exact header registry once
            # so blank columns cannot disappear from the UI again.
            headers = _recover_legacy_library_headers(source_path, records)
            with conn:
                _set_meta(conn, "cc_master_headers_json", json.dumps(headers, ensure_ascii=False))
                _set_meta(conn, "header_registry_recovered_utc", datetime.utcnow().replace(microsecond=0).isoformat() + "Z")

        # Live scanner data can introduce a genuinely new field after migration.
        # Keep such fields without sacrificing the preserved blank-column registry.
        for rec in records:
            for field in rec:
                if field.startswith("_") or field == "Thumbnail":
                    continue
                if field not in headers:
                    headers.append(field)
        if "Package" in headers:
            headers = ["Package"] + [field for field in headers if field != "Package"]

        stored_status_options = [
            item.strip() for item in str(settings.get("Status Options", "") or "").replace("|", ";").split(";")
            if item.strip()
        ]
        if not stored_status_options:
            stored_status_options = sorted(
                {str(rec.get("Status", "") or "").strip() for rec in records if str(rec.get("Status", "") or "").strip()},
                key=str.casefold,
            )

        base = MDBData(
            path=source_path,
            headers=headers,
            records=records,
            casp_observations_by_package=casp_by_package,
            status_options=stored_status_options,
            library_options=[],
            profile_options=profiles,
            file_locations=file_locations,
            merge_profile_options=merge_profiles,
            profile_layouts=profile_layouts,
            metadata={**settings, **meta, "_SQLitePrimaryDbPath": str(db), "_SQLitePrimary": "1"},
            thumbnail_count=sum(1 for rec in records if str(rec.get("Thumbnail", "") or "").casefold() == "yes" or bool(rec.get("_ThumbnailTarget"))),
            casp_count=sum(len(items) for items in casp_by_package.values()),
            package_history_count=len(records),
            source_sha256=str(meta.get("source_sha256", "") or ""),
            source_size=int(str(meta.get("source_size", "0") or "0") or 0),
            source_mtime_ns=int(str(meta.get("source_mtime_ns", "0") or "0") or 0),
        )
    finally:
        conn.close()

    # Reuse the already-proven TEST_015 live-row overlay, then layer user edits
    # on top so manual organisation is never replaced by a rescan.
    merged = overlay_live_sqlite_library(base, db)

    conn = sqlite3.connect(str(db))
    try:
        _ensure_primary_library_tables(conn)
        manual = {str(row[0]): _read_json_record(row[1]) for row in conn.execute("SELECT record_key, fields_json FROM manual_package_fields")}
        deleted = {str(row[0]) for row in conn.execute("SELECT record_key FROM deleted_library_records")}
    finally:
        conn.close()

    filtered: List[Dict[str, str]] = []
    for rec in merged.records:
        key = _sqlite_record_key(rec)
        rec["_SQLiteRecordKey"] = key
        if key in deleted:
            continue
        changes = manual.get(key)
        if changes:
            rec.update(changes)
        filtered.append(rec)
    merged.records = filtered
    merged.thumbnail_count = sum(1 for rec in filtered if str(rec.get("Thumbnail", "") or "").casefold() == "yes" or bool(rec.get("_ThumbnailTarget")))
    merged.package_history_count = max(merged.package_history_count, len(filtered))
    metadata = dict(merged.metadata)
    metadata["_SQLitePrimaryDbPath"] = str(db)
    metadata["_SQLitePrimary"] = "1"
    # The overlay count is a raw scanner-table count and can include rows hidden
    # by Library deletion tombstones.  Keep a separate post-filter count so the
    # GUI/log reports what the user can actually see.
    metadata["_SQLiteVisibleRecordCount"] = str(len(filtered))
    merged.metadata = metadata
    return merged


def save_sqlite_manual_changes(
    data: MDBData,
    db_path: Union[str, Path],
    pending_changes: Dict[int, Dict[str, str]],
    profile_options: Optional[List[str]] = None,
    status_options: Optional[List[str]] = None,
    delete_rows: Optional[Iterable[int]] = None,
    file_locations: Optional[List[Tuple[str, str]]] = None,
    merge_profile_options: Optional[List[str]] = None,
    profile_layouts: Optional[List[Tuple[str, str, str, str]]] = None,
) -> SQLiteSaveResult:
    db = Path(db_path).resolve()
    if not db.is_file():
        raise FileNotFoundError(db)
    backup_dir = db.parent / "SQLite_Backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
    backup = backup_dir / f"{db.stem}_pre_manual_save_{stamp}{db.suffix}"
    shutil.copy2(str(db), str(backup))

    by_row: Dict[int, Dict[str, str]] = {}
    for rec in data.records:
        try:
            row_num = int(str(rec.get("_ExcelRow", "0") or "0"))
        except ValueError:
            continue
        if row_num:
            by_row[row_num] = rec

    deleted = {int(value) for value in (delete_rows or [])}
    changed_rows = 0
    changed_fields = 0
    now = datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

    conn = sqlite3.connect(str(db), timeout=30.0)
    conn.execute("PRAGMA busy_timeout=30000")
    try:
        _ensure_primary_library_tables(conn)
        with conn:
            for row_num, fields in (pending_changes or {}).items():
                rec = by_row.get(int(row_num))
                if not rec or int(row_num) in deleted:
                    continue
                key = _sqlite_record_key(rec)
                prior = conn.execute("SELECT fields_json FROM manual_package_fields WHERE record_key=?", (key,)).fetchone()
                payload = _read_json_record(prior[0]) if prior else {}
                payload.update({str(k): "" if v is None else str(v) for k, v in fields.items()})
                conn.execute(
                    "INSERT OR REPLACE INTO manual_package_fields(record_key, fields_json, updated_utc) VALUES (?, ?, ?)",
                    (key, json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")), now),
                )
                changed_rows += 1
                changed_fields += len(fields)

            for row_num in deleted:
                rec = by_row.get(int(row_num))
                if not rec:
                    continue
                key = _sqlite_record_key(rec)
                conn.execute("INSERT OR REPLACE INTO deleted_library_records(record_key, deleted_utc) VALUES (?, ?)", (key, now))
                conn.execute("DELETE FROM manual_package_fields WHERE record_key=?", (key,))

            if profile_options is not None:
                conn.execute("DELETE FROM library_profiles")
                conn.executemany("INSERT OR IGNORE INTO library_profiles(name) VALUES (?)", [(str(v).strip(),) for v in profile_options if str(v).strip()])
            if merge_profile_options is not None:
                conn.execute("DELETE FROM merge_profiles")
                conn.executemany("INSERT OR IGNORE INTO merge_profiles(name) VALUES (?)", [(str(v).strip(),) for v in merge_profile_options if str(v).strip()])
            if file_locations is not None:
                conn.execute("DELETE FROM file_locations")
                conn.executemany(
                    "INSERT OR REPLACE INTO file_locations(name, folder_path) VALUES (?, ?)",
                    [(str(name).strip(), str(path).strip()) for name, path in file_locations if str(name).strip() and str(path).strip()],
                )
            if profile_layouts is not None:
                conn.execute("DELETE FROM profile_layouts")
                conn.executemany(
                    "INSERT INTO profile_layouts(profile, package_hash, package_name, file_location) VALUES (?, ?, ?, ?)",
                    [(str(a), str(b), str(c), str(d)) for a, b, c, d in profile_layouts],
                )
            if status_options is not None:
                conn.execute("INSERT OR REPLACE INTO workbook_settings(key, value) VALUES (?, ?)", ("Status Options", "; ".join(str(v) for v in status_options)))
            _set_meta(conn, "storage_mode", "sqlite_primary_test_016")
            _set_meta(conn, "last_manual_save_utc", now)

        integrity = conn.execute("PRAGMA integrity_check").fetchone()
        if not integrity or str(integrity[0]).casefold() != "ok":
            raise RuntimeError("SQLite integrity_check failed after manual save: " + (str(integrity[0]) if integrity else "no result"))
    except Exception:
        conn.close()
        shutil.copy2(str(backup), str(db))
        raise
    else:
        conn.close()

    return SQLiteSaveResult(
        backup_path=backup,
        changed_rows=changed_rows,
        changed_fields=changed_fields,
        deleted_rows=len(deleted),
    )