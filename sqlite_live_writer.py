"""R27 SQLite-primary live scan writer for CAS Curator.

Scanner package rows and CASP observations are committed together. A database
backup is created before the scan write. Batch completion is recorded in the
same transaction, so interrupted scans cannot mark files as processed.
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import sqlite3
import time
from datetime import datetime, timezone
from pathlib import Path

PERFORMANCE_SCORE_VERSION_METADATA_KEY = "Performance Score Version"
CURRENT_PERFORMANCE_SCORE_VERSION = "2"


def _now():
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _sha256_file(path):
    digest = hashlib.sha256()
    with open(str(path), "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _value(row, index, default=""):
    try:
        item = row[index]
    except (IndexError, TypeError):
        return default
    return default if item is None else item


def _text(value):
    return "" if value is None else str(value)


def _identity_key(row):
    digest = _text(_value(row, 21)).strip().lower()
    if digest:
        return "hash:" + digest
    source = _text(_value(row, 138)).strip()
    if source:
        return "path:" + os.path.normcase(os.path.abspath(source)).lower()
    return "name:" + _text(_value(row, 0)).strip().lower() + "|casp:" + _text(_value(row, 1)).strip().lower()


def _row_record(row):
    indexes = {
        "Package": 0, "CASP Instance IDs": 1, "GEOM Instance IDs": 2,
        "IMG Instance IDs": 3, "VPXY Instance IDs": 4, "CC Type": 6,
        "Content Type": 7, "Risk": 8, "Risk Score": 9, "HQ Warning": 10,
        "Poly Warning": 11, "Highest Polys": 12, "Biggest Texture": 13,
        "Skin Type": 14, "Package Hash": 21, "Health Status": 22,
        "Health Issues": 25, "Original Source Path": 138,
        "Texture Resolutions": 143, "GEOM Total Bytes": 144,
        "4K Texture Count": 145, "Performance Data Version": 146,
    }
    record = {name: _value(row, index) for name, index in indexes.items()}
    for key in record:
        if key not in ("Highest Polys", "GEOM Total Bytes", "4K Texture Count"):
            record[key] = _text(record[key])
    record["_ScannerRow"] = list(row)
    return record


def _ensure_tables(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS scan_runs(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        started_utc TEXT NOT NULL, finished_utc TEXT,
        source_mdb TEXT NOT NULL, source_mdb_sha256 TEXT NOT NULL,
        package_count INTEGER NOT NULL DEFAULT 0,
        casp_count INTEGER NOT NULL DEFAULT 0,
        status TEXT NOT NULL, elapsed_seconds REAL NOT NULL DEFAULT 0
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS scan_packages_live(
        identity_key TEXT PRIMARY KEY,
        package TEXT NOT NULL, package_hash TEXT, source_path TEXT,
        casp_ids TEXT, cc_type TEXT, content_type TEXT, health_status TEXT,
        risk TEXT, risk_score TEXT, highest_polys TEXT, biggest_texture TEXT,
        texture_resolutions TEXT, geom_total_bytes TEXT, texture_4k_count TEXT,
        performance_data_version TEXT, last_scanned TEXT NOT NULL,
        last_scan_run_id INTEGER, record_json TEXT NOT NULL,
        FOREIGN KEY(last_scan_run_id) REFERENCES scan_runs(id)
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS scan_casp_live(
        package_identity_key TEXT NOT NULL, casp_instance TEXT NOT NULL,
        package TEXT, cc_type TEXT, data_type TEXT, age_gender TEXT,
        category TEXT, body_type TEXT, overlay_priority TEXT,
        last_scan_run_id INTEGER, record_json TEXT NOT NULL,
        PRIMARY KEY(package_identity_key, casp_instance),
        FOREIGN KEY(package_identity_key) REFERENCES scan_packages_live(identity_key)
            ON DELETE CASCADE,
        FOREIGN KEY(last_scan_run_id) REFERENCES scan_runs(id)
    )""")
    for sql in (
        "CREATE INDEX IF NOT EXISTS idx_scan_packages_live_package ON scan_packages_live(package COLLATE NOCASE)",
        "CREATE INDEX IF NOT EXISTS idx_scan_packages_live_hash ON scan_packages_live(package_hash)",
        "CREATE INDEX IF NOT EXISTS idx_scan_packages_live_source_path ON scan_packages_live(source_path COLLATE NOCASE)",
        "CREATE INDEX IF NOT EXISTS idx_scan_packages_live_cc_type ON scan_packages_live(cc_type COLLATE NOCASE)",
        "CREATE INDEX IF NOT EXISTS idx_scan_casp_live_instance ON scan_casp_live(casp_instance COLLATE NOCASE)",
        "CREATE INDEX IF NOT EXISTS idx_scan_casp_live_package ON scan_casp_live(package COLLATE NOCASE)",
    ):
        conn.execute(sql)


def _set_meta(conn, key, value):
    conn.execute("INSERT OR REPLACE INTO meta(key,value) VALUES (?,?)",
                 (str(key), _text(value)))


def _restore_deleted(conn, package_hash):
    digest = _text(package_hash).strip().lower()
    if len(digest) != 64 or any(ch not in "0123456789abcdef" for ch in digest):
        return 0
    if not conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='deleted_library_records'"
    ).fetchone():
        return 0
    ordinary = "hash:" + digest + "|"
    live = "live:hash:" + digest
    cursor = conn.execute(
        "DELETE FROM deleted_library_records "
        "WHERE substr(lower(record_key),1,?)=? OR substr(lower(record_key),1,?)=?",
        (len(ordinary), ordinary, len(live), live),
    )
    return max(0, int(cursor.rowcount or 0))


def _backup_database(path):
    root = path.parent / "SQLite_Backups"
    root.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    backup = root / (path.stem + "_pre_scan_" + stamp + path.suffix)
    # SQLite's backup API handles WAL-backed databases correctly, unlike
    # copying the .db file while an active WAL contains uncheckpointed pages.
    with sqlite3.connect(str(path), timeout=30) as original:
        with sqlite3.connect(str(backup), timeout=30) as destination:
            original.backup(destination)
    return backup


def _restore_backup(db_path, backup_path):
    # Restoring a backup via SQLite's backup API preserves transactional
    # correctness even if an active WAL mode is in use.
    with sqlite3.connect(str(backup_path), timeout=30) as original:
        with sqlite3.connect(str(db_path), timeout=30) as destination:
            original.backup(destination)


def write_scan_to_sqlite(rows, casp_records, db_path, source_mdb=None, batch_signatures=None):
    """Write one successful scan transaction to an existing SQLite Library.

    The caller must supply completed scanner rows, never partially scanned
    records. This function never modifies the source .package/.sim files.
    """
    started = time.perf_counter()
    db_path = Path(db_path).resolve()
    if not db_path.is_file():
        raise RuntimeError("SQLite primary Library not found: " + str(db_path))
    rows = list(rows)
    casp_records = list(casp_records)
    source_mdb_text = ""
    source_mdb_hash = ""
    if source_mdb:
        candidate = Path(source_mdb).resolve()
        if candidate.is_file():
            source_mdb_text = str(candidate)
            source_mdb_hash = _sha256_file(candidate)

    # Validate all mappings before taking any write lock.
    identities = {}
    by_name = {}
    for row in rows:
        record = _row_record(row)
        key = _identity_key(row)
        if not record["Package"]:
            raise ValueError("Scanner row has no package name")
        if key in identities and identities[key]["Original Source Path"] != record["Original Source Path"]:
            raise RuntimeError("Ambiguous identity for different physical packages: " + key)
        identities[key] = record
        name = record["Package"].casefold()
        if name in by_name and by_name[name] != key:
            by_name[name] = None
        elif name not in by_name:
            by_name[name] = key
    for casp in casp_records:
        if not _text(casp.get("CASP Instance", "")).strip():
            continue
        name = _text(casp.get("Package", "")).strip().casefold()
        if not by_name.get(name):
            raise RuntimeError("Ambiguous or missing CASP-to-package mapping: " + name)

    backup_path = _backup_database(db_path)
    now = _now()
    conn = sqlite3.connect(str(db_path), timeout=30)
    conn.execute("PRAGMA busy_timeout=30000")
    conn.execute("PRAGMA foreign_keys=ON")
    restored_deleted = 0
    touched = 0
    casp_touched = 0
    run_id = None
    try:
        with conn:
            _ensure_tables(conn)
            _set_meta(conn, "storage_mode", "sqlite_primary_test_018")
            _set_meta(conn, "last_live_scan_utc", now)
            migrated = conn.execute("SELECT value FROM meta WHERE key='source_workbook'").fetchone()
            if not batch_signatures or not (migrated and _text(migrated[0]).strip()):
                _set_meta(conn, PERFORMANCE_SCORE_VERSION_METADATA_KEY, CURRENT_PERFORMANCE_SCORE_VERSION)
            _set_meta(conn, "last_live_score_model_version", CURRENT_PERFORMANCE_SCORE_VERSION)

            cur = conn.execute(
                "INSERT INTO scan_runs(started_utc,source_mdb,source_mdb_sha256,"
                "package_count,casp_count,status,elapsed_seconds)"
                " VALUES(?,?,?,?,?,?,?)",
                (now, source_mdb_text, source_mdb_hash,
                 len(rows), len(casp_records), "writing", 0.0),
            )
            run_id = int(cur.lastrowid)
            for identity, rec in identities.items():
                restored_deleted += _restore_deleted(conn, rec["Package Hash"])
                conn.execute("""INSERT INTO scan_packages_live(
                    identity_key,package,package_hash,source_path,casp_ids,cc_type,
                    content_type,health_status,risk,risk_score,highest_polys,
                    biggest_texture,texture_resolutions,geom_total_bytes,
                    texture_4k_count,performance_data_version,last_scanned,
                    last_scan_run_id,record_json)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(identity_key) DO UPDATE SET
                    package=excluded.package,package_hash=excluded.package_hash,
                    source_path=excluded.source_path,casp_ids=excluded.casp_ids,
                    cc_type=excluded.cc_type,content_type=excluded.content_type,
                    health_status=excluded.health_status,risk=excluded.risk,
                    risk_score=excluded.risk_score,highest_polys=excluded.highest_polys,
                    biggest_texture=excluded.biggest_texture,
                    texture_resolutions=excluded.texture_resolutions,
                    geom_total_bytes=excluded.geom_total_bytes,
                    texture_4k_count=excluded.texture_4k_count,
                    performance_data_version=excluded.performance_data_version,
                    last_scanned=excluded.last_scanned,
                    last_scan_run_id=excluded.last_scan_run_id,
                    record_json=excluded.record_json""", (
                        identity, rec["Package"], rec["Package Hash"],
                        rec["Original Source Path"], rec["CASP Instance IDs"],
                        rec["CC Type"], rec["Content Type"], rec["Health Status"],
                        rec["Risk"], rec["Risk Score"], _text(rec["Highest Polys"]),
                        rec["Biggest Texture"], rec["Texture Resolutions"],
                        _text(rec["GEOM Total Bytes"]), _text(rec["4K Texture Count"]),
                        rec["Performance Data Version"], now, run_id,
                        json.dumps(rec, ensure_ascii=False, separators=(",", ":")),
                    ))
                conn.execute("DELETE FROM scan_casp_live WHERE package_identity_key=?",
                             (identity,))
                touched += 1

            for casp in casp_records:
                instance = _text(casp.get("CASP Instance", "")).strip()
                if not instance:
                    continue
                name = _text(casp.get("Package", "")).strip().casefold()
                identity = by_name[name]
                conn.execute("""INSERT OR REPLACE INTO scan_casp_live(
                    package_identity_key,casp_instance,package,cc_type,data_type,
                    age_gender,category,body_type,overlay_priority,last_scan_run_id,
                    record_json) VALUES (?,?,?,?,?,?,?,?,?,?,?)""", (
                        identity, instance, _text(casp.get("Package")),
                        _text(casp.get("CC Type")), _text(casp.get("Data Type")),
                        _text(casp.get("Age Gender")), _text(casp.get("Category")),
                        _text(casp.get("Body Type")),
                        _text(casp.get("Overlay Priority")), run_id,
                        json.dumps(casp, ensure_ascii=False, separators=(",", ":")),
                    ))
                casp_touched += 1

            if batch_signatures:
                from scan_batch_progress import commit_batch_progress, commit_routing_batch
                commit_batch_progress(conn, batch_signatures)
                commit_routing_batch(conn, batch_signatures)

            conn.execute("UPDATE scan_runs SET finished_utc=?,status=?,elapsed_seconds=? WHERE id=?",
                         (_now(), "ok", time.perf_counter() - started, run_id))

        result = conn.execute("PRAGMA integrity_check").fetchone()
        if not result or str(result[0]).lower() != "ok":
            raise RuntimeError("SQLite integrity check failed")
        expected = len(identities)
        actual = conn.execute(
            "SELECT count(*) FROM scan_packages_live WHERE last_scan_run_id=?",
            (run_id,),
        ).fetchone()[0]
        if actual != expected:
            raise RuntimeError("SQLite written package count differs: %d vs %d" % (actual, expected))
        casp_actual = conn.execute(
            "SELECT count(*) FROM scan_casp_live WHERE last_scan_run_id=?",
            (run_id,),
        ).fetchone()[0]
        if casp_actual > len(casp_records):
            raise RuntimeError("SQLite CASP count exceeds scanner results")
    except Exception:
        conn.close()
        _restore_backup(db_path, backup_path)
        raise
    else:
        conn.close()
    return {
        "db_path": str(db_path),
        "backup_path": str(backup_path),
        "packages": touched,
        "batch_completed": len(batch_signatures or ()),
        "casp": casp_touched,
        "restored_deleted_records": restored_deleted,
        "elapsed_seconds": time.perf_counter() - started,
    }
