"""R27 SQLite-primary live scanner writer: transactional regression coverage."""
import hashlib
import sqlite3
import tempfile
import unittest
from pathlib import Path

from scan_batch_progress import (
    _file_signature, _path_key, select_next_batch,
)
from sqlite_live_writer import write_scan_to_sqlite


def row_for(path):
    row = [""] * 147
    row[0] = path.name
    row[1] = "1234567890ABCDEF"
    row[6] = "Top"
    row[7] = "Custom Content"
    row[21] = hashlib.sha256(path.read_bytes()).hexdigest()
    row[22] = "OK"
    row[138] = str(path)
    row[146] = "2"
    return row


class SQLiteWriterTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.source = self.root / "source"
        self.source.mkdir()
        self.files = [self.source / ("CC_%d.package" % n) for n in range(3)]
        for i, p in enumerate(self.files):
            p.write_bytes(("testpackage-%d" % i).encode("ascii"))
        self.original = {p: p.read_bytes() for p in self.files}
        self.db = self.root / "CAS_Curator.db"
        with sqlite3.connect(self.db) as conn:
            conn.execute("CREATE TABLE meta(key TEXT PRIMARY KEY,value TEXT)")

    def _signatures(self, paths):
        return [(_path_key(self.source), _path_key(p), *_file_signature(p))
                for p in paths]

    def test_scanner_package_casp_and_progress_commit_together(self):
        result = write_scan_to_sqlite(
            [row_for(p) for p in self.files[:2]],
            [{"Package": self.files[0].name,
              "CASP Instance": "1234567890ABCDEF", "Age Gender": "Adult Female"}],
            self.db, batch_signatures=self._signatures(self.files[:2]),
        )
        self.assertEqual(result["packages"], 2)
        self.assertEqual(result["casp"], 1)
        self.assertEqual(result["batch_completed"], 2)
        self.assertTrue(Path(result["backup_path"]).is_file())
        with sqlite3.connect(self.db) as conn:
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM scan_packages_live").fetchone()[0], 2)
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM scan_casp_live").fetchone()[0], 1)
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM scan_batch_progress").fetchone()[0], 2)
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM scan_routing_batch").fetchone()[0], 2)
            self.assertEqual(conn.execute("PRAGMA integrity_check").fetchone()[0], "ok")
            self.assertEqual(conn.execute(
                "SELECT value FROM meta WHERE key='Performance Score Version'").fetchone()[0], "2")
        self.assertEqual({p: p.read_bytes() for p in self.files}, self.original)

    def test_migrated_partial_scan_does_not_claim_global_score_refresh(self):
        with sqlite3.connect(self.db) as conn:
            conn.execute("INSERT INTO meta VALUES (?,?)", ("source_workbook", "old.xlsx"))
            conn.execute("INSERT INTO meta VALUES (?,?)", ("Performance Score Version", "legacy"))
        write_scan_to_sqlite([row_for(self.files[0])], [], self.db,
                             batch_signatures=self._signatures(self.files[:1]))
        with sqlite3.connect(self.db) as conn:
            self.assertEqual(conn.execute(
                "SELECT value FROM meta WHERE key='Performance Score Version'").fetchone()[0], "legacy")
            self.assertEqual(conn.execute(
                "SELECT value FROM meta WHERE key='last_live_score_model_version'").fetchone()[0], "2")

    def test_failed_casp_mapping_leaves_database_unchanged(self):
        invalid = [{"Package": "Not scanned.package", "CASP Instance": "ABCDEF"}]
        with self.assertRaisesRegex(RuntimeError, "mapping"):
            write_scan_to_sqlite([row_for(self.files[0])], invalid, self.db,
                                 batch_signatures=self._signatures(self.files[:1]))
        with sqlite3.connect(self.db) as conn:
            self.assertFalse(conn.execute(
                "SELECT 1 FROM sqlite_master WHERE name='scan_batch_progress'").fetchone())
        self.assertEqual({p: p.read_bytes() for p in self.files}, self.original)

    def test_database_write_error_restores_backup_and_prevents_progress(self):
        # Force failure AFTER writer starts a transaction; not merely input validation.
        with sqlite3.connect(self.db) as conn:
            conn.execute("CREATE TABLE scan_runs(id INTEGER PRIMARY KEY, started_utc TEXT)")
        with self.assertRaises(sqlite3.OperationalError):
            write_scan_to_sqlite([row_for(self.files[0])], [], self.db,
                                 batch_signatures=self._signatures(self.files[:1]))
        with sqlite3.connect(self.db) as conn:
            self.assertEqual(conn.execute(
                "SELECT COUNT(*) FROM scan_runs").fetchone()[0], 0)
            self.assertFalse(conn.execute(
                "SELECT 1 FROM sqlite_master WHERE name='scan_batch_progress'").fetchone())
            self.assertEqual(conn.execute("PRAGMA integrity_check").fetchone()[0], "ok")
        self.assertEqual({p: p.read_bytes() for p in self.files}, self.original)

    def test_wal_mode_backup_contains_committed_changes(self):
        # SQLite WAL may hold pages outside the main .db file. A pre-scan
        # backup must include all committed state, not just .db bytes.
        with sqlite3.connect(self.db) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("INSERT INTO meta VALUES (?,?)", ("example_setting", "keep me"))
        result = write_scan_to_sqlite([row_for(self.files[0])], [], self.db)
        with sqlite3.connect(result["backup_path"]) as backup:
            self.assertEqual(backup.execute(
                "SELECT value FROM meta WHERE key='example_setting'").fetchone()[0], "keep me")
            self.assertFalse(backup.execute(
                "SELECT 1 FROM sqlite_master WHERE name='scan_packages_live'").fetchone())
        with sqlite3.connect(self.db) as conn:
            self.assertEqual(conn.execute(
                "SELECT COUNT(*) FROM scan_packages_live").fetchone()[0], 1)

    def test_completed_scan_resumes_remaining_files(self):
        selected = select_next_batch(self.files, self.db, self.source, 2)
        self.assertEqual(selected["selected"], self.files[:2])
        write_scan_to_sqlite(
            [row_for(p) for p in selected["selected"]], [], self.db,
            batch_signatures=selected["signatures"],
        )
        second = select_next_batch(self.files, self.db, self.source, 2)
        self.assertEqual(second["selected"], self.files[2:3])
        self.assertEqual(second["unchanged"], 2)


if __name__ == "__main__":
    unittest.main()
