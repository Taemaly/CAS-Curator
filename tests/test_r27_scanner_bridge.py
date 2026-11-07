"""Test opt-in public SQLite scanner dispatch without scanning or moving CC.

A mock scanner module prints environment variables but touches no packages.
"""
from __future__ import annotations

import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class PublicScannerDispatchTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.src = self.root / "source"
        self.src.mkdir()
        self.src_file = self.src / "untouched.package"
        self.src_file.write_bytes(b"original CC bytes unchanged")
        self.dest = self.root / "sorted"
        self.dest.mkdir()
        self.db = self.root / "CAS_Curator.db"
        with sqlite3.connect(self.db) as conn:
            conn.execute("CREATE TABLE meta(key TEXT PRIMARY KEY, value TEXT)")
        self.xlsx = self.root / "legacy.xlsx"
        self.xlsx.write_bytes(b"stub-not-used-by-mock-scanner")
        for name in ("scanner_integration_runner.py", "scanner_sqlite_bridge.py"):
            shutil.copyfile(ROOT / name, self.root / name)
        core = self.root / "scanner_core_reference"
        core.mkdir()
        (core / "Sims3_CC_Health_Scanner_R27_SQLITE.py").write_text(
            "# sqlite_live_writer - mock for regression tests only\n"
            "import os\n"
            "print('MOCK_CORE: storage=%s move=%s batch=%s' % (\n"
            "    os.environ.get('S3CC_SQLITE_LIVE_WRITE','not-set'),\n"
            "    os.environ.get('S3CC_MOVE_FILES','not-set'),\n"
            "    os.environ.get('S3CC_SCAN_BATCH_SIZE','not-set')))\n",
            encoding="utf-8",
        )
        # Legacy XLSX runner still imports its published core, not R27.
        # Install a separate mock so both launch paths can be tested.
        shutil.copyfile(
            core / "Sims3_CC_Health_Scanner_R27_SQLITE.py",
            core / "Sims3_CC_Health_Scanner_CORE_INTEGRATED_TEST.py",
        )

    def run_runner(self, *args):
        result = subprocess.run(
            [sys.executable, str(self.root / "scanner_integration_runner.py"),
             "--source", str(self.src), *args],
            cwd=str(self.root), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, timeout=15,
        )
        self.assertEqual(self.src_file.read_bytes(), b"original CC bytes unchanged")
        self.assertEqual(list(self.dest.iterdir()), [])
        return result

    def test_sqlite_individual_dispatches_to_r27_without_moves(self):
        out = self.run_runner("--sqlite-db", str(self.db), "--mode", "INDIVIDUAL",
                              "--edition", "PUBLIC")
        self.assertEqual(out.returncode, 0, out.stdout)
        self.assertIn("MOCK_CORE: storage=1 move=0 batch=0", out.stdout)

    def test_big_review_requires_bounded_batch(self):
        out = self.run_runner("--sqlite-db", str(self.db),
                              "--mode", "BIG_REVIEW_BATCH", "--edition", "PUBLIC")
        self.assertNotEqual(out.returncode, 0)
        self.assertIn("requires a bounded batch", out.stdout)

    def test_bounded_big_review_has_explicit_limit(self):
        out = self.run_runner("--sqlite-db", str(self.db), "--mode",
                              "BIG_REVIEW_BATCH", "--batch-size", "500",
                              "--edition", "PUBLIC")
        self.assertEqual(out.returncode, 0, out.stdout)
        self.assertIn("MOCK_CORE: storage=1 move=0 batch=500", out.stdout)

    def test_missing_database_fails_before_scanner_import(self):
        out = self.run_runner("--sqlite-db", str(self.root / "missing.db"))
        self.assertNotEqual(out.returncode, 0)
        self.assertNotIn("MOCK_CORE:", out.stdout)

    def test_invalid_sqlite_database_fails_before_import(self):
        invalid = self.root / "invalid.db"
        invalid.write_bytes(b"not sqlite")
        out = self.run_runner("--sqlite-db", str(invalid))
        self.assertNotEqual(out.returncode, 0)
        self.assertIn("SQLite Library preflight failed", out.stdout)

    def test_unapproved_routing_is_blocked(self):
        out = self.run_runner("--sqlite-db", str(self.db), "--mode", "BIG_REVIEW_BATCH",
                              "--edition", "PUBLIC", "--real-routing",
                              "--route-current-batch", "--routing-root", str(self.dest))
        self.assertNotEqual(out.returncode, 0)
        self.assertIn("approved cached Preview missing", out.stdout)
        self.assertNotIn("MOCK_CORE:", out.stdout)

    def test_unbounded_routing_is_blocked(self):
        out = self.run_runner("--sqlite-db", str(self.db), "--mode", "BIG_REVIEW_BATCH",
                              "--edition", "PUBLIC", "--preview-routing",
                              "--routing-root", str(self.dest))
        self.assertNotEqual(out.returncode, 0)
        self.assertIn("last completed bounded Big Review batch", out.stdout)

    def test_sqlite_thumbnail_backfill_is_blocked(self):
        out = self.run_runner("--sqlite-db", str(self.db), "--mode", "THUMBNAIL_BACKFILL",
                              "--edition", "PUBLIC")
        self.assertNotEqual(out.returncode, 0)
        self.assertIn("not yet supported", out.stdout)

    def test_missing_upgraded_core_refuses_legacy_fallback(self):
        core = self.root / "scanner_core_reference/Sims3_CC_Health_Scanner_R27_SQLITE.py"
        core.write_text("print('SHOULD NOT RUN')\n", encoding="utf-8")
        out = self.run_runner("--sqlite-db", str(self.db), "--edition", "PUBLIC")
        self.assertNotEqual(out.returncode, 0)
        self.assertIn("R27 SQLite scanner core not installed", out.stdout)
        self.assertNotIn("SHOULD NOT RUN", out.stdout)

    def test_legacy_mdb_scanner_call_remains_unmodified(self):
        out = self.run_runner("--mdb", str(self.xlsx), "--mode", "INDIVIDUAL",
                              "--edition", "PUBLIC")
        self.assertEqual(out.returncode, 0, out.stdout)
        self.assertIn("MOCK_CORE: storage=not-set move=0 batch=not-set", out.stdout)


if __name__ == "__main__":
    unittest.main()
