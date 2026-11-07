"""SQLite batch resume and routing scope regression checks."""
import sqlite3
import tempfile
import unittest
from pathlib import Path
from scan_batch_progress import (
    _path_key, _file_signature, select_next_batch, commit_batch_progress,
    commit_routing_batch, load_routing_batch, reset_source_progress,
)


class BoundedBatchTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.source = self.root / "source"
        self.source.mkdir()
        self.db = self.root / "CAS_Curator.db"
        sqlite3.connect(self.db).close()
        self.files = [self.source / f"{i:03d}.package" for i in range(8)]
        for i, path in enumerate(self.files):
            path.write_bytes(bytes([i]) * 8)

    def test_resume_after_successful_commit(self):
        first = select_next_batch(self.files, self.db, self.source, 3)
        self.assertEqual(len(first["selected"]), 3)
        with sqlite3.connect(self.db) as conn:
            commit_batch_progress(conn, first["signatures"])
        again = select_next_batch(self.files, self.db, self.source, 3)
        self.assertEqual(again["selected"], self.files[3:6])
        self.assertEqual(again["unchanged"], 3)
        self.assertEqual(reset_source_progress(self.db, self.source), 3)
        self.assertEqual(select_next_batch(self.files, self.db, self.source, 3)["selected"], self.files[:3])

    def test_saved_routing_scope_and_change_detection(self):
        signatures = [(_path_key(self.source), _path_key(p), *_file_signature(p)) for p in self.files[:3]]
        with sqlite3.connect(self.db) as conn:
            commit_routing_batch(conn, signatures)
        selected, total = load_routing_batch(self.db, self.source)
        self.assertEqual((len(selected), total), (3, 3))
        self.assertNotIn(_path_key(self.files[4]), selected)
        self.files[0].write_bytes(b"edited-content")
        with self.assertRaisesRegex(RuntimeError, "changed"):
            load_routing_batch(self.db, self.source)

    def test_rejects_bad_bound_and_missing_batch(self):
        with self.assertRaises(ValueError):
            select_next_batch(self.files, self.db, self.source, 0)
        with self.assertRaisesRegex(RuntimeError, "No saved batch"):
            load_routing_batch(self.db, self.source)


if __name__ == "__main__":
    unittest.main()
