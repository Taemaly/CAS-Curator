"""Safety regression tests for the R27 cached SQLite routing backend.

All packages are synthetic bytes in temporary directories. No Sims 3 installation
or existing user CC files are touched.
"""
from __future__ import annotations

import hashlib
import json
import os
import sqlite3
import tempfile
import unittest
from pathlib import Path

from scan_batch_progress import _path_key, _file_signature, commit_routing_batch
from cached_routing_preview import preview
from cached_routing_move import execute
from routing_preview_parser import parse_preview_rows


class CachedRoutingChecks(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.src = self.root / "source"
        self.dst = self.root / "sorted"
        self.runs = self.root / "runs"
        for folder in (self.src, self.dst, self.runs):
            folder.mkdir()
        self.db = self.root / "library.sqlite"
        self.files = [self.src / ("item%d.package" % i) for i in range(3)]
        with sqlite3.connect(self.db) as conn:
            conn.execute("CREATE TABLE scan_packages_live (source_path TEXT, record_json TEXT, package_hash TEXT)")
            sig = []
            for i, path in enumerate(self.files):
                path.write_bytes(("synthetic%d" % i).encode("ascii"))
                sha = hashlib.sha256(path.read_bytes()).hexdigest()
                conn.execute("INSERT INTO scan_packages_live VALUES (?,?,?)",
                             (str(path), json.dumps({"Content Type": "EA Override"}), sha))
                sig.append((_path_key(self.src), _path_key(path), *_file_signature(path)))
            commit_routing_batch(conn, sig)
        self.preview = preview(self.db, self.src, self.dst, self.runs)

    def test_exact_batch_moves_and_ignores_unrelated_files(self):
        extra = self.src / "unrelated.package"
        extra.write_bytes(b"do not move")
        manifest = execute(self.db, self.src, self.dst, self.preview, self.runs)
        self.assertTrue(extra.exists())
        self.assertEqual(len(list(self.dst.rglob("*.package"))), 3)
        self.assertFalse(any(p.exists() for p in self.files))
        self.assertEqual(len(manifest.read_text(encoding="utf-8").splitlines()), 4)

    def test_missing_approval_blocks_every_move(self):
        Path(str(self.preview) + ".approved.json").unlink()
        with self.assertRaisesRegex(RuntimeError, "Approved cached preview is missing"):
            execute(self.db, self.src, self.dst, self.preview, self.runs)
        self.assertTrue(all(p.exists() for p in self.files))

    def test_mutated_source_blocks_every_move(self):
        self.files[-1].write_bytes(b"altered contents")
        with self.assertRaises(RuntimeError):
            execute(self.db, self.src, self.dst, self.preview, self.runs)
        self.assertTrue(all(p.exists() for p in self.files))

    def test_changed_destination_blocks_all_moves(self):
        (self.dst / "EA_OVERRIDE").mkdir()
        (self.dst / "EA_OVERRIDE" / self.files[-1].name).write_bytes(b"already here")
        with self.assertRaisesRegex(RuntimeError, "Destination already exists"):
            execute(self.db, self.src, self.dst, self.preview, self.runs)
        self.assertTrue(all(p.exists() for p in self.files))

    def test_other_destination_root_blocks_all_moves(self):
        other = self.root / "different"
        other.mkdir()
        with self.assertRaisesRegex(RuntimeError, "differs from Preview"):
            execute(self.db, self.src, other, self.preview, self.runs)
        self.assertTrue(all(p.exists() for p in self.files))

    def test_symlink_destination_added_after_preview_is_rejected(self):
        outside = self.root / "outside"
        outside.mkdir()
        alias = self.dst / "EA_OVERRIDE"
        try:
            alias.symlink_to(outside, target_is_directory=True)
        except (OSError, NotImplementedError):
            self.skipTest("Directory symlinks unavailable on this host")
        with self.assertRaisesRegex(RuntimeError, "escapes selected root"):
            execute(self.db, self.src, self.dst, self.preview, self.runs)
        self.assertEqual(list(outside.iterdir()), [])
        self.assertTrue(all(p.exists() for p in self.files))

    def test_symlink_destination_present_during_preview_is_rejected(self):
        outside = self.root / "outside"
        outside.mkdir()
        alias = self.dst / "EA_OVERRIDE"
        try:
            alias.symlink_to(outside, target_is_directory=True)
        except (OSError, NotImplementedError):
            self.skipTest("Directory symlinks unavailable on this host")
        with self.assertRaisesRegex(RuntimeError, "escapes selected root"):
            preview(self.db, self.src, self.dst, self.runs)
        self.assertEqual(list(outside.iterdir()), [])
        self.assertTrue(all(p.exists() for p in self.files))


class CachedDuplicateChecks(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.src = self.root / "source"
        self.dest = self.root / "dest"
        self.out = self.root / "runs"
        for folder in (self.src, self.dest, self.out):
            folder.mkdir()
        self.db = self.root / "library.sqlite"
        self.saved = self.root / "saved.package"
        self.saved.write_bytes(b"original binary")
        self.copy = self.src / "duplicate.package"
        self.copy.write_bytes(self.saved.read_bytes())
        self.normal = self.src / "normal.package"
        self.normal.write_bytes(b"another package")
        with sqlite3.connect(self.db) as conn:
            conn.execute("CREATE TABLE scan_packages_live (source_path TEXT, record_json TEXT, package_hash TEXT)")
            conn.execute("INSERT INTO scan_packages_live VALUES (?,?,?)",
                         (str(self.saved), json.dumps({"Content Type": "EA Override"}),
                          hashlib.sha256(self.saved.read_bytes()).hexdigest()))
            conn.execute("INSERT INTO scan_packages_live VALUES (?,?,?)",
                         (str(self.normal), json.dumps({"Content Type": "EA Override"}),
                          hashlib.sha256(self.normal.read_bytes()).hexdigest()))
            sig = [(_path_key(self.src), _path_key(p), *_file_signature(p))
                   for p in (self.copy, self.normal)]
            commit_routing_batch(conn, sig)

    def test_independent_hash_match_allows_duplicate_route(self):
        draft = preview(self.db, self.src, self.dest, self.out)
        self.assertIn("DUPLICATES", draft.read_text(encoding="utf-8"))
        execute(self.db, self.src, self.dest, draft, self.out)
        self.assertTrue(self.saved.is_file())
        self.assertEqual(len(list(self.dest.rglob("*.package"))), 2)

    def test_missing_independent_saved_hash_blocks_preview(self):
        with sqlite3.connect(self.db) as conn:
            conn.execute("DELETE FROM scan_packages_live WHERE source_path=?", (str(self.saved),))
        with self.assertRaisesRegex(RuntimeError, "No independent saved SHA256 match"):
            preview(self.db, self.src, self.dest, self.out)
        self.assertTrue(self.copy.exists())


class PreviewParserChecks(unittest.TestCase):
    def test_headerless_first_row_is_not_dropped(self):
        text = "EA_OVERRIDE\tNo\t/a.package\t/target/a.package\tREADY\n"
        self.assertEqual(len(parse_preview_rows(text)), 1)

    def test_legacy_header_is_skipped(self):
        text = "Category\tDuplicate\tSource\tDestination\tNotes\nEA_OVERRIDE\tNo\t/a.package\t/target/a.package\tREADY\n"
        self.assertEqual(len(parse_preview_rows(text)), 1)


if __name__ == "__main__":
    unittest.main()
