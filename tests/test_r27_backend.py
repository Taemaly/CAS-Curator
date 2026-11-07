"""Regression checks for R27 backend modules staged for the public port.

These synthetic tests do not establish compatibility with every real Sims 3
package, game version, or installed Windows build.
"""
from __future__ import annotations

import hashlib
import tempfile
import unittest
from pathlib import Path

from merge_builder_core import (
    PACKAGE_MANIFEST_TYPE,
    MergeResource,
    _read_sims3_dbpf_index,
    _write_sims3_dbpf,
    build_merge_from_temporary_copies,
)
from sqlite_store import (
    create_fresh_sqlite_library,
    load_sqlite_library,
    validate_sqlite_library,
)

# Uncompressed S3PI-compatible storage word (marker 1, compression 0).
STORAGE_WORD = 0x00010000


def resource(kind: int, instance: int, contents: bytes) -> MergeResource:
    return MergeResource((kind, 0, instance), contents, len(contents), STORAGE_WORD)


def stored_resources(path: Path):
    entries, data = _read_sims3_dbpf_index(path)
    return {
        (entry.resource_type, entry.resource_group, entry.instance): (
            data[entry.offset:entry.offset + entry.stored_size],
            entry.mem_size,
            entry.compression_word,
        )
        for entry in entries
    }


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class MergeBuilderBackendTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.a = self.root / "original-a.package"
        self.b = self.root / "original-b.package"
        _write_sims3_dbpf(self.a, [
            resource(0xABCDE001, 1, b"alpha"),
            resource(0xABCDE003, 3, b"shared"),
            resource(PACKAGE_MANIFEST_TYPE, 1, b"manifest A"),
        ])
        _write_sims3_dbpf(self.b, [
            resource(0xABCDE002, 2, b"beta"),
            resource(0xABCDE003, 3, b"shared"),
            resource(PACKAGE_MANIFEST_TYPE, 1, b"manifest B"),
        ])

    def test_merges_exact_resources_and_preserves_originals(self):
        originals = {p: sha256(p) for p in (self.a, self.b)}
        output = self.root / "merged.package"
        report = build_merge_from_temporary_copies(
            [self.a, self.b], output, temp_root=self.root / "working"
        )
        self.assertEqual(report["resources_written"], 3)
        self.assertEqual(report["manifests_omitted"], 2)
        self.assertEqual(report["exact_duplicates_omitted"], 1)
        self.assertEqual(report["originals_verified"], 2)
        self.assertEqual(sha256(output), report["output_sha256"])
        actual = stored_resources(output)
        self.assertEqual(len(actual), 3)
        self.assertEqual(actual[(0xABCDE001, 0, 1)], (b"alpha", 5, STORAGE_WORD))
        self.assertEqual(actual[(0xABCDE002, 0, 2)], (b"beta", 4, STORAGE_WORD))
        self.assertEqual(actual[(0xABCDE003, 0, 3)], (b"shared", 6, STORAGE_WORD))
        self.assertFalse(any(key[0] == PACKAGE_MANIFEST_TYPE for key in actual))
        self.assertEqual({p: sha256(p) for p in originals}, originals)

    def test_conflicting_tgi_blocks_output(self):
        _write_sims3_dbpf(self.b, [
            resource(0xABCDE001, 1, b"different bytes"),
        ])
        output = self.root / "blocked.package"
        before = {p: sha256(p) for p in (self.a, self.b)}
        with self.assertRaisesRegex(ValueError, "Real TGI collision"):
            build_merge_from_temporary_copies(
                [self.a, self.b], output, temp_root=self.root / "working"
            )
        self.assertFalse(output.exists())
        self.assertEqual({p: sha256(p) for p in before}, before)

    def test_refuses_to_overwrite_an_original(self):
        before = sha256(self.a)
        with self.assertRaisesRegex(ValueError, "cannot replace"):
            build_merge_from_temporary_copies(
                [self.a, self.b], self.a, temp_root=self.root / "working"
            )
        self.assertEqual(sha256(self.a), before)


class SQLiteBackendTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.db = Path(self.temp.name) / "CAS_Curator.db"

    def test_create_validate_and_load_empty_library(self):
        result = create_fresh_sqlite_library(self.db)
        self.assertEqual(result.db_path, self.db.resolve())
        validation = validate_sqlite_library(self.db)
        self.assertEqual(validation["integrity"], "ok")
        self.assertEqual(validation["cc_master_rows"], 0)
        self.assertEqual(validation["casp_rows"], 0)
        library = load_sqlite_library(self.db)
        self.assertEqual(len(library.records), 0)
        self.assertIn("Status", library.headers)

    def test_recreate_makes_backup_of_existing_database(self):
        create_fresh_sqlite_library(self.db)
        original = sha256(self.db)
        result = create_fresh_sqlite_library(self.db)
        self.assertIsNotNone(result.backup_path)
        self.assertTrue(result.backup_path.is_file())
        self.assertEqual(sha256(result.backup_path), original)
        self.assertEqual(validate_sqlite_library(self.db)["integrity"], "ok")


if __name__ == "__main__":
    unittest.main()
