from __future__ import annotations

import hashlib
import os
import shutil
import struct
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from temp_working import managed_temporary_directory

TGI = Tuple[int, int, int]

# Sims 3 Package Manifest. Multiple packages commonly carry this launcher/package
# metadata under the same TGI, so it is expected to collide when packages are
# combined. Merge Builder will omit Package Manifest resources from a future
# merged output rather than arbitrarily choosing one package's metadata.
PACKAGE_MANIFEST_TYPE = 0x73E93EEB
RESOURCE_TYPE_NAMES = {
    PACKAGE_MANIFEST_TYPE: "Package Manifest",
}


@dataclass
class PackageScan:
    path: Path
    resource_count: int = 0
    unreadable_count: int = 0
    resource_error: str = ""
    error: str = ""


@dataclass
class ResourceOccurrence:
    package_path: Path
    resource_hash: str
    size: int


@dataclass(frozen=True)
class DbpfIndexEntry:
    resource_type: int
    resource_group: int
    instance: int
    offset: int
    stored_size: int
    mem_size: int
    compression_word: int


@dataclass(frozen=True)
class MergeResource:
    key: TGI
    stored_bytes: bytes
    mem_size: int
    compression_word: int


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _resource_digest(entry: DbpfIndexEntry, stored: bytes) -> str:
    h = hashlib.sha256()
    h.update(struct.pack("<III", entry.stored_size, entry.mem_size, entry.compression_word))
    h.update(stored)
    return h.hexdigest()


def _expected_storage_word(stored_size: int, mem_size: int) -> int:
    """Return the canonical S3PI storage word for an unchanged resource chunk."""
    compressed = 0xFFFF if stored_size < mem_size else 0x0000
    return (0x0001 << 16) | compressed


def _validate_storage_entry(entry: DbpfIndexEntry, path: Path, key: TGI) -> int:
    """Reject storage/index combinations this first writer test should not guess about."""
    if entry.mem_size == 0xFFFFFFFF or entry.stored_size == 0:
        raise ValueError(f"{path.name}: unsupported special/empty DBPF resource { _tgi_text(key) }")
    if entry.stored_size > entry.mem_size:
        raise ValueError(f"{path.name}: unsupported DBPF storage sizes for { _tgi_text(key) }")

    expected = _expected_storage_word(entry.stored_size, entry.mem_size)
    if entry.compression_word != expected:
        actual_compression = entry.compression_word & 0xFFFF
        actual_marker = (entry.compression_word >> 16) & 0xFFFF
        expected_compression = expected & 0xFFFF
        raise ValueError(
            f"{path.name}: unsupported DBPF storage flags for { _tgi_text(key) } "
            f"(compression 0x{actual_compression:04X}, marker 0x{actual_marker:04X}; "
            f"expected compression 0x{expected_compression:04X}, marker 0x0001)"
        )
    return expected


def _prepare_merge_resources(paths: List[Path]):
    """Return deterministic resources for a new merged package.

    Package Manifest resources are intentionally omitted. Exact duplicate TGIs
    are deduplicated. Any same-TGI/different-data collision or unusual storage
    record blocks this first writer test rather than guessing.
    """
    planned: Dict[TGI, Tuple[str, MergeResource]] = {}
    manifests_omitted = 0
    exact_duplicates_omitted = 0

    for path in paths:
        entries, package_bytes = _read_sims3_dbpf_index(path)
        for entry in entries:
            key = (entry.resource_type, entry.resource_group, entry.instance)
            if _is_package_manifest(key):
                manifests_omitted += 1
                continue
            stored = package_bytes[entry.offset:entry.offset + entry.stored_size]
            if len(stored) != entry.stored_size:
                raise ValueError(f"{path.name}: stored resource is truncated")
            # S3PI treats files smaller than Memsize as compressed and otherwise
            # as uncompressed. Refuse any storage flags that do not match that
            # canonical Sims 3/S3PI form in this first writer test.
            compression_word = _validate_storage_entry(entry, path, key)
            digest = _resource_digest(entry, stored)
            resource = MergeResource(key, bytes(stored), int(entry.mem_size), compression_word)
            previous = planned.get(key)
            if previous is None:
                planned[key] = (digest, resource)
            elif previous[0] == digest:
                exact_duplicates_omitted += 1
            else:
                raise ValueError(f"Real TGI collision blocks merge: { _tgi_text(key) }")

    resources = [planned[key][1] for key in sorted(planned)]
    return resources, manifests_omitted, exact_duplicates_omitted


def _write_sims3_dbpf(path: Path, resources: List[MergeResource]):
    """Write a new DBPF 2.0 package using S3PI-compatible index fields."""
    header = bytearray(96)
    header[0:4] = b"DBPF"
    struct.pack_into("<I", header, 4, 2)   # Major
    struct.pack_into("<I", header, 8, 0)   # Minor
    struct.pack_into("<I", header, 36, len(resources))
    struct.pack_into("<I", header, 40, 0)  # old index position field unused
    struct.pack_into("<I", header, 60, 3)  # Sims 3 index version

    index_rows = []
    with path.open("wb") as f:
        f.write(header)
        for resource in resources:
            key = resource.key
            offset = f.tell()
            stored = resource.stored_bytes
            f.write(stored)
            stored_size = len(stored)
            mem_size = resource.mem_size
            expected_storage_word = _expected_storage_word(stored_size, mem_size)
            if resource.compression_word != expected_storage_word:
                raise ValueError(f"Writer plan contains unsupported storage metadata for { _tgi_text(key) }")
            size_field = stored_size | 0x80000000
            compression_word = resource.compression_word
            rtype, group, instance = key
            index_rows.append((
                rtype, group, (instance >> 32) & 0xFFFFFFFF, instance & 0xFFFFFFFF,
                offset, size_field, mem_size, compression_word,
            ))

        index_pos = f.tell()
        f.write(struct.pack("<I", 0))  # full entries; no constant T/G/I fields
        for row in index_rows:
            f.write(struct.pack("<8I", *row))
        index_size = 4 + (32 * len(index_rows))

        f.seek(44)
        f.write(struct.pack("<I", index_size))
        f.seek(64)
        f.write(struct.pack("<I", index_pos))
        f.flush()
        os.fsync(f.fileno())


def _verify_merged_package(path: Path, expected: List[MergeResource]):
    entries, package_bytes = _read_sims3_dbpf_index(path)
    if len(entries) != len(expected):
        raise ValueError(f"Verification failed: expected {len(expected)} resources, found {len(entries)}")
    expected_map = {r.key: r for r in expected}
    if len(expected_map) != len(expected):
        raise ValueError("Verification failed: merge plan contains duplicate TGIs")
    seen = set()
    for entry in entries:
        key = (entry.resource_type, entry.resource_group, entry.instance)
        if key in seen:
            raise ValueError(f"Verification failed: duplicate output TGI { _tgi_text(key) }")
        seen.add(key)
        exp = expected_map.get(key)
        if exp is None:
            raise ValueError(f"Verification failed: unexpected output TGI { _tgi_text(key) }")
        stored = package_bytes[entry.offset:entry.offset + entry.stored_size]
        if stored != exp.stored_bytes or entry.mem_size != exp.mem_size:
            raise ValueError(f"Verification failed: resource changed { _tgi_text(key) }")
        if entry.compression_word != exp.compression_word:
            raise ValueError(f"Verification failed: storage metadata changed { _tgi_text(key) }")
        if _is_package_manifest(key):
            raise ValueError("Verification failed: Package Manifest was written")
    if seen != set(expected_map):
        raise ValueError("Verification failed: one or more planned resources are missing")
    return len(entries)


def build_merge_from_temporary_copies(original_paths: List[Path], output_path: Path, *, temp_root: Optional[Path] = None):
    """Build and verify a merged package without writing to any original package."""
    originals = [Path(p).resolve() for p in original_paths]
    before = {str(p).casefold(): _sha256_file(p) for p in originals}
    output_resolved = output_path.resolve()
    if any(str(output_resolved).casefold() == str(p).casefold() for p in originals):
        raise ValueError("Output path cannot replace one of the original packages")

    with managed_temporary_directory(root=temp_root, prefix="CASCurator_Merge_") as temp_root:
        copies = []
        for i, original in enumerate(originals, start=1):
            copy_path = temp_root / f"{i:03d}_{original.name}"
            shutil.copy2(original, copy_path)
            copies.append(copy_path)

        resources, manifests_omitted, exacts_omitted = _prepare_merge_resources(copies)
        if not resources:
            raise ValueError("No mergeable resources remain after safety filtering")
        temp_output = temp_root / "CASCurator_Merged_TEST.package"
        _write_sims3_dbpf(temp_output, resources)
        written = _verify_merged_package(temp_output, resources)

        after = {str(p).casefold(): _sha256_file(p) for p in originals}
        changed = [p.name for p in originals if before[str(p).casefold()] != after[str(p).casefold()]]
        if changed:
            raise RuntimeError("Safety verification failed: original package hash changed: " + ", ".join(changed))

        output_path.parent.mkdir(parents=True, exist_ok=True)
        stage = output_path.with_name(output_path.name + ".cascurator_tmp")
        try:
            shutil.copy2(temp_output, stage)
            os.replace(stage, output_path)
        finally:
            if stage.exists():
                stage.unlink(missing_ok=True)

    # Verify the final copied output too.
    _verify_merged_package(output_path, resources)
    return {
        "resources_written": written,
        "manifests_omitted": manifests_omitted,
        "exact_duplicates_omitted": exacts_omitted,
        "originals_verified": len(originals),
        "output_sha256": _sha256_file(output_path),
    }


def _read_u32(data: bytes, offset: int) -> int:
    if offset < 0 or offset + 4 > len(data):
        raise ValueError("Unexpected end of DBPF data")
    return struct.unpack_from("<I", data, offset)[0]


def _read_sims3_dbpf_index(path: Path):
    """Read the Sims 3 DBPF 2.0 resource index without CLR/S3PI.

    This is intentionally index-only. It never rewrites/decompresses resources.
    For exact-duplicate detection we hash the bytes exactly as stored in the
    package; anything that differs is treated conservatively as a collision.
    """
    data = path.read_bytes()
    if len(data) < 96 or data[:4] != b"DBPF":
        raise ValueError("Not a valid DBPF package")

    major = _read_u32(data, 4)
    minor = _read_u32(data, 8)
    if (major, minor) != (2, 0):
        raise ValueError(f"Unsupported DBPF version {major}.{minor}; expected Sims 3 DBPF 2.0")

    index_count = _read_u32(data, 36)
    index_size = _read_u32(data, 44)
    index_version = _read_u32(data, 60)
    index_pos = _read_u32(data, 64)

    if index_count == 0:
        return [], data
    if index_version not in (0, 3):
        raise ValueError(f"Unsupported Sims 3 DBPF index version {index_version}")
    if index_pos == 0 or index_pos + 4 > len(data):
        raise ValueError("DBPF index position is invalid")
    if index_size and index_pos + index_size > len(data):
        raise ValueError("DBPF index extends past end of package")

    cursor = index_pos
    index_type = _read_u32(data, cursor)
    cursor += 4
    if index_type & ~0xFF:
        raise ValueError(f"Unsupported DBPF index type 0x{index_type:08X}")

    constants = [None] * 8
    for field in range(8):
        if index_type & (1 << field):
            constants[field] = _read_u32(data, cursor)
            cursor += 4

    entries = []
    for _ in range(index_count):
        fields = list(constants)
        for field in range(8):
            if fields[field] is None:
                fields[field] = _read_u32(data, cursor)
                cursor += 4

        rtype, group, inst_hi, inst_lo, chunk_offset, size_field, mem_size, compressed = fields
        stored_size = size_field & 0x7FFFFFFF
        if chunk_offset > len(data) or chunk_offset + stored_size > len(data):
            raise ValueError("A DBPF resource points outside the package")
        entries.append(DbpfIndexEntry(
            int(rtype), int(group), (int(inst_hi) << 32) | int(inst_lo),
            int(chunk_offset), int(stored_size), int(mem_size), int(compressed),
        ))

    if index_size and cursor > index_pos + index_size:
        raise ValueError("DBPF index entries exceed the declared index size")
    return entries, data


def _resource_type_name(key: TGI) -> str:
    return RESOURCE_TYPE_NAMES.get(key[0], "")


def _is_package_manifest(key: TGI) -> bool:
    return key[0] == PACKAGE_MANIFEST_TYPE


def _tgi_text(key: TGI) -> str:
    rtype, group, instance = key
    base = f"0x{rtype:08X} / 0x{group:08X} / 0x{instance:016X}"
    name = _resource_type_name(key)
    return f"{name} — {base}" if name else base


def analyse_packages(paths: List[Path], progress=None):
    scans: List[PackageScan] = []
    by_tgi: Dict[TGI, List[ResourceOccurrence]] = defaultdict(list)

    for pkg_index, path in enumerate(paths):
        if progress is not None:
            progress.setLabelText(f"Reading DBPF index: {path.name}…")
            progress.setValue(pkg_index)
            if progress.wasCanceled():
                raise RuntimeError("Preview cancelled")

        scan = PackageScan(path=path)
        try:
            entries, package_bytes = _read_sims3_dbpf_index(path)
            scan.resource_count = len(entries)
            for entry in entries:
                try:
                    stored = package_bytes[entry.offset:entry.offset + entry.stored_size]
                    if len(stored) != entry.stored_size:
                        raise ValueError("Stored resource is truncated")
                    # Hash exact stored bytes plus storage metadata. This is
                    # deliberately conservative: if two resources do not match
                    # byte-for-byte as stored, Preview calls them a collision.
                    key = (entry.resource_type, entry.resource_group, entry.instance)
                    _validate_storage_entry(entry, path, key)
                    digest = _resource_digest(entry, stored)
                    by_tgi[key].append(ResourceOccurrence(path, digest, entry.stored_size))
                except Exception as exc:
                    scan.unreadable_count += 1
                    if not scan.resource_error:
                        scan.resource_error = f"{type(exc).__name__}: {exc}"
        except Exception as exc:
            scan.error = f"{type(exc).__name__}: {exc}"
        scans.append(scan)

    exact_duplicate_keys: Dict[TGI, List[ResourceOccurrence]] = {}
    conflicting_keys: Dict[TGI, List[ResourceOccurrence]] = {}
    expected_manifest_keys: Dict[TGI, List[ResourceOccurrence]] = {}
    for key, occurrences in by_tgi.items():
        if len(occurrences) < 2:
            continue
        # Package Manifest is launcher/package metadata, not a normal CAS runtime
        # resource. A future writer will omit all Package Manifest resources from
        # merged output, so a manifest duplicate/collision is expected and does
        # not block the merge preview.
        if _is_package_manifest(key):
            expected_manifest_keys[key] = occurrences
            continue
        hashes = {item.resource_hash for item in occurrences}
        if len(hashes) == 1:
            exact_duplicate_keys[key] = occurrences
        else:
            conflicting_keys[key] = occurrences

    if progress is not None:
        progress.setValue(len(paths))

    return scans, by_tgi, exact_duplicate_keys, conflicting_keys, expected_manifest_keys


