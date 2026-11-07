"""Reconstruct the verified R27 SQLite scanner from the public core + patch.

The public/legacy scanner stays byte-for-byte unchanged. This script never
alters that scanner; it creates a separate opt-in R27 module.

The patch file is LZMA-compressed JSON, Base64-encoded to keep the staging
branch compact. Both the public base and reconstructed output are SHA256
checked; any mismatch aborts before writing.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import lzma
from pathlib import Path

BASE_SHA256 = "43307adec9c6983814e4c8ee0eaffb5abf6abf0f7516ac7400bc1bbacdf3a40f"
R27_SHA256 = "2607fe5258dfe4fc0886caa5e6b616461a677c9de934182a533b5912c788b9d7"


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def reconstruct(root: Path) -> bytes:
    source = root / "scanner_core_reference" / "Sims3_CC_Health_Scanner_CORE_INTEGRATED_TEST.py"
    patch = root / "scanner_core_reference" / "R27_SQLITE_CORE_PATCH.xz.b64"
    original = source.read_bytes()
    if _sha256(original) != BASE_SHA256:
        raise RuntimeError("Legacy scanner source changed; refusing unverified R27 patch.")
    operations = json.loads(lzma.decompress(base64.b64decode(patch.read_text(encoding="ascii"))))
    lines = original.decode("utf-8-sig").split("\n")
    previous = len(lines)
    for start, end, replacement in reversed(operations):
        if not (0 <= start <= end <= previous):
            raise RuntimeError("R27 patch has overlapping or invalid hunks.")
        lines[start:end] = replacement
        previous = start
    updated = "\n".join(lines).encode("utf-8")
    if _sha256(updated) != R27_SHA256:
        raise RuntimeError("R27 scanner SHA256 mismatch; refusing modified core.")
    return updated


def main():
    parser = argparse.ArgumentParser(description="Materialize the exact tested R27 SQLite scanner core")
    parser.add_argument("--check", action="store_true", help="Verify reconstruction without writing")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    updated = reconstruct(root)
    if not args.check:
        target = root / "scanner_core_reference" / "Sims3_CC_Health_Scanner_R27_SQLITE.py"
        if target.exists() and _sha256(target.read_bytes()) != R27_SHA256:
            raise RuntimeError("Destination scanner already exists with different contents; not replacing.")
        target.write_bytes(updated)
        print("Created:", target)
    print("R27 scanner verified:", _sha256(updated), "bytes:", len(updated))


if __name__ == "__main__":
    main()
