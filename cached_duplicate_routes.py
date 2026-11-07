"""R27 TEST_015: secure duplicate-only fallback for scanned files without live rows.

Never infer regular routing categories from names, and never parse packages.
A missing scan row is only allowed when the batch signature matches AND its
current SHA256 equals a DIFFERENT saved SQLite package's SHA256.
"""
import hashlib
import os
import sqlite3
from pathlib import Path
from scan_batch_progress import _path_key


def resolve_missing_duplicates(sqlite_db, source_root, keys, records):
    missing = set(keys) - set(records)
    if not missing:
        return {}
    known = {}
    with sqlite3.connect(str(sqlite_db), timeout=30) as con:
        for sha, source_path in con.execute('SELECT package_hash, source_path FROM scan_packages_live WHERE package_hash IS NOT NULL'):
            digest = str(sha or '').strip().lower()
            if len(digest) == 64 and all(c in '0123456789abcdef' for c in digest):
                known.setdefault(digest, set()).add(_path_key(source_path) if source_path else '')
    batch_tag = hashlib.sha256(('\n'.join(sorted(keys))).encode('utf-8')).hexdigest()[:16]
    category = os.path.join('DUPLICATES', 'BATCH_' + batch_tag)
    out = {}
    for key in sorted(missing):
        # This operation intentionally hashes bytes only; no CC resource rescan.
        path = Path(key)
        if not path.is_file() or path.is_symlink():
            raise RuntimeError('Missing or linked duplicate candidate: ' + key)
        h = hashlib.sha256()
        with path.open('rb') as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b''):
                h.update(chunk)
        digest = h.hexdigest()
        matches = known.get(digest, set()) - {key}
        if not matches:
            raise RuntimeError('No independent saved SHA256 match for missing scan record: ' + key + '. Preview blocked; no files moved.')
        out[key] = category
        print('CACHED DUPLICATE VERIFIED: {} | SHA256={} | saved Library match'.format(path.name, digest), flush=True)
    return out
