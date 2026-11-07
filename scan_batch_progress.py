"""Resumable, bounded SQLite-primary scanner batches (R23).

Progress is tied to the selected SQLite Library and the exact source root.
Entries are saved only inside the successful scanner SQLite write transaction.
No source files are moved or modified by this module.
"""
from __future__ import print_function

import os
import sqlite3


def _path_key(path):
    return os.path.normcase(os.path.normpath(os.path.abspath(str(path)))).casefold()


def _file_signature(path):
    stat = os.stat(path)
    return int(stat.st_size), int(getattr(stat, 'st_mtime_ns', round(stat.st_mtime * 1000000000)))


def _has_progress_table(conn):
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='scan_batch_progress'"
    ).fetchone() is not None


def select_next_batch(all_paths, sqlite_db, source_root, limit):
    """Pick the next unprocessed or changed files without trusting legacy Library rows.

    Legacy/migrated Library packages are NOT considered processed by this feature;
    their previous metadata may be stale. Only an R23+ completed scan marks progress.
    Returns selected paths and their *pre-scan* signatures for atomic commit.
    """
    limit = int(limit)
    if limit <= 0:
        raise ValueError('A bounded batch requires a positive limit')
    if not os.path.isfile(str(sqlite_db)):
        raise FileNotFoundError('Open SQLite Library is missing: ' + str(sqlite_db))
    source_key = _path_key(source_root)
    done = {}
    with sqlite3.connect(str(sqlite_db), timeout=30) as conn:
        if _has_progress_table(conn):
            done = {
                str(path): (int(size), int(mtime))
                for path, size, mtime in conn.execute(
                    'SELECT source_path, file_size, mtime_ns FROM scan_batch_progress WHERE source_root=?',
                    (source_key,),
                )
            }
    selected = []
    signatures = []
    skipped_unchanged = 0
    pending = 0
    for path in all_paths:
        key = _path_key(path)
        # If a file disappears after discovery, let the scanner handle/report it.
        # Do not mark a non-existent file as completed.
        try:
            size, mtime = _file_signature(path)
        except OSError:
            pending += 1
            if len(selected) < limit:
                selected.append(path)
            continue
        if done.get(key) == (size, mtime):
            skipped_unchanged += 1
            continue
        pending += 1
        if len(selected) < limit:
            selected.append(path)
            signatures.append((source_key, key, size, mtime))
    return {
        'selected': selected,
        'signatures': signatures,
        'source_root': source_key,
        'unchanged': skipped_unchanged,
        'pending': pending,
        'total': len(all_paths),
    }


def commit_batch_progress(conn, signatures):
    """Called *within* the same transaction as the successful live scan rows."""
    if not signatures:
        return 0
    conn.execute('''CREATE TABLE IF NOT EXISTS scan_batch_progress (
        source_root TEXT NOT NULL,
        source_path TEXT NOT NULL,
        file_size INTEGER NOT NULL,
        mtime_ns INTEGER NOT NULL,
        completed_utc TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (source_root, source_path)
    )''')
    conn.executemany('''INSERT INTO scan_batch_progress
        (source_root, source_path, file_size, mtime_ns) VALUES (?, ?, ?, ?)
        ON CONFLICT(source_root, source_path) DO UPDATE SET
            file_size=excluded.file_size,
            mtime_ns=excluded.mtime_ns,
            completed_utc=CURRENT_TIMESTAMP''', signatures)
    return len(signatures)


def reset_source_progress(sqlite_db, source_root):
    """Explicit user-requested restart for this Library+source root only."""
    with sqlite3.connect(str(sqlite_db), timeout=30) as conn:
        if not _has_progress_table(conn):
            return 0
        cursor = conn.execute(
            'DELETE FROM scan_batch_progress WHERE source_root=?', (_path_key(source_root),)
        )
        return max(0, int(cursor.rowcount or 0))


def is_stale_library_path_match(paths):
    """A known hash whose *every* old path is missing is probably relocated.

    Do not use this when a live path still exists, because that is a real copy.
    """
    return bool(paths) and all(not os.path.isfile(str(path)) for path in paths)


# R27: Last *successfully committed* bounded batch. Routing never infers this
# from all historical progress rows, which could silently select thousands.
def commit_routing_batch(conn, signatures):
    if not signatures:
        return
    conn.execute("""CREATE TABLE IF NOT EXISTS scan_routing_batch (
        source_root TEXT NOT NULL, source_path TEXT NOT NULL,
        file_size INTEGER NOT NULL, mtime_ns INTEGER NOT NULL,
        PRIMARY KEY(source_root, source_path))""")
    root = signatures[0][0]
    if any(sig[0] != root for sig in signatures):
        raise ValueError("Mixed source roots in batch")
    conn.execute('DELETE FROM scan_routing_batch WHERE source_root=?', (root,))
    conn.executemany('INSERT INTO scan_routing_batch (source_root, source_path, file_size, mtime_ns) VALUES (?, ?, ?, ?)', signatures)


def load_routing_batch(sqlite_db, source_root):
    """Fail closed on absent/changed batch files; never select the entire root."""
    root = _path_key(source_root)
    with sqlite3.connect(str(sqlite_db), timeout=30) as conn:
        table = conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='scan_routing_batch'").fetchone()
        if not table:
            raise RuntimeError('No saved batch found. Scan a bounded batch before routing.')
        records = conn.execute('SELECT source_path, file_size, mtime_ns FROM scan_routing_batch WHERE source_root=? ORDER BY source_path', (root,)).fetchall()
    if not records:
        raise RuntimeError('No saved batch found for this source and Library.')
    selected = set()
    for path, size, mtime in records:
        if not os.path.isfile(path):
            # Already moved paths are harmless, but must not broaden scope.
            continue
        current = _file_signature(path)
        if current != (int(size), int(mtime)):
            raise RuntimeError('Saved batch file changed; scan again before routing: ' + path)
        if os.path.commonpath([root, _path_key(path)]) != root:
            raise RuntimeError('Saved batch contains path outside source root: ' + path)
        selected.add(_path_key(path))
    return selected, len(records)
