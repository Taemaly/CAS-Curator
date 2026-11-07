"""R27 TEST_015: real DEV bounded-batch routing without rescanning packages.

Only executes a previously approved cached plan, reconstructing categories from
SQLite scan records and validating all source signatures before the first move.
"""
import json
import os
import sqlite3
import shutil
from datetime import datetime
from pathlib import Path
from scan_batch_progress import load_routing_batch, _path_key
from cached_routing_preview import _route_category, _plan, _validate_child
from routing_plan_guard import validate_approved_plan
from cached_duplicate_routes import resolve_missing_duplicates


def _original_case(path):
    """Recover original on-disk case where scanned SQLite keys are lowercased."""
    p = Path(path)
    try:
        names = {item.name.casefold(): item.name for item in p.parent.iterdir()}
        return str(p.with_name(names.get(p.name.casefold(), p.name)))
    except OSError:
        return str(p)


def execute(sqlite_db, source_root, destination_root, approved_preview, output_root):
    approved_preview = Path(approved_preview)
    keys, total = load_routing_batch(sqlite_db, source_root)
    if len(keys) != total:
        raise RuntimeError('Saved batch has missing files; no files moved')
    with sqlite3.connect(str(sqlite_db), timeout=30) as con:
        con.row_factory = sqlite3.Row
        records = {}
        for row in con.execute('SELECT source_path, record_json FROM scan_packages_live'):
            path = row['source_path']
            if path and _path_key(path) in keys:
                key = _path_key(path)
                if key in records:
                    raise RuntimeError('Duplicate SQLite scan entry: ' + key)
                records[key] = json.loads(row['record_json'])
    duplicates = resolve_missing_duplicates(sqlite_db, source_root, keys, records)
    items = []
    for key in sorted(keys):
        _validate_child(key, source_root)
        category = duplicates[key] if key in duplicates else _route_category(records[key])
        items.append((_original_case(key), category))
    plan = _plan(items, source_root, destination_root)
    validate_approved_plan(str(approved_preview), sqlite_db, source_root, destination_root, plan)
    if not plan:
        raise RuntimeError('Empty approved batch; no files moved')

    # Preflight ALL inputs and destinations, not only the first. No overwrites.
    for src, cat, dst, _collision in plan:
        if not os.path.isfile(src):
            raise RuntimeError('Source disappeared: ' + src)
        if os.path.lexists(dst):
            raise RuntimeError('Destination already exists: ' + dst + '. Refresh Preview; no files moved.')
        if os.path.islink(src):
            raise RuntimeError('Symlink source is not supported: ' + src)
        _validate_child(src, source_root)
        _validate_child(dst, destination_root)
    # One root per approved plan; no broad source scanning.
    stamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S-%f')
    session = Path(output_root) / ('Cached_Routing_Move_' + stamp)
    session.mkdir(parents=True, exist_ok=False)
    manifest = session / 'Move_Manifest.txt'
    log_path = session / 'Health_Scan_Log.txt'
    actions = []
    print('R27 TEST_016 CACHED ROUTING: approved plan only; NO PACKAGE RESCAN', flush=True)
    print('TOTAL PACKAGES TO SCAN: 0', flush=True)
    print('ROUTING APPROVED FILES: {}'.format(len(plan)), flush=True)
    try:
        print('MOVING FILES...', flush=True)
        for src, cat, dst, _ in plan:
            # Refuse a mid-move destination collision rather than overwrite it.
            if os.path.lexists(dst):
                raise RuntimeError('Destination appeared during routing: ' + dst)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            # Move may involve different volumes. shutil.move uses a copy/remove fallback.
            shutil.move(src, dst)
            actions.append((src, dst, cat, 'MOVED'))
            print('MOVED: ' + dst, flush=True)
    except Exception as exc:
        errors=[]
        for src,dst,_,_ in reversed(actions):
            try:
                if os.path.lexists(src):
                    raise RuntimeError('original source path exists, cannot restore')
                shutil.move(dst,src)
            except Exception as undo_exc:
                errors.append('{}: {}'.format(dst,undo_exc))
        if errors:
            raise RuntimeError('Move interrupted; recovery incomplete, inspect files manually. Cause: {}. Restore issues: {}'.format(exc,'; '.join(errors)))
        raise RuntimeError('Routing cancelled and completed moves restored: {}'.format(exc))
    manifest.write_text('Source\tDestination\tCategory\tStatus\n' + ''.join('{}\t{}\t{}\t{}\n'.format(*a) for a in actions),encoding='utf-8')
    log_path.write_text('R27 TEST_016 cached routing (approved plan), no rescan\n'+'\n'.join('MOVED: '+a[1] for a in actions)+'\n',encoding='utf-8')
    print('Move manifest saved: ' + str(manifest),flush=True)
    print('Console log saved: ' + str(log_path),flush=True)
    print('Scan time: 0 minutes 0 seconds',flush=True)
    print('Cached routing finished; moved {} approved files, scanned 0 packages.'.format(len(actions)), flush=True)
    return manifest
