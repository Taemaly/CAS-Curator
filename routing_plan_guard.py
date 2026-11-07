"""R27 TEST_011: Verify DEV actual moves against approved cached preview.

Never approve destinations by file count alone. This validation executes before
any move; changed source signatures, destinations or categories fail closed.
"""
import json
import os
from pathlib import Path
from scan_batch_progress import _path_key, _file_signature, load_routing_batch


def write_approved_plan(preview_path, sqlite_db, source_root, destination_root, planned):
    keys, count = load_routing_batch(sqlite_db, source_root)
    if len(keys) != count or len(planned) != count:
        raise RuntimeError('Saved batch is incomplete; cannot approve routing preview')
    entries = []
    for src, category, destination, _collision in planned:
        key = _path_key(src)
        if key not in keys:
            raise RuntimeError('Preview includes a source outside the batch: ' + src)
        size, mtime = _file_signature(src)
        entries.append({'source':key, 'category':category,
                        'destination':_path_key(destination), 'size':size, 'mtime_ns':mtime})
    if len({entry['source'] for entry in entries}) != count:
        raise RuntimeError('Duplicate or missing sources in routing preview')
    payload = {'format':1,'sqlite_db':_path_key(sqlite_db),
               'source_root':_path_key(source_root),'destination_root':_path_key(destination_root),
               'entries':sorted(entries,key=lambda x:x['source'])}
    plan_path=Path(str(preview_path)+'.approved.json')
    plan_path.write_text(json.dumps(payload,indent=2),encoding='utf-8')
    return plan_path


def validate_approved_plan(preview_path, sqlite_db, source_root, destination_root, physical_plan):
    sidecar=Path(str(preview_path)+'.approved.json')
    if not sidecar.is_file() or not Path(preview_path).is_file():
        raise RuntimeError('Approved cached preview is missing; run Preview Routing again. No files moved.')
    plan=json.loads(sidecar.read_text(encoding='utf-8'))
    if plan.get('format')!=1 or any(plan.get(field)!=_path_key(value) for field,value in (
        ('sqlite_db',sqlite_db),('source_root',source_root),('destination_root',destination_root))):
        raise RuntimeError('Routing source, destination or Library differs from Preview; no files moved.')
    keys,count=load_routing_batch(sqlite_db,source_root)
    if len(keys)!=count:
        raise RuntimeError('Saved batch is incomplete or moved; no files moved.')
    approved=plan.get('entries')
    if not isinstance(approved,list) or len(approved)!=count or len(physical_plan)!=count:
        raise RuntimeError('Routing plan count changed since Preview; no files moved.')
    actual=[]
    for src,cat,dest,_collision in physical_plan:
        key=_path_key(src)
        if key not in keys:
            raise RuntimeError('Routing includes unrelated source: '+str(src))
        size,mtime=_file_signature(src)
        actual.append({'source':key,'category':cat,'destination':_path_key(dest),
                       'size':size,'mtime_ns':mtime})
    if sorted(actual,key=lambda x:x['source'])!=approved:
        raise RuntimeError('Routing destinations or source signatures differ from cached Preview; no files moved. Generate a new preview.')
    print('R27 APPROVED PREVIEW VERIFIED: {} exact source/category/destination/signature matches'.format(count),flush=True)
    return count
