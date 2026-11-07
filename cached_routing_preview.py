"""R27 TEST_015: DEV cached batch Routing Preview, read-only.

Refuse unsupported/incomplete records rather than rescanning or guessing routes.
Does not move files, open packages, extract thumbnails, or write SQLite.
"""
import json
import os
import sqlite3
from datetime import datetime
from pathlib import Path
from scan_batch_progress import load_routing_batch, _path_key
from cached_duplicate_routes import resolve_missing_duplicates


def _validate_child(source, root):
    if os.path.commonpath([_path_key(source), _path_key(root)]) != _path_key(root):
        raise RuntimeError('Source path outside scan root: ' + str(source))


def _route_category(row):
    content = str(row.get('Content Type') or '').strip()
    cc_type = str(row.get('CC Type') or '').strip()
    health = str(row.get('Health Status') or '').strip()
    skin = str(row.get('Skin Type') or '').strip()
    casp = str(row.get('CASP Instance IDs') or '').strip()
    # Corrupt resource subcategories depend on a larger set of scanner-specific
    # broken-resource flags. Do not collapse distinct destinations into UNKNOWN.
    if health.upper() == 'BROKEN':
        raise ValueError('BROKEN package needs saved detailed routing classification')
    if content == 'EA Match':
        return 'EA_MATCH'
    if content == 'EA Override':
        return 'EA_OVERRIDE'
    if content == 'EA Candidate' and skin not in ('Non-Default', 'Default Replacement'):
        return 'EA_CANDIDATE'
    if content == 'Custom Content' and cc_type and cc_type != 'Unknown':
        root = 'CASP_OBSERVED' if casp else 'CUSTOM_CONTENT'
        return os.path.join(root, 'Custom_Sim' if cc_type == 'Custom Sim' else cc_type)
    raise ValueError('Unknown, skin or other special route requires verified scanner routing data')


def _plan(items, source_root, destination_root):
    # Matches scanner routing collision handling, preserving shortest parent suffix.
    plan=[]; collisions={}
    for path, category in items:
        dest=os.path.join(destination_root, category, os.path.basename(path))
        plan.append([path,category,dest,False])
        collisions.setdefault(os.path.normpath(dest).casefold(),[]).append(len(plan)-1)
    for indices in collisions.values():
        if len(indices)<2: continue
        parent_parts=[]
        for i in indices:
            relative=os.path.relpath(plan[i][0],source_root)
            if relative.startswith('..'+os.sep) or os.path.isabs(relative):
                raise RuntimeError('Source path escapes scan root')
            parent=os.path.dirname(relative)
            parent_parts.append([] if parent in ('', '.') else parent.split(os.sep))
        chosen=None
        for depth in range(1,max(map(len,parent_parts))+2):
            tails=[tuple(p[-depth:]) if p else ('_SCAN_ROOT_',) for p in parent_parts]
            if len({os.path.normpath(os.path.join(*tail)).casefold() for tail in tails})==len(tails):
                chosen=tails;break
        if chosen is None:raise RuntimeError('Destination filename collision cannot be resolved')
        for i,parts in zip(indices,chosen):
            path,cat,_,_=plan[i]
            plan[i][2]=os.path.join(destination_root,cat,'SOURCE_FOLDERS',*parts,os.path.basename(path))
            plan[i][3]=True
    destinations=[os.path.normpath(item[2]).casefold() for item in plan]
    if len(destinations)!=len(set(destinations)):
        raise RuntimeError('Duplicate destinations in routing plan')
    return plan


def preview(sqlite_db,source_root,destination_root,preview_root):
    keys,saved_count=load_routing_batch(sqlite_db,source_root)
    if len(keys)!=saved_count:
        raise RuntimeError('Some saved batch files were moved or missing. Cached preview blocked; no rescan was started.')
    with sqlite3.connect(str(sqlite_db),timeout=30) as con:
        con.row_factory=sqlite3.Row
        names={row['name'] for row in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if 'scan_packages_live' not in names:
            raise RuntimeError('No SQLite live scan records found for cached preview.')
        records={}
        for item in con.execute('SELECT source_path,record_json FROM scan_packages_live'):
            if item['source_path']:
                key=_path_key(item['source_path'])
                if key in keys:
                    if key in records: raise RuntimeError('Ambiguous multiple records for '+key)
                    records[key]=json.loads(item['record_json'])
    duplicates=resolve_missing_duplicates(sqlite_db,source_root,keys,records)
    items=[]; problems=[]
    for key in sorted(keys):
        _validate_child(key,source_root)
        if key in duplicates:
            cat=duplicates[key]
        else:
            try: cat=_route_category(records[key])
            except ValueError as exc:problems.append('%s: %s'%(Path(key).name,exc));continue
        items.append((key,cat))
    if problems:
        raise RuntimeError('%s files need special routing data absent from the saved scan. Preview blocked; no packages rescanned. Example: %s'%(len(problems),problems[0]))
    plan=_plan(items,source_root,destination_root)
    stamp=datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    session=Path(preview_root)/('Cached_Routing_Preview_'+stamp)
    session.mkdir(parents=True,exist_ok=True)
    output=session/'Routing_Preview.txt'
    lines=[]
    print('ROUTING PREVIEW: SQLITE CACHED BATCH - NO PACKAGE RESCAN / NO THUMBNAIL EXTRACTION')
    print('ROUTING PREVIEW: DRY RUN - NO SOURCE FILES WILL BE MOVED')
    print('ROUTING PREVIEW SOURCE RECORDS:',len(plan))
    for path,cat,dest,collision in plan:
        note='DESTINATION EXISTS' if os.path.exists(dest) else ('READY (SOURCE FOLDER KEPT: NAME COLLISION)' if collision else 'READY')
        values=[cat,'No',path,dest,note]
        print('ROUTING PREVIEW ITEM\t'+'\t'.join(v.replace('\t',' ') for v in values))
        lines.append('\t'.join(values))
    output.write_text('\n'.join(lines)+'\n',encoding='utf-8')
    from routing_plan_guard import write_approved_plan
    write_approved_plan(output,sqlite_db,source_root,destination_root,plan)
    print('Routing preview saved:',output)
    print('ROUTING PREVIEW COUNT:',len(plan))
    print('Cached preview finished; SQLite unchanged; packages not rescanned.')
    return output
