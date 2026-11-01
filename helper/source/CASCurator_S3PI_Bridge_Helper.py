#!/usr/bin/env python3
# SPDX-License-Identifier: GPL-3.0-or-later
from __future__ import print_function

"""Separate S3PI package snapshot helper for CAS Curator.

This GPL helper is the only scanner component that imports clr/pythonnet/S3PI.
It opens one Sims 3 package, records the package resource index, direct
(decompressed) resource bytes, WrapperDealer parse success/failure, wrapper
bytes, and the small structured CASP/GEOM values the existing scanner uses.
It never accesses the CAS Curator MDB and never moves/routes source files.
"""

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback

SCHEMA = "CAS_CURATOR_S3PI_BRIDGE_SNAPSHOT_V1"
CASP_TYPE = 0x034AEECB
GEOM_TYPE = 0x015A1849
DLLS = (
    'System.Custom.dll', 's3pi.Interfaces.dll', 's3pi.Settings.dll', 's3pi.Package.dll', 's3pi.WrapperDealer.dll',
    's3pi.CASPartResource.dll', 's3pi.DefaultResource.dll',
    's3pi.GenericRCOLResource.dll', 's3pi.RCOLChunks.dll', 's3pi.RigResource.dll',
    's3pi.JazzResource.dll', 's3pi.ScriptResource.dll', 's3pi.MeshChunks.dll',
    's3pi.ModularResource.dll', 's3pi.NGMPHashMapResource.dll', 's3pi.CatalogResource.dll',
)


def sha256_file(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def locate_s3pi(folder):
    folder = Path(folder).expanduser().resolve(strict=True)
    missing = [name for name in DLLS if not (folder / name).is_file()]
    if missing:
        raise ValueError('Required S3PI DLLs missing: ' + ', '.join(missing))
    return folder


def load_s3pi(dll_dir):
    import clr
    sys.path.insert(0, str(dll_dir))
    for dll in DLLS:
        clr.AddReference(str(dll_dir / dll))
    from CASPartResource import CASPartResourceHandler
    CASPartResourceHandler()
    from s3pi.Package import Package
    from s3pi.WrapperDealer import WrapperDealer
    return Package, WrapperDealer


def resolve_get_resource_method(package):
    import System
    from System.Reflection import BindingFlags
    package_type = package.GetType()
    preferred = None
    fallback = None
    methods = package_type.GetMethods(
        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic
    )
    for method in methods:
        try:
            if method.Name != 'GetResource':
                continue
            params = method.GetParameters()
            if len(params) != 1:
                continue
            if fallback is None:
                fallback = method
            name = str(params[0].ParameterType.FullName or '')
            if 'IResourceIndexEntry' in name:
                preferred = method
                break
        except Exception:
            continue
    return preferred or fallback


def read_direct_resource(package, entry, method):
    import System
    if method is None:
        raise RuntimeError('Could not resolve package GetResource method')
    stream = None
    try:
        args = System.Array[System.Object]([entry])
        stream = method.Invoke(package, args)
        if stream is None or not stream.CanRead:
            raise RuntimeError('Package returned an unreadable resource stream')
        if stream.CanSeek:
            stream.Position = 0
        data = bytearray()
        buffer = System.Array.CreateInstance(System.Byte, 65536)
        while True:
            count = stream.Read(buffer, 0, buffer.Length)
            if count <= 0:
                break
            data.extend(int(buffer[i]) for i in range(count))
        return bytes(data)
    except Exception as exc:
        actual = getattr(exc, 'InnerException', None) or exc
        try:
            error_name = str(actual.GetType().Name)
        except Exception:
            error_name = type(actual).__name__
        raise RuntimeError('%s: %s' % (error_name, actual))
    finally:
        try:
            if stream is not None:
                stream.Dispose()
        except Exception:
            try:
                if stream is not None:
                    stream.Close()
            except Exception:
                pass


def safe_field(resource, field):
    try:
        return str(resource[field])
    except Exception:
        return ''


def snapshot_package(package_path, snapshot_dir, Package, WrapperDealer):
    package_path = Path(package_path).expanduser().resolve(strict=True)
    if not package_path.is_file() or package_path.suffix.lower() != '.package':
        raise ValueError('Expected one existing .package file')

    snapshot_dir = Path(snapshot_dir).expanduser().resolve()
    if snapshot_dir.exists():
        shutil.rmtree(str(snapshot_dir))
    (snapshot_dir / 'raw').mkdir(parents=True)
    (snapshot_dir / 'wrapper').mkdir(parents=True)

    before_hash = sha256_file(package_path)
    before_stat = package_path.stat()
    package = None
    close_error = None
    records = []

    try:
        package = Package.OpenPackage(0, str(package_path))
        resources = package.GetResourceList
        method = resolve_get_resource_method(package)

        for i in range(int(resources.Count)):
            entry = resources[i]
            rtype = int(entry.ResourceType) & 0xFFFFFFFF
            group = int(entry.ResourceGroup) & 0xFFFFFFFF
            instance = int(entry.Instance) & 0xFFFFFFFFFFFFFFFF
            try:
                filesize = int(entry.Filesize)
            except Exception:
                filesize = 0

            rec = {
                'index': i,
                'resource_type': rtype,
                'resource_group': group,
                'instance': instance,
                'filesize': filesize,
                'storage_ok': False,
                'storage_error': None,
                'raw_relpath': None,
                'wrapper_ok': False,
                'wrapper_error': None,
                'wrapper_asbytes_ok': False,
                'wrapper_asbytes_error': None,
                'wrapper_relpath': None,
                'wrapper_type': None,
                'casp_fields': {},
                'geom_chunks': [],
            }

            # Direct DBPF/decompression layer.
            try:
                raw = read_direct_resource(package, entry, method)
                raw_rel = 'raw/r_%06d.bin' % i
                (snapshot_dir / raw_rel).write_bytes(raw)
                rec['storage_ok'] = True
                rec['raw_relpath'] = raw_rel
            except Exception as exc:
                rec['storage_error'] = str(exc)

            # S3PI wrapper/parser layer.
            resource = None
            try:
                resource = WrapperDealer.GetResource(0, package, entry)
                rec['wrapper_type'] = type(resource).__name__
                rec['wrapper_ok'] = True
                try:
                    wrapper_raw = bytes(resource.AsBytes)
                    wrap_rel = 'wrapper/r_%06d.bin' % i
                    (snapshot_dir / wrap_rel).write_bytes(wrapper_raw)
                    rec['wrapper_asbytes_ok'] = True
                    rec['wrapper_relpath'] = wrap_rel
                except Exception as bytes_exc:
                    rec['wrapper_asbytes_error'] = '%s: %s' % (type(bytes_exc).__name__, bytes_exc)

                if rtype == CASP_TYPE:
                    for field in (
                        'Clothing', 'DataType', 'AgeGender', 'ClothingCategory',
                        'BodyType', 'OverlayPriority', 'BodyTypeFlags',
                        'CategoryFlags', 'CasPart1Index', 'CasPart2Index',
                        'LODInfo',
                    ):
                        rec['casp_fields'][field] = safe_field(resource, field)

                if rtype == GEOM_TYPE:
                    chunks = resource['ChunkEntries'].Value
                    for j in range(int(chunks.Count)):
                        block = chunks[j].RCOLBlock
                        vertices = 0
                        faces = 0
                        try:
                            if hasattr(block, 'VertexData'):
                                vertices = int(block.VertexData.Count)
                            if hasattr(block, 'Faces'):
                                faces = int(block.Faces.Count)
                        except Exception:
                            vertices = 0
                            faces = 0
                        rec['geom_chunks'].append({
                            'vertices': vertices,
                            'faces': faces,
                        })
            except Exception as exc:
                rec['wrapper_error'] = '%s: %s' % (type(exc).__name__, exc)
            finally:
                resource = None

            records.append(rec)
    finally:
        if package is not None:
            try:
                Package.ClosePackage(0, package)
            except Exception as exc:
                close_error = '%s: %s' % (type(exc).__name__, exc)

    after_stat = package_path.stat()
    after_hash = sha256_file(package_path)
    if before_hash != after_hash:
        raise RuntimeError('Source package hash changed while helper was reading it')
    if (before_stat.st_size, before_stat.st_mtime_ns) != (after_stat.st_size, after_stat.st_mtime_ns):
        raise RuntimeError('Source package metadata changed while helper was reading it')
    if close_error:
        raise RuntimeError('Package close failed: ' + close_error)

    payload = {
        'schema': SCHEMA,
        'helper_process': True,
        'package_path': str(package_path),
        'package_sha256': before_hash,
        'resource_count': len(records),
        'resources': records,
    }
    (snapshot_dir / 'manifest.json').write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + '\n',
        encoding='utf-8'
    )
    return payload


def main(argv=None):
    ap = argparse.ArgumentParser(description='CAS Curator separate S3PI bridge snapshot helper')
    ap.add_argument('--s3pi', required=True)
    ap.add_argument('--package', required=True)
    ap.add_argument('--snapshot-dir', required=True)
    args = ap.parse_args(argv)

    snapshot_dir = Path(args.snapshot_dir).expanduser().resolve()
    try:
        dll_dir = locate_s3pi(args.s3pi)
        Package, WrapperDealer = load_s3pi(dll_dir)
        payload = snapshot_package(args.package, snapshot_dir, Package, WrapperDealer)
        print('BRIDGE HELPER PASS: resources=%d snapshot=%s' % (
            payload['resource_count'], snapshot_dir
        ))
        return 0
    except Exception as exc:
        try:
            snapshot_dir.mkdir(parents=True, exist_ok=True)
            (snapshot_dir / 'helper_error.json').write_text(
                json.dumps({
                    'schema': SCHEMA,
                    'helper_process': True,
                    'error': '%s: %s' % (type(exc).__name__, exc),
                    'traceback_tail': traceback.format_exc().splitlines()[-10:],
                }, ensure_ascii=False, indent=2) + '\n',
                encoding='utf-8'
            )
        except Exception:
            pass
        print('BRIDGE HELPER FAILED: %s: %s' % (type(exc).__name__, exc))
        return 1


if __name__ == '__main__':
    sys.exit(main())
