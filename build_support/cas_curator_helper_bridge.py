from __future__ import print_function

"""Parent-side compatibility bridge for CAS Curator's separate S3PI helper.

This module intentionally does NOT import clr, pythonnet, System, or s3pi.
The GPL S3PI parser runs only in CASCurator_S3PI_Bridge_Helper.exe. The helper
writes a read-only snapshot for one package; these small proxy classes expose
only the Package/WrapperDealer surface used by the validated scanner.
"""

import atexit
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

SCHEMA = "CAS_CURATOR_S3PI_BRIDGE_SNAPSHOT_V1"
_FORBIDDEN_PREFIXES = ("clr", "pythonnet", "System", "s3pi", "CASPartResource")
_OPEN_PACKAGES = []


def forbidden_loaded_modules():
    bad = []
    for name in sys.modules:
        if any(name == prefix or name.startswith(prefix + ".") for prefix in _FORBIDDEN_PREFIXES):
            bad.append(name)
    return sorted(bad)


def _assert_boundary(stage):
    bad = forbidden_loaded_modules()
    if bad:
        raise RuntimeError(
            "S3PI boundary failure at %s: parent process loaded %s"
            % (stage, ", ".join(bad))
        )


def _runtime_root():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _helper_paths():
    root = _runtime_root()
    helper_env = str(os.environ.get("CAS_CURATOR_S3PI_BRIDGE_HELPER", "") or "").strip()
    s3pi_env = str(os.environ.get("CAS_CURATOR_S3PI_BRIDGE_DLL_DIR", "") or "").strip()

    if getattr(sys, "frozen", False):
        helper_root = root.parent / "helper"
    else:
        # Build/source staging normally keeps helper beside the scanner folder;
        # fall back to a helper subfolder when running directly from project root.
        helper_root = root.parent / "helper"
        if not helper_root.is_dir():
            helper_root = root / "helper"

    candidate_exe = Path(helper_env).expanduser() if helper_env else helper_root / "CASCurator_S3PI_Bridge_Helper.exe"
    candidate_py = helper_root / "source" / "CASCurator_S3PI_Bridge_Helper.py"
    s3pi_dir = Path(s3pi_env).expanduser() if s3pi_env else helper_root / "s3pi"

    candidate_exe = candidate_exe.resolve()
    candidate_py = candidate_py.resolve()
    s3pi_dir = s3pi_dir.resolve()

    if not candidate_exe.is_file():
        # The compiled public scanner requires the separate helper executable.
        # A source path is returned only for diagnostics/build tests.
        if not candidate_py.is_file():
            raise RuntimeError("Separate S3PI bridge helper is missing: %s" % candidate_exe)
    if not s3pi_dir.is_dir():
        raise RuntimeError("Separate S3PI DLL folder is missing: %s" % s3pi_dir)
    return candidate_exe, candidate_py, s3pi_dir


def _safe_rmtree(path):
    try:
        if path and Path(path).exists():
            shutil.rmtree(str(path), ignore_errors=True)
    except Exception:
        pass


class _CountList(object):
    def __init__(self, values):
        self._values = list(values)

    @property
    def Count(self):
        return len(self._values)

    def __len__(self):
        return len(self._values)

    def __getitem__(self, index):
        return self._values[index]

    def __iter__(self):
        return iter(self._values)


class _EntryProxy(object):
    def __init__(self, package, rec, bridge_index):
        self._package = package
        self._rec = rec
        self._bridge_index = int(bridge_index)
        self.ResourceType = int(rec.get("resource_type", 0))
        self.ResourceGroup = int(rec.get("resource_group", 0))
        self.Instance = int(rec.get("instance", 0))
        self.Filesize = int(rec.get("filesize", 0) or 0)


class _StringFieldProxy(object):
    def __init__(self, text):
        self._text = str(text if text is not None else "")
        self.Value = self._text

    def __str__(self):
        return self._text

    def __repr__(self):
        return self._text


class _SimpleCountHolder(object):
    def __init__(self, count):
        self.Count = int(count or 0)


class _GeomBlockProxy(object):
    def __init__(self, item):
        item = item or {}
        self.VertexData = _SimpleCountHolder(item.get("vertices", 0))
        self.Faces = _SimpleCountHolder(item.get("faces", 0))


class _GeomChunkProxy(object):
    def __init__(self, item):
        self.RCOLBlock = _GeomBlockProxy(item)


class _ChunkFieldProxy(object):
    def __init__(self, chunks):
        self.Value = _CountList([_GeomChunkProxy(item) for item in (chunks or [])])


class _DummyStream(object):
    def Dispose(self):
        return None

    def Close(self):
        return None


class _TypeProxy(object):
    def __init__(self, name):
        self.Name = str(name or "BridgeResourceProxy")
        self.FullName = self.Name

    def __str__(self):
        return self.FullName

    def __repr__(self):
        return self.FullName


class _ResourceProxy(object):
    def __init__(self, package, rec):
        self._package = package
        self._rec = rec
        self._snapshot_dir = package._snapshot_dir
        self._wrapper_bytes = None
        self.Stream = _DummyStream()
        self.ContentFields = list((rec.get("casp_fields") or {}).keys())

    @property
    def AsBytes(self):
        if self._wrapper_bytes is not None:
            return self._wrapper_bytes
        if not self._rec.get("wrapper_asbytes_ok"):
            error = self._rec.get("wrapper_asbytes_error") or "S3PI wrapper AsBytes failed in separate helper"
            raise RuntimeError(str(error))
        rel = self._rec.get("wrapper_relpath")
        if not rel:
            raise RuntimeError("Bridge snapshot has no wrapper bytes")
        path = self._snapshot_dir / rel
        self._wrapper_bytes = path.read_bytes()
        return self._wrapper_bytes

    def __getitem__(self, field):
        if field == "ChunkEntries":
            return _ChunkFieldProxy(self._rec.get("geom_chunks") or [])
        fields = self._rec.get("casp_fields") or {}
        if field in fields:
            return _StringFieldProxy(fields.get(field, ""))
        raise KeyError(field)

    def GetType(self):
        return _TypeProxy(self._rec.get("wrapper_type") or "BridgeResourceProxy")

    def __repr__(self):
        return "<BridgeResourceProxy type=%s>" % (self._rec.get("wrapper_type") or "unknown")


class _PackageProxy(object):
    def __init__(self, source_path, snapshot_dir, payload, helper_stdout=""):
        self._source_path = Path(source_path).resolve()
        self._snapshot_dir = Path(snapshot_dir).resolve()
        self._payload = payload
        self._helper_stdout = helper_stdout
        self._closed = False
        self._records = list(payload.get("resources") or [])
        self._entries = _CountList([
            _EntryProxy(self, rec, i) for i, rec in enumerate(self._records)
        ])
        self.GetResourceList = self._entries

    def _record_for_entry(self, entry):
        try:
            idx = int(entry._bridge_index)
            if entry._package is not self:
                raise ValueError("Entry belongs to a different bridge package")
            return self._records[idx]
        except Exception:
            raise IndexError("Bridge resource index out of range")

    def _bridge_raw_bytes(self, entry):
        rec = self._record_for_entry(entry)
        if not rec.get("storage_ok"):
            raise RuntimeError(rec.get("storage_error") or "Bridge package storage read failed")
        rel = rec.get("raw_relpath")
        if not rel:
            raise RuntimeError("Bridge snapshot has no direct resource bytes")
        return (self._snapshot_dir / rel).read_bytes()

    def _bridge_storage_error(self, entry):
        rec = self._record_for_entry(entry)
        if rec.get("storage_ok"):
            return ""
        return str(rec.get("storage_error") or "Bridge package storage read failed")

    def _bridge_wrapper_resource(self, entry):
        rec = self._record_for_entry(entry)
        if not rec.get("wrapper_ok"):
            raise RuntimeError(rec.get("wrapper_error") or "S3PI wrapper parse failed in separate helper")
        return _ResourceProxy(self, rec)

    def close(self):
        if self._closed:
            return
        self._closed = True
        _safe_rmtree(self._snapshot_dir)


class Package(object):
    @staticmethod
    def OpenPackage(_api_version, filename):
        _assert_boundary("before helper launch")
        package_path = Path(filename).expanduser().resolve(strict=True)
        helper_exe, helper_py, s3pi_dir = _helper_paths()
        snapshot_dir = Path(tempfile.mkdtemp(prefix="cas_curator_s3pi_bridge_"))

        if helper_exe.is_file():
            command = [
                str(helper_exe),
                "--s3pi", str(s3pi_dir),
                "--package", str(package_path),
                "--snapshot-dir", str(snapshot_dir),
            ]
        else:
            # Build/source diagnostic fallback only. The public compiled runtime
            # always ships the separate helper EXE.
            command = [
                sys.executable, str(helper_py),
                "--s3pi", str(s3pi_dir),
                "--package", str(package_path),
                "--snapshot-dir", str(snapshot_dir),
            ]

        proc = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            universal_newlines=True,
            cwd=str(helper_exe.parent if helper_exe.is_file() else helper_py.parent),
        )
        helper_stdout = proc.stdout or ""
        if proc.returncode != 0:
            error_text = helper_stdout.strip()
            err_json = snapshot_dir / "helper_error.json"
            if err_json.is_file():
                try:
                    error_text = json.loads(err_json.read_text(encoding="utf-8")).get("error") or error_text
                except Exception:
                    pass
            _safe_rmtree(snapshot_dir)
            raise RuntimeError("Separate S3PI bridge helper failed: %s" % (error_text or "exit %s" % proc.returncode))

        manifest = snapshot_dir / "manifest.json"
        if not manifest.is_file():
            _safe_rmtree(snapshot_dir)
            raise RuntimeError("Separate S3PI helper did not produce manifest.json")
        payload = json.loads(manifest.read_text(encoding="utf-8"))
        if payload.get("schema") != SCHEMA or payload.get("helper_process") is not True:
            _safe_rmtree(snapshot_dir)
            raise RuntimeError("Unexpected S3PI bridge snapshot schema")
        resources = payload.get("resources") or []
        if int(payload.get("resource_count", -1)) != len(resources):
            _safe_rmtree(snapshot_dir)
            raise RuntimeError("S3PI bridge snapshot resource count mismatch")

        # Extra read-only integrity check of the package handed to the helper.
        expected_hash = str(payload.get("package_sha256") or "").strip().lower()
        if expected_hash:
            h = hashlib.sha256()
            with package_path.open("rb") as f:
                for block in iter(lambda: f.read(1024 * 1024), b""):
                    h.update(block)
            if h.hexdigest().lower() != expected_hash:
                _safe_rmtree(snapshot_dir)
                raise RuntimeError("Source package changed after separate S3PI helper read")

        _assert_boundary("after helper launch")
        package = _PackageProxy(package_path, snapshot_dir, payload, helper_stdout)
        _OPEN_PACKAGES.append(package)
        return package

    @staticmethod
    def ClosePackage(_api_version, package):
        try:
            package.close()
        finally:
            try:
                _OPEN_PACKAGES.remove(package)
            except ValueError:
                pass
        _assert_boundary("after package close")


class WrapperDealer(object):
    @staticmethod
    def GetResource(_api_version, package, entry):
        if not isinstance(package, _PackageProxy):
            raise ValueError("WrapperDealer proxy expected a helper-backed package")
        return package._bridge_wrapper_resource(entry)


def _cleanup_all():
    for package in list(_OPEN_PACKAGES):
        try:
            package.close()
        except Exception:
            pass
    _OPEN_PACKAGES[:] = []


atexit.register(_cleanup_all)
_assert_boundary("bridge import")
