from pathlib import Path
import re, sys

src = Path(sys.argv[1])
out = Path(sys.argv[2])
text = src.read_text(encoding='utf-8')

# Parent runtime must never import pythonnet/CLR/S3PI.
text = re.sub(r'(?m)^import clr\n', '', text, count=1)

start = text.index('S3PE_PATH = r"G:\\Game Tools\\Sims\\Sims 3 tools\\s3pe"')
end_marker = 'from s3pi.WrapperDealer import *\n'
end = text.index(end_marker, start) + len(end_marker)
replacement = r'''# -------------------------
# SEPARATE S3PI HELPER BOUNDARY (compiled public runtime)
# -------------------------
# The proprietary scanner parent never imports clr/System/s3pi. One separate
# GPL helper process opens each package and writes a read-only snapshot. The
# compatibility bridge exposes only the small Package/WrapperDealer surface
# already used by the validated scanner.
from cas_curator_helper_bridge import Package, WrapperDealer, _assert_boundary
_assert_boundary("scanner startup")
print("S3PI BOUNDARY: separate helper process / parent has no clr or s3pi imports")

_runtime_dir = os.path.dirname(
    os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__)
)
EA_DATABASE = os.path.join(
    _runtime_dir,
    "EA_Reference",
    "EA_resources_HASH_TEST_v5_ALL_PACKS.json",
)

with open(EA_DATABASE, "r", encoding="utf-8") as f:
    EA_RESOURCES = json.load(f)

EA_LOOKUP = set()
EA_HASH_LOOKUP = {}
EA_UNVERIFIABLE_TGIS = set()
EA_REFERENCE_INVALID_RECORDS = 0

for r in EA_RESOURCES:
    key = (
        int(r["type"], 16),
        int(r["group"], 16),
        int(r["instance"], 16)
    )

    EA_LOOKUP.add(key)

    hash_method = str(r.get("hash_method", "raw") or "raw").strip().lower()
    digest = str(r.get("sha256", "") or "").strip().lower()

    if hash_method == "unverifiable":
        EA_UNVERIFIABLE_TGIS.add(key)
        continue

    if hash_method not in ("raw", "wrapper") or len(digest) != 64:
        EA_REFERENCE_INVALID_RECORDS += 1
        continue

    EA_HASH_LOOKUP.setdefault(key, {}).setdefault(hash_method, set()).add(digest)

if EA_REFERENCE_INVALID_RECORDS:
    raise RuntimeError(
        "EA hash reference contains %s invalid hash records"
        % EA_REFERENCE_INVALID_RECORDS
    )

print(
    "EA HASH REFERENCE: TGIs=%s | exact variants=%s | unverifiable TGIs=%s"
    % (
        format(len(EA_LOOKUP), ","),
        format(
            sum(
                len(hashes)
                for methods in EA_HASH_LOOKUP.values()
                for hashes in methods.values()
            ),
            ","
        ),
        format(len(EA_UNVERIFIABLE_TGIS), ","),
    )
)
'''
# Replacement re-emits EA block, so consume through old WrapperDealer imports.
text = text[:start] + replacement + text[end:]

# Replace raw package read reflection with helper snapshot bytes.
pat = re.compile(r'def _get_raw_resource_bytes\(package, entry\):.*?\n\ndef _get_resource_sha256\(package, entry\):', re.S)
m = pat.search(text)
if not m:
    raise SystemExit('Could not locate _get_raw_resource_bytes block')
raw_repl = '''def _get_raw_resource_bytes(package, entry):\n    """Return direct/decompressed resource bytes captured by the separate helper."""\n    if not hasattr(package, "_bridge_raw_bytes"):\n        raise RuntimeError("Expected a helper-backed package proxy")\n    return package._bridge_raw_bytes(entry)\n\n\ndef _get_resource_sha256(package, entry):'''
text = text[:m.start()] + raw_repl + text[m.end():]

# Replace CLR-reflection package storage test with helper-produced storage result.
pat = re.compile(r'def check_package_storage_health\(package, resources, filename\):.*?\n\ndef _health_issue_instance_and_text\(raw_error\):', re.S)
m = pat.search(text)
if not m:
    raise SystemExit('Could not locate check_package_storage_health block')
storage_repl = '''def check_package_storage_health(package, resources, filename):\n    """Use package-storage/decompression results produced inside the separate helper."""\n    failures = []\n    failed_ids = []\n    if not hasattr(package, "_bridge_storage_error"):\n        print("PACKAGE STORAGE CHECK UNAVAILABLE: expected a helper-backed package proxy")\n        return failures, failed_ids\n\n    for i in range(resources.Count):\n        entry = resources[i]\n        try:\n            error_text = package._bridge_storage_error(entry)\n        except Exception as exc:\n            error_text = "%s: %s" % (type(exc).__name__, exc)\n        if not error_text:\n            continue\n\n        rid = f"0x{int(entry.Instance):016X}"\n        rtype = f"0x{int(entry.ResourceType):08X}"\n        id_text = f"{rtype}/{rid}"\n        reason = str(error_text)\n        failures.append(f"PACKAGE STORAGE {id_text}: {reason}")\n        failed_ids.append(id_text)\n        print(\n            f"PACKAGE STORAGE HEALTH FAILURE: "\n            f"Type={rtype} | Instance={rid} | {reason}"\n        )\n\n    if failures:\n        print(\n            f"PACKAGE STORAGE HEALTH SUMMARY: "\n            f"{os.path.basename(filename)} | "\n            f"BROKEN_ENTRIES={len(failures)} | "\n            f"TOTAL={resources.Count}"\n        )\n    return failures, failed_ids\n\n\ndef _health_issue_instance_and_text(raw_error):'''
text = text[:m.start()] + storage_repl + text[m.end():]

# Parent scanner has no CLR GC. Child helper exits after each snapshot, so plain
# Python GC is sufficient before file moves.
pat = re.compile(r'if MOVE_FILES:\n\s*import clr\n\s*from System import GC\n\s*GC\.Collect\(\)\n\s*GC\.WaitForPendingFinalizers\(\)\n\s*GC\.Collect\(\)\n')
m = pat.search(text)
if not m:
    raise SystemExit('Could not locate CLR GC block')
text = text[:m.start()] + 'if MOVE_FILES:\n\n    import gc\n    gc.collect()\n' + text[m.end():]

# Add immutable build note without changing scanner logic.
text = text.replace(
    'PRIMARY_BUILD_MARKER = (\n    "BUILD: DEV/PUBLIC SCRIPTCLASS TYPEDEF MATCH PERMANENT - 2026-09-12"\n)\n',
    'PRIMARY_BUILD_MARKER = (\n    "BUILD: DEV/PUBLIC SCRIPTCLASS TYPEDEF MATCH PERMANENT - 2026-09-12"\n)\nHELPER_BOUNDARY_BUILD = "CAS Curator v1.1.0 separate S3PI helper boundary"\n',
    1,
)

out.write_text(text, encoding='utf-8', newline='\n')
print(out)
