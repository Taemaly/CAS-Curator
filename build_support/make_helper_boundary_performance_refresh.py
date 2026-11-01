from pathlib import Path
import re
import sys

if len(sys.argv) != 3:
    raise SystemExit("usage: make_helper_boundary_performance_refresh.py SOURCE OUTPUT")

src = Path(sys.argv[1])
out = Path(sys.argv[2])
text = src.read_text(encoding="utf-8")

pattern = re.compile(
    r"def _find_s3pi_dir\(\):.*?\n\ndef _dds_dimensions\(raw\):",
    re.DOTALL,
)
replacement = '''def load_s3pi():
    """Use the same separate S3PI helper boundary as the compiled scanner."""
    from cas_curator_helper_bridge import Package, WrapperDealer, _assert_boundary
    _assert_boundary("performance score refresh")
    return Package, WrapperDealer


def _dds_dimensions(raw):'''
new_text, count = pattern.subn(replacement, text, count=1)
if count != 1:
    raise SystemExit("Could not locate the expected direct S3PI fallback block; refusing to guess.")

# Safety checks: the build-only copy must not contain the old CLR/S3PI import path.
for forbidden in ("import clr", "clr.AddReference", "from s3pi.Package import Package", "from s3pi.WrapperDealer import WrapperDealer"):
    if forbidden in new_text:
        raise SystemExit(f"Forbidden direct S3PI/CLR code remained after patch: {forbidden}")
if "from cas_curator_helper_bridge import Package, WrapperDealer, _assert_boundary" not in new_text:
    raise SystemExit("Helper-boundary import was not written.")

out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(new_text, encoding="utf-8", newline="\n")
