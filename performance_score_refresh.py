from __future__ import print_function

import argparse
import hashlib
import json
import os
import re
import sys
from pathlib import Path

IMG_TYPE = 0x00B2D882
GEOM_TYPE = 0x015A1849

PERFORMANCE_REFRESH_FIELDS = (
    "Risk",
    "Risk Score",
    "HQ Warning",
    "Poly Warning",
    "Biggest Texture",
)


def _as_int(value, default=0):
    try:
        text = str(value if value is not None else "").strip().replace(",", "")
        if not text:
            return int(default)
        return int(float(text))
    except Exception:
        return int(default)


def get_poly_warning(poly_count, cc_type, mixed_casp_types=None):
    """Return the current recalibrated poly warning without importing scanner core."""
    poly_count = _as_int(poly_count, 0)
    cc_type = str(cc_type or "").strip()

    if cc_type == "Mixed CAS" and mixed_casp_types:
        priority = {"Normal": 0, "HEAVY": 1, "VERY HEAVY": 2}
        warnings = [get_poly_warning(poly_count, item) for item in mixed_casp_types if str(item or "").strip()]
        if warnings:
            return max(warnings, key=lambda warning: priority.get(warning, 0))

    if cc_type in ("Shoes", "Top", "Bottom"):
        if poly_count >= 12000:
            return "VERY HEAVY"
        if poly_count >= 7000:
            return "HEAVY"
        return "Normal"

    if cc_type == "Body":
        if poly_count >= 20000:
            return "VERY HEAVY"
        if poly_count >= 13000:
            return "HEAVY"
        return "Normal"

    if cc_type == "Hair":
        if poly_count >= 40000:
            return "VERY HEAVY"
        if poly_count >= 30000:
            return "HEAVY"
        return "Normal"

    if cc_type in (
        "Accessory", "Earrings", "Glasses_F", "Ring_Lt", "Bracelets",
        "Glove", "Socks", "ArmBand", "Beard", "Moles", "Tattoo",
        "Makeup", "Eyebrow", "Lipstick", "Blush", "Eyeshadow", "Eyeliner",
    ):
        if poly_count >= 12000:
            return "VERY HEAVY"
        if poly_count >= 7000:
            return "HEAVY"
        return "Normal"

    if cc_type in ("Skin", "Custom Sim", "Morph Slider"):
        return "Normal"

    if poly_count >= 30000:
        return "VERY HEAVY"
    if poly_count >= 15000:
        return "HEAVY"
    return "Normal"


def calculate_performance_score(hq_warning, poly_warning, geom_bytes, fourk, cc_type):
    """Current recalibrated Performance Risk Score."""
    if str(cc_type or "").strip() == "Custom Sim":
        return "0/10", "Safe"

    score = 0
    hq_warning = str(hq_warning or "")
    poly_warning = str(poly_warning or "")
    geom_bytes = _as_int(geom_bytes, 0)
    fourk = _as_int(fourk, 0)

    if hq_warning == "4K TEXTURE":
        score += 5
        if fourk >= 3:
            score += 1
    elif hq_warning == "2048 HIGH":
        score += 1

    if poly_warning == "VERY HEAVY":
        score += 5
    elif poly_warning == "HEAVY":
        score += 3

    geom_mb = geom_bytes / float(1024 * 1024)
    if geom_mb >= 2:
        score += 2
    elif geom_mb >= 1:
        score += 1

    score = min(score, 10)
    if score <= 1:
        level = "Safe"
    elif score <= 3:
        level = "Moderate"
    elif score <= 5:
        level = "High"
    else:
        level = "Extreme"
    return "%d/10" % score, level


def score_fields_from_measurements(item, geom_bytes, fourk, two_k, biggest_texture):
    cc_type = str(item.get("cc_type", "") or "").strip()
    highest_polys = _as_int(item.get("highest_polys", 0), 0)
    mixed_types = item.get("mixed_casp_types") or []

    if _as_int(fourk, 0) > 0:
        hq_warning = "4K TEXTURE"
    elif _as_int(two_k, 0) > 0:
        hq_warning = "2048 HIGH"
    else:
        hq_warning = "Normal"

    poly_warning = get_poly_warning(highest_polys, cc_type, mixed_types)
    risk_score, risk = calculate_performance_score(
        hq_warning, poly_warning, geom_bytes, fourk, cc_type
    )
    return {
        "Risk": risk,
        "Risk Score": risk_score,
        "HQ Warning": hq_warning,
        "Poly Warning": poly_warning,
        "Biggest Texture": str(biggest_texture or ""),
    }


def measurement_fields_from_measurements(measured):
    """MDB fields that preserve the raw inputs needed by future score models."""
    return {
        "Texture Resolutions": str(measured.get("texture_resolutions", "") or ""),
        "GEOM Total Bytes": str(_as_int(measured.get("geom_bytes", 0), 0)),
        "4K Texture Count": str(_as_int(measured.get("fourk", 0), 0)),
        "Performance Data Version": "1",
    }


def _package_key(value):
    return Path(str(value or "").strip()).name.casefold()


def _texture_dimensions_from_text(value):
    text = str(value or "")
    found = []
    for width, height in re.findall(r"(\d{2,5})\s*[xX×]\s*(\d{2,5})", text):
        try:
            found.append((int(width), int(height)))
        except Exception:
            continue
    return found


def _normalise_texture_text(value):
    dims = _texture_dimensions_from_text(value)
    if not dims:
        return str(value or "").strip().casefold()
    width, height = max(dims, key=lambda pair: pair[0] * pair[1])
    return "%dx%d" % (width, height)


def _row_dict(headers, row):
    return {str(headers[index] or "").strip(): row[index] if index < len(row) else None for index in range(len(headers))}


def _read_sheet_records(ws):
    rows = ws.iter_rows(values_only=True)
    try:
        headers = list(next(rows))
    except StopIteration:
        return []
    return [_row_dict(headers, list(row)) for row in rows]


def load_health_scan_report(path):
    """Load the exact performance measurements preserved in Health_Scan_Report.xlsx."""
    from openpyxl import load_workbook

    report_path = Path(path)
    if not report_path.is_file():
        raise FileNotFoundError("Health Scan Report not found: %s" % report_path)
    wb = load_workbook(str(report_path), read_only=True, data_only=True)
    try:
        if "CC Summary" not in wb.sheetnames or "Technical Details" not in wb.sheetnames:
            raise ValueError("This workbook is not a compatible Health_Scan_Report.xlsx (CC Summary / Technical Details missing).")
        summary_records = _read_sheet_records(wb["CC Summary"])
        tech_records = _read_sheet_records(wb["Technical Details"])
    finally:
        wb.close()

    summary_index = {}
    for record in summary_records:
        key = _package_key(record.get("Package"))
        if key:
            summary_index.setdefault(key, []).append(record)
    tech_index = {}
    for record in tech_records:
        key = _package_key(record.get("Package"))
        if key:
            tech_index.setdefault(key, []).append(record)

    return {
        "path": str(report_path.resolve()),
        "summary": summary_index,
        "technical": tech_index,
    }


def _unique_report_values(records, fields, label, package_name):
    if not records:
        raise KeyError("%s has no %s row in the Health Scan Report." % (package_name, label))
    signatures = {}
    for record in records:
        signature = tuple(str(record.get(field, "") if record.get(field, "") is not None else "").strip() for field in fields)
        signatures.setdefault(signature, record)
    if len(signatures) != 1:
        raise ValueError(
            "The Health Scan Report contains more than one different %s row for '%s'; CAS Curator will not guess."
            % (label, package_name)
        )
    return next(iter(signatures.values()))


def measurements_from_mdb(item):
    """Return complete stored performance measurements for current MDB rows."""
    if str(item.get("performance_data_version", "") or "").strip() != "1":
        raise ValueError("MDB row does not contain current stored performance measurements.")
    data = item.get("mdb_measurements") or {}
    biggest_texture = str(data.get("biggest_texture", "") or "").strip()
    resolutions = str(data.get("texture_resolutions", "") or "")
    dimensions = _texture_dimensions_from_text(resolutions)
    if not dimensions:
        dimensions = _texture_dimensions_from_text(biggest_texture)
    two_k = sum(1 for width, height in dimensions if max(width, height) >= 2048 and max(width, height) < 4096)
    return {
        "geom_bytes": _as_int(data.get("geom_total_bytes", 0), 0),
        "fourk": _as_int(data.get("fourk_texture_count", 0), 0),
        "two_k": two_k,
        "biggest_texture": biggest_texture,
        "texture_resolutions": resolutions,
        "report_highest_poly": 0,
        "report_cc_type": "",
    }


def measurements_from_health_scan_report(item, report_index):
    package_name = Path(str(item.get("package", "") or "").strip()).name
    key = package_name.casefold()
    summary = _unique_report_values(
        report_index.get("summary", {}).get(key, []),
        ("CC Type", "Highest Poly Count", "Biggest Texture"),
        "CC Summary",
        package_name,
    )
    technical = _unique_report_values(
        report_index.get("technical", {}).get(key, []),
        ("All Resolutions", "GEOM Total Bytes", "4K Count"),
        "Technical Details",
        package_name,
    )

    # Guard filename-only report matching with values already stored in the MDB.
    stored_poly = _as_int(item.get("highest_polys", 0), 0)
    report_poly = _as_int(summary.get("Highest Poly Count", 0), 0)
    if stored_poly and report_poly and stored_poly != report_poly:
        raise ValueError("Health Scan Report Highest Poly does not match the MDB row.")

    stored_texture = _normalise_texture_text((item.get("old") or {}).get("Biggest Texture", ""))
    report_texture = _normalise_texture_text(summary.get("Biggest Texture", ""))
    if stored_texture and report_texture and stored_texture != report_texture:
        raise ValueError("Health Scan Report Biggest Texture does not match the MDB row.")

    biggest_texture = str(summary.get("Biggest Texture", "") or "").strip()
    fourk = _as_int(technical.get("4K Count", 0), 0)
    all_resolutions = str(technical.get("All Resolutions", "") or "")
    dimensions = _texture_dimensions_from_text(all_resolutions)
    if not dimensions:
        dimensions = _texture_dimensions_from_text(biggest_texture)
    two_k = 1 if any(max(width, height) >= 2048 and max(width, height) < 4096 for width, height in dimensions) else 0

    return {
        "geom_bytes": _as_int(technical.get("GEOM Total Bytes", 0), 0),
        "fourk": fourk,
        "two_k": two_k,
        "biggest_texture": biggest_texture,
        "texture_resolutions": all_resolutions,
        "report_highest_poly": report_poly,
        "report_cc_type": str(summary.get("CC Type", "") or "").strip(),
    }


def result_from_measurements(item, measured, source, source_path=""):
    old = {
        field: str(item.get("old", {}).get(field, "") or "")
        for field in PERFORMANCE_REFRESH_FIELDS
    }
    scoring_item = dict(item)
    if not _as_int(scoring_item.get("highest_polys", 0), 0) and _as_int(measured.get("report_highest_poly", 0), 0):
        scoring_item["highest_polys"] = measured.get("report_highest_poly", 0)
    new_fields = score_fields_from_measurements(
        scoring_item,
        measured["geom_bytes"],
        measured["fourk"],
        measured["two_k"],
        measured["biggest_texture"],
    )
    changed = any(old.get(field, "") != new_fields.get(field, "") for field in PERFORMANCE_REFRESH_FIELDS)
    measurement_fields = measurement_fields_from_measurements(measured)
    backfill_required = str(item.get("performance_data_version", "") or "").strip() != "1"
    return {
        "excel_row": int(item.get("excel_row", 0) or 0),
        "package": str(item.get("package", "") or ""),
        "package_hash": str(item.get("package_hash", "") or ""),
        "status": "changed" if changed else "unchanged",
        "path": str(source_path or ""),
        "source": str(source or ""),
        "old": old,
        "new": new_fields,
        "measurements": measured,
        "measurement_fields": measurement_fields,
        "backfill_required": backfill_required,
        "error": "",
    }


def _sha256_file(path):
    h = hashlib.sha256()
    with open(str(path), "rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _candidate_paths(item, managed_index=None, fallback_index=None):
    """Return possible physical package paths in safest source order.

    Priority is intentional:
      0 current tracked File Location (exact filename)
      1 current tracked File Location (recursive recovery inside that location)
      2 any other configured CAS Curator managed File Location
      3 stored Original Source Path
      4 user-chosen fallback/search folder
    Every candidate is still SHA-256 verified by resolve_package_path when a
    stored Package Hash is available.
    """
    filename = Path(str(item.get("package", "") or "").strip()).name
    if not filename.lower().endswith(".package"):
        return []
    candidates = []

    location_path = str(item.get("file_location_path", "") or "").strip()
    if location_path:
        folder = Path(location_path)
        if folder.is_dir():
            direct = folder / filename
            if direct.is_file():
                candidates.append((0, direct, "Current File Location"))
            else:
                try:
                    for found in folder.rglob(filename):
                        if found.is_file():
                            candidates.append((1, found, "Current File Location"))
                            if len(candidates) >= 80:
                                break
                except OSError:
                    pass

    if managed_index:
        for candidate in managed_index.get(filename.casefold(), []):
            if candidate.is_file():
                candidates.append((2, candidate, "Managed File Location"))

    for raw in item.get("original_source_paths") or []:
        text = str(raw or "").strip()
        if not text:
            continue
        candidate = Path(text)
        if candidate.is_dir():
            candidate = candidate / filename
        if candidate.is_file():
            candidates.append((3, candidate, "Original Source Path"))

    if fallback_index:
        for candidate in fallback_index.get(filename.casefold(), []):
            if candidate.is_file():
                candidates.append((4, candidate, "Chosen Folder"))

    unique = {}
    for priority, candidate, source in candidates:
        try:
            resolved = candidate.resolve()
        except OSError:
            resolved = candidate.absolute()
        key = os.path.normcase(str(resolved)).casefold()
        previous = unique.get(key)
        if previous is None or priority < previous[0]:
            unique[key] = (priority, resolved, source)
    return list(unique.values())


def resolve_package_path(item, managed_index=None, fallback_index=None):
    filename = Path(str(item.get("package", "") or "").strip()).name
    expected_hash = str(item.get("package_hash", "") or "").strip().lower()
    candidates = _candidate_paths(
        item,
        managed_index=managed_index,
        fallback_index=fallback_index,
    )
    if not candidates:
        raise FileNotFoundError("Could not locate the original package '%s'." % filename)

    checked = []
    for priority, candidate, source in candidates:
        try:
            actual_hash = _sha256_file(candidate)
        except OSError:
            continue
        checked.append((priority, candidate, actual_hash, source))

    if expected_hash:
        matching = [entry for entry in checked if entry[2].lower() == expected_hash]
        if not matching:
            raise ValueError("Located file(s), but none match the Package Hash stored in the MDB.")
    else:
        matching = checked

    if not matching:
        raise FileNotFoundError("Could not read a usable source file for '%s'." % filename)
    best_priority = min(entry[0] for entry in matching)
    best = [entry for entry in matching if entry[0] == best_priority]
    if len(best) != 1:
        raise ValueError("More than one equally valid copy was found; CAS Curator will not guess.")
    return best[0][1], best[0][3]


def _find_s3pi_dir():
    candidates = []
    env_path = str(os.environ.get("CAS_CURATOR_S3PE_PATH", "") or "").strip()
    if env_path:
        candidates.append(Path(env_path))

    # Source-project location used by the validated scanner core.
    candidates.append(Path(r"G:\Game Tools\Sims\Sims 3 tools\s3pe"))

    # Bundled scanner-runtime candidates. PyInstaller exposes _MEIPASS while
    # source/portable layouts commonly keep scanner DLLs beside the executable.
    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        candidates.append(Path(meipass))
        candidates.append(Path(meipass) / "s3pe")
        candidates.append(Path(meipass) / "scanner")
    try:
        exe_dir = Path(sys.executable).resolve().parent
        candidates.extend([exe_dir, exe_dir / "s3pe", exe_dir / "scanner"])
    except Exception:
        pass
    script_dir = Path(__file__).resolve().parent
    candidates.extend([script_dir, script_dir / "s3pe", script_dir / "scanner"])

    needed = ("s3pi.Interfaces.dll", "s3pi.Package.dll", "s3pi.WrapperDealer.dll")
    seen = set()
    for candidate in candidates:
        key = os.path.normcase(str(candidate)).casefold()
        if key in seen:
            continue
        seen.add(key)
        if candidate.is_dir() and all((candidate / name).is_file() for name in needed):
            return candidate
    return None


def load_s3pi():
    import clr

    s3pi_dir = _find_s3pi_dir()
    assemblies = ("s3pi.Interfaces", "s3pi.Package", "s3pi.WrapperDealer", "s3pi.DefaultResource")
    if s3pi_dir is not None:
        if str(s3pi_dir) not in sys.path:
            sys.path.append(str(s3pi_dir))
        for name in assemblies:
            dll = s3pi_dir / (name + ".dll")
            if dll.is_file():
                clr.AddReference(str(dll))
    else:
        # Frozen runtimes may already expose the assemblies by name.
        for name in assemblies:
            clr.AddReference(name)

    from s3pi.Package import Package
    from s3pi.WrapperDealer import WrapperDealer
    return Package, WrapperDealer


def _dds_dimensions(raw):
    if raw is None or len(raw) < 128:
        return None
    if raw[0:4] != b"DDS ":
        return None
    header_size = int.from_bytes(raw[4:8], "little") if hasattr(int, "from_bytes") else 0
    if header_size != 124:
        return None
    height = int.from_bytes(raw[12:16], "little")
    width = int.from_bytes(raw[16:20], "little")
    if width <= 0 or height <= 0:
        return None
    return width, height


def measure_package(path, Package, WrapperDealer):
    package = None
    resources_to_keep_alive = []
    geom_bytes = 0
    fourk = 0
    two_k = 0
    biggest_area = 0
    biggest_texture = ""
    resolutions = {}
    try:
        package = Package.OpenPackage(0, str(path))
        resources = package.GetResourceList
        for index in range(resources.Count):
            entry = resources[index]
            rtype = int(entry.ResourceType)
            if rtype == GEOM_TYPE:
                try:
                    geom_bytes += max(0, int(entry.Filesize))
                except Exception:
                    pass
                continue
            if rtype != IMG_TYPE:
                continue
            resource = None
            try:
                resource = WrapperDealer.GetResource(0, package, entry)
                resources_to_keep_alive.append(resource)
                raw = bytes(resource.AsBytes)
                size = _dds_dimensions(raw)
            except Exception:
                size = None
            if not size:
                continue
            width, height = size
            text = "%dx%d" % (width, height)
            resolutions[text] = resolutions.get(text, 0) + 1
            area = int(width) * int(height)
            if area > biggest_area:
                biggest_area = area
                biggest_texture = "%dx%d" % (width, height)
            if width >= 4096 or height >= 4096:
                fourk += 1
            elif width >= 2048 or height >= 2048:
                two_k += 1
        return {
            "geom_bytes": geom_bytes,
            "fourk": fourk,
            "two_k": two_k,
            "biggest_texture": biggest_texture,
            "texture_resolutions": "; ".join(
                "%s (%d)" % (key, value) for key, value in resolutions.items()
            ),
        }
    finally:
        if package is not None:
            try:
                Package.ClosePackage(0, package)
            except Exception:
                pass
        resources_to_keep_alive[:] = []


def refresh_item(item, Package, WrapperDealer, managed_index=None, fallback_index=None):
    old = {
        field: str(item.get("old", {}).get(field, "") or "")
        for field in PERFORMANCE_REFRESH_FIELDS
    }
    try:
        path, source = resolve_package_path(
            item, managed_index=managed_index, fallback_index=fallback_index
        )
        measured = measure_package(path, Package, WrapperDealer)
        return result_from_measurements(item, measured, source, str(path))
    except Exception as exc:
        return {
            "excel_row": int(item.get("excel_row", 0) or 0),
            "package": str(item.get("package", "") or ""),
            "package_hash": str(item.get("package_hash", "") or ""),
            "status": "skipped",
            "path": "",
            "source": "Physical package",
            "old": old,
            "new": {},
            "measurements": {},
            "measurement_fields": {},
            "backfill_required": False,
            "error": "%s: %s" % (type(exc).__name__, exc),
        }


def build_fallback_index(search_roots):
    """Index package filenames once so a large CC tree is not recursively searched per MDB row."""
    index = {}
    seen_roots = set()
    for raw in search_roots or []:
        raw = str(raw or "").strip()
        if not raw:
            continue
        root = Path(raw)
        try:
            root = root.resolve()
        except OSError:
            root = root.absolute()
        key = os.path.normcase(str(root)).casefold()
        if key in seen_roots or not root.is_dir():
            continue
        seen_roots.add(key)
        try:
            for path in root.rglob("*.package"):
                if not path.is_file():
                    continue
                index.setdefault(path.name.casefold(), []).append(path)
        except OSError:
            continue
    return index


def main(argv=None):
    parser = argparse.ArgumentParser(description="Refresh CAS Curator Performance Risk Scores without a full health rescan")
    parser.add_argument("--manifest", required=True)
    parser.add_argument("--result", required=True)
    args = parser.parse_args(argv)

    manifest_path = Path(args.manifest)
    result_path = Path(args.result)
    with manifest_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    items = list(payload.get("items") or [])
    managed_roots = list(payload.get("managed_search_roots") or [])
    fallback_roots = list(payload.get("fallback_search_roots") or [])
    report_path = str(payload.get("health_scan_report_path") or "").strip()
    report_optional = bool(payload.get("health_scan_report_optional", False))

    report_index = None
    if report_path:
        try:
            report_index = load_health_scan_report(report_path)
            print("REPORT|LOADED|%s" % report_path.replace("|", "_"), flush=True)
        except Exception as exc:
            if report_optional:
                print(
                    "REPORT|SKIPPED|%s" % (
                        ("%s: %s" % (type(exc).__name__, exc)).replace("|", "_")
                    ),
                    flush=True,
                )
                report_index = None
            else:
                result = {
                    "fatal_error": "Could not read the selected Health Scan Report: %s: %s" % (type(exc).__name__, exc),
                    "results": [],
                }
                with result_path.open("w", encoding="utf-8") as handle:
                    json.dump(result, handle, ensure_ascii=False, indent=2)
                print("FATAL|" + result["fatal_error"], flush=True)
                return 2

    Package = None
    WrapperDealer = None
    package_reader_error = None
    managed_index = build_fallback_index(managed_roots)
    if managed_roots:
        print(
            "MANAGED_INDEX|%d|%d" % (
                len(managed_index),
                sum(len(paths) for paths in managed_index.values()),
            ),
            flush=True,
        )

    fallback_index = build_fallback_index(fallback_roots)
    if fallback_roots:
        print(
            "INDEX|%d|%d" % (
                len(fallback_index),
                sum(len(paths) for paths in fallback_index.values()),
            ),
            flush=True,
        )

    results = []
    total = len(items)
    for index, item in enumerate(items, start=1):
        package_name = str(item.get("package", "") or "")
        print("PROGRESS|%d|%d|%s" % (index, total, package_name.replace("|", "_")), flush=True)

        mdb_error = ""
        try:
            measured = measurements_from_mdb(item)
            results.append(result_from_measurements(item, measured, "Master Database", ""))
            continue
        except Exception as exc:
            mdb_error = "%s: %s" % (type(exc).__name__, exc)

        report_error = ""
        if report_index is not None:
            try:
                measured = measurements_from_health_scan_report(item, report_index)
                results.append(
                    result_from_measurements(
                        item, measured, "Health Scan Report", report_index.get("path", report_path)
                    )
                )
                continue
            except Exception as exc:
                report_error = "%s: %s" % (type(exc).__name__, exc)

        if Package is None and package_reader_error is None:
            try:
                Package, WrapperDealer = load_s3pi()
            except Exception as exc:
                package_reader_error = "Could not load the Sims 3 package reader: %s: %s" % (type(exc).__name__, exc)

        if Package is not None:
            physical = refresh_item(
                item, Package, WrapperDealer,
                managed_index=managed_index, fallback_index=fallback_index
            )
            if physical.get("status") == "skipped":
                prefix = []
                if mdb_error:
                    prefix.append("MDB: %s" % mdb_error)
                if report_error:
                    prefix.append("Report: %s" % report_error)
                if prefix:
                    prefix.append("Package fallback: %s" % physical.get("error", ""))
                    physical["error"] = " | ".join(prefix)
            results.append(physical)
        else:
            old = {field: str(item.get("old", {}).get(field, "") or "") for field in PERFORMANCE_REFRESH_FIELDS}
            error_parts = []
            if mdb_error:
                error_parts.append("MDB: " + mdb_error)
            if report_error:
                error_parts.append("Report: " + report_error)
            if package_reader_error:
                error_parts.append("Package fallback: " + package_reader_error)
            results.append({
                "excel_row": int(item.get("excel_row", 0) or 0),
                "package": package_name,
                "package_hash": str(item.get("package_hash", "") or ""),
                "status": "skipped",
                "path": "",
                "source": "",
                "old": old,
                "new": {},
                "measurements": {},
                "measurement_fields": {},
                "backfill_required": False,
                "error": " | ".join(error_parts) or "No usable refresh source was available.",
            })

    summary = {
        "total": total,
        "changed": sum(1 for item in results if item.get("status") == "changed"),
        "unchanged": sum(1 for item in results if item.get("status") == "unchanged"),
        "skipped": sum(1 for item in results if item.get("status") == "skipped"),
    }
    with result_path.open("w", encoding="utf-8") as handle:
        json.dump({"fatal_error": "", "summary": summary, "results": results}, handle, ensure_ascii=False, indent=2)
    print("DONE|{changed}|{unchanged}|{skipped}".format(**summary), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
