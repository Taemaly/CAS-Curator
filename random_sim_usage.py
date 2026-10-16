from __future__ import annotations

"""Persistent DEV-only Random Sim usage/history links.

This module intentionally does not edit the Master Database or scanner-owned fields.
It stores a small sidecar registry keyed by stable package identity so Library records
can show which exported Random Sim collections use the original CC package.
"""

from pathlib import Path
from typing import Dict, List, Mapping, MutableMapping, Sequence
import hashlib
import json
import os
import tempfile

from random_sim_exporter import package_identity

REGISTRY_FORMAT = "CAS Curator Random Sim Usage Registry v1"


def default_usage_registry_path() -> Path:
    """Return a stable per-user DEV registry path without touching the MDB."""
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or Path.home())
        return base / "CASCurator" / "DEV" / "random_sim_usage.json"
    base = Path(os.environ.get("XDG_DATA_HOME") or (Path.home() / ".local" / "share"))
    return base / "CASCurator" / "DEV" / "random_sim_usage.json"


def empty_usage_registry() -> Dict[str, object]:
    return {"format": REGISTRY_FORMAT, "sims": {}, "packages": {}}


def _normalise_registry(data: object) -> Dict[str, object]:
    if not isinstance(data, dict):
        return empty_usage_registry()
    sims = data.get("sims") if isinstance(data.get("sims"), dict) else {}
    packages = data.get("packages") if isinstance(data.get("packages"), dict) else {}
    return {"format": REGISTRY_FORMAT, "sims": dict(sims), "packages": dict(packages)}


def load_usage_registry(path: Path | str | None = None) -> Dict[str, object]:
    target = Path(path) if path is not None else default_usage_registry_path()
    if not target.is_file():
        return empty_usage_registry()
    try:
        return _normalise_registry(json.loads(target.read_text(encoding="utf-8")))
    except Exception:
        # Corrupt experimental metadata must never block CAS Curator or MDB access.
        return empty_usage_registry()


def save_usage_registry(registry: Mapping[str, object], path: Path | str | None = None) -> Path:
    """Atomically save the DEV sidecar usage registry."""
    target = Path(path) if path is not None else default_usage_registry_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = _normalise_registry(dict(registry))
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    fd, tmp_name = tempfile.mkstemp(prefix=target.name + ".", suffix=".tmp", dir=str(target.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
        os.replace(tmp_name, target)
    except Exception:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
        raise
    return target


def _sim_id(manifest: Mapping[str, object], export_folder: Path | str) -> str:
    seed = "\n".join(
        [
            str(manifest.get("set_name", "")),
            str(manifest.get("created_at", "")),
            str(Path(export_folder).resolve()),
        ]
    )
    return hashlib.sha256(seed.encode("utf-8", errors="replace")).hexdigest()[:20]


def register_export(
    registry: Mapping[str, object],
    manifest: Mapping[str, object],
    export_folder: Path | str,
) -> Dict[str, object]:
    """Return a registry updated with one successfully exported Random Sim collection.

    Both directly selected CC and explicit dependency-only packages are indexed. UI
    callers can distinguish roles. Existing Status (including "Used in Game") is not
    read, replaced, or modified here.
    """
    updated = _normalise_registry(dict(registry))
    sims: MutableMapping[str, object] = updated["sims"]  # type: ignore[assignment]
    packages: MutableMapping[str, object] = updated["packages"]  # type: ignore[assignment]

    folder = Path(export_folder).resolve()
    sim_id = _sim_id(manifest, folder)
    sim = {
        "sim_id": sim_id,
        "name": str(manifest.get("set_name", "") or folder.name),
        "created_at": str(manifest.get("created_at", "")),
        "folder": str(folder),
        "manifest": str(folder / "CAS_Curator_Collection_Manifest.json"),
        "master_database": str(manifest.get("master_database", "")),
        "app_version": str(manifest.get("app_version", "")),
        "selected_piece_count": int(manifest.get("selected_piece_count", 0) or 0),
        "unique_selected_package_count": int(manifest.get("unique_selected_package_count", 0) or 0),
        "dependency_package_count": int(manifest.get("dependency_package_count", 0) or 0),
        "generation_settings": dict(manifest.get("generation_settings", {}) or {}),
    }
    sims[sim_id] = sim

    for item in manifest.get("items", []) or []:
        if not isinstance(item, Mapping):
            continue
        identity = str(item.get("identity", "") or "").strip()
        if not identity:
            continue
        role = str(item.get("kind", "") or "selected").strip().casefold()
        if role not in {"selected", "dependency"}:
            role = "selected"
        link = {
            "sim_id": sim_id,
            "role": role,
            "categories": list(item.get("category_labels", []) or []),
            "package": str(item.get("package", "") or ""),
            "exported_as": str(item.get("exported_as", "") or ""),
        }
        current = packages.get(identity)
        links = list(current) if isinstance(current, list) else []
        links = [entry for entry in links if not (isinstance(entry, Mapping) and entry.get("sim_id") == sim_id)]
        links.append(link)
        packages[identity] = links

    return updated


def usages_for_record(record: Mapping[str, object], registry: Mapping[str, object]) -> List[Dict[str, object]]:
    """Return resolved Sim-link records for an MDB record, newest first."""
    identity = package_identity(record)
    packages = registry.get("packages") if isinstance(registry, Mapping) else None
    sims = registry.get("sims") if isinstance(registry, Mapping) else None
    if not isinstance(packages, Mapping) or not isinstance(sims, Mapping):
        return []
    raw_links = packages.get(identity)
    if not isinstance(raw_links, Sequence) or isinstance(raw_links, (str, bytes)):
        return []

    result: List[Dict[str, object]] = []
    for raw in raw_links:
        if not isinstance(raw, Mapping):
            continue
        sim_id = str(raw.get("sim_id", "") or "")
        sim = sims.get(sim_id)
        if not isinstance(sim, Mapping):
            continue
        merged = dict(sim)
        merged.update(dict(raw))
        result.append(merged)
    result.sort(key=lambda row: str(row.get("created_at", "")), reverse=True)
    return result


def selected_usages_for_record(record: Mapping[str, object], registry: Mapping[str, object]) -> List[Dict[str, object]]:
    return [row for row in usages_for_record(record, registry) if str(row.get("role", "")) == "selected"]


def dependency_usages_for_record(record: Mapping[str, object], registry: Mapping[str, object]) -> List[Dict[str, object]]:
    return [row for row in usages_for_record(record, registry) if str(row.get("role", "")) == "dependency"]


def sim_history(registry: Mapping[str, object]) -> List[Dict[str, object]]:
    """Return exported Random Sim records, newest first, with usage counts.

    Counts are reconstructed from the package-link index so older v1 registries also
    gain history summaries without a migration step.
    """
    sims = registry.get("sims") if isinstance(registry, Mapping) else None
    packages = registry.get("packages") if isinstance(registry, Mapping) else None
    if not isinstance(sims, Mapping):
        return []
    counts: Dict[str, Dict[str, int]] = {}
    if isinstance(packages, Mapping):
        for raw_links in packages.values():
            if not isinstance(raw_links, Sequence) or isinstance(raw_links, (str, bytes)):
                continue
            for raw in raw_links:
                if not isinstance(raw, Mapping):
                    continue
                sim_id = str(raw.get("sim_id", "") or "")
                if not sim_id:
                    continue
                bucket = counts.setdefault(sim_id, {"selected": 0, "dependency": 0})
                role = str(raw.get("role", "selected") or "selected").casefold()
                bucket["dependency" if role == "dependency" else "selected"] += 1

    result: List[Dict[str, object]] = []
    for sim_id, raw_sim in sims.items():
        if not isinstance(raw_sim, Mapping):
            continue
        sim = dict(raw_sim)
        sim.setdefault("sim_id", str(sim_id))
        bucket = counts.get(str(sim.get("sim_id", sim_id)), {"selected": 0, "dependency": 0})
        sim["selected_count"] = int(bucket.get("selected", 0))
        sim["dependency_count"] = int(bucket.get("dependency", 0))
        result.append(sim)
    result.sort(key=lambda row: str(row.get("created_at", "")), reverse=True)
    return result


def usages_for_sim(
    sim_id: object,
    registry: Mapping[str, object],
    role: str | None = None,
) -> List[Dict[str, object]]:
    """Return package-link records belonging to one exported Random Sim.

    Each row includes ``identity`` plus the stored link metadata. ``role`` may be
    ``selected`` or ``dependency``; omit it for both.
    """
    wanted_sim = str(sim_id or "")
    if not wanted_sim:
        return []
    wanted_role = str(role or "").casefold().strip()
    packages = registry.get("packages") if isinstance(registry, Mapping) else None
    if not isinstance(packages, Mapping):
        return []
    result: List[Dict[str, object]] = []
    for identity, raw_links in packages.items():
        if not isinstance(raw_links, Sequence) or isinstance(raw_links, (str, bytes)):
            continue
        for raw in raw_links:
            if not isinstance(raw, Mapping):
                continue
            if str(raw.get("sim_id", "") or "") != wanted_sim:
                continue
            row_role = str(raw.get("role", "selected") or "selected").casefold()
            if wanted_role and row_role != wanted_role:
                continue
            row = dict(raw)
            row["identity"] = str(identity)
            row["role"] = row_role
            result.append(row)
    result.sort(
        key=lambda row: (
            ", ".join(str(v) for v in row.get("categories", []) or []).casefold(),
            str(row.get("package", "") or "").casefold(),
        )
    )
    return result
