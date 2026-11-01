from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
import urllib.error
import urllib.request
from pathlib import Path
from typing import Callable, Dict, Iterable, Optional, Tuple
from urllib.parse import urlparse

GITHUB_REPOSITORY = "Taemaly/CAS-Curator"
LATEST_STABLE_RELEASE_API = f"https://api.github.com/repos/{GITHUB_REPOSITORY}/releases/latest"
USER_AGENT = "CAS-Curator-Updater/1.0"
_ALLOWED_DOWNLOAD_HOSTS = {"github.com", "objects.githubusercontent.com", "release-assets.githubusercontent.com"}
_VERSION_RE = re.compile(r"(?<!\d)(\d+)\.(\d+)\.(\d+)(?:[-+][0-9A-Za-z.-]+)?")
_SHA256_RE = re.compile(r"\b([0-9a-fA-F]{64})\b")


class UpdateError(RuntimeError):
    pass


def parse_semver(value: str) -> Optional[Tuple[int, int, int]]:
    """Extract a simple x.y.z tuple from app/release text."""
    match = _VERSION_RE.search(str(value or ""))
    if not match:
        return None
    return tuple(int(part) for part in match.groups())


def is_newer_version(current: str, candidate: str) -> bool:
    current_v = parse_semver(current)
    candidate_v = parse_semver(candidate)
    if current_v is None or candidate_v is None:
        return False
    return candidate_v > current_v


def _https_github_url(url: str) -> bool:
    parsed = urlparse(str(url or ""))
    return parsed.scheme.lower() == "https" and (parsed.hostname or "").lower() in _ALLOWED_DOWNLOAD_HOSTS


def _request(url: str, *, accept: str = "application/vnd.github+json") -> urllib.request.Request:
    return urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": accept,
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )


def fetch_latest_stable_release(timeout: int = 12) -> Dict[str, object]:
    """Read GitHub's official latest non-prerelease release metadata."""
    try:
        with urllib.request.urlopen(_request(LATEST_STABLE_RELEASE_API), timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (urllib.error.URLError, OSError, ValueError, json.JSONDecodeError) as exc:
        raise UpdateError(f"Could not check GitHub Releases: {exc}") from exc
    if not isinstance(payload, dict):
        raise UpdateError("GitHub returned an unexpected update response.")
    if payload.get("draft") or payload.get("prerelease"):
        raise UpdateError("The latest stable-release endpoint returned a draft or pre-release unexpectedly.")
    tag = str(payload.get("tag_name") or "").strip()
    if parse_semver(tag) is None:
        raise UpdateError("The latest release does not have a recognised version tag.")
    return payload


def release_assets(release: Dict[str, object]) -> list[Dict[str, object]]:
    raw = release.get("assets", []) if isinstance(release, dict) else []
    return [item for item in raw if isinstance(item, dict)]


def choose_installer_asset(release: Dict[str, object]) -> Optional[Dict[str, object]]:
    """Choose an official Windows installer, never a ZIP/source archive."""
    assets = release_assets(release)
    candidates = []
    for asset in assets:
        name = str(asset.get("name") or "").strip()
        lower = name.lower()
        url = str(asset.get("browser_download_url") or "")
        if not name or not _https_github_url(url):
            continue
        if not lower.endswith((".exe", ".msi")):
            continue
        score = 0
        if "installer" in lower or "setup" in lower:
            score += 100
        if "cas" in lower and "curator" in lower:
            score += 40
        if lower.endswith(".exe"):
            score += 10
        if "portable" in lower:
            score -= 80
        candidates.append((score, name.lower(), asset))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (-item[0], item[1]))
    best_score, _, best = candidates[0]
    return best if best_score > 0 else None


def choose_checksum_asset(release: Dict[str, object], installer_name: str) -> Optional[Dict[str, object]]:
    assets = release_assets(release)
    wanted = {
        f"{installer_name}.sha256".lower(),
        f"{installer_name}.sha256.txt".lower(),
        "sha256sums.txt",
        "sha256sums",
        "checksums.txt",
        "checksums.sha256",
    }
    ranked = []
    for asset in assets:
        name = str(asset.get("name") or "").strip()
        url = str(asset.get("browser_download_url") or "")
        if not name or not _https_github_url(url):
            continue
        lower = name.lower()
        if lower in wanted:
            rank = 0 if lower.startswith(installer_name.lower()) else 1
            ranked.append((rank, lower, asset))
    if not ranked:
        return None
    ranked.sort(key=lambda item: (item[0], item[1]))
    return ranked[0][2]


def _download_bytes(url: str, timeout: int = 20) -> bytes:
    if not _https_github_url(url):
        raise UpdateError("Update asset URL is not an approved GitHub HTTPS address.")
    try:
        with urllib.request.urlopen(_request(url, accept="application/octet-stream"), timeout=timeout) as response:
            return response.read()
    except (urllib.error.URLError, OSError) as exc:
        raise UpdateError(f"Could not download update verification data: {exc}") from exc


def expected_sha256(checksum_text: str, installer_name: str) -> Optional[str]:
    """Parse common SHA256SUMS or one-hash-per-file formats without guessing."""
    text = str(checksum_text or "")
    target = installer_name.lower()
    digests = []
    for line in text.splitlines():
        match = _SHA256_RE.search(line)
        if not match:
            continue
        digest = match.group(1).lower()
        digests.append(digest)
        if target and target in line.lower():
            return digest
    # A per-installer .sha256 file often contains only the bare hash. If a
    # multi-file checksum list does not name this installer, refuse to guess.
    return digests[0] if len(digests) == 1 else None


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        while True:
            block = handle.read(chunk_size)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


def default_update_directory(tag_name: str) -> Path:
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    safe_tag = re.sub(r"[^0-9A-Za-z._-]+", "_", str(tag_name or "update")).strip("._") or "update"
    return Path(base) / "CASCurator" / "Updates" / safe_tag


def download_verified_installer(
    release: Dict[str, object],
    progress: Optional[Callable[[int, int], None]] = None,
    timeout: int = 30,
) -> Dict[str, object]:
    installer = choose_installer_asset(release)
    if installer is None:
        raise UpdateError("This release does not contain a recognised Windows installer asset.")
    installer_name = str(installer.get("name") or "").strip()
    checksum_asset = choose_checksum_asset(release, installer_name)
    if checksum_asset is None:
        raise UpdateError(
            "This release does not contain a SHA-256 checksum for its installer. "
            "CAS Curator will not run an unverified update automatically."
        )

    checksum_text = _download_bytes(str(checksum_asset.get("browser_download_url") or ""), timeout=timeout).decode(
        "utf-8", errors="replace"
    )
    expected = expected_sha256(checksum_text, installer_name)
    if not expected:
        raise UpdateError("The release checksum file does not contain a SHA-256 value for the installer.")

    destination_dir = default_update_directory(str(release.get("tag_name") or "update"))
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination = destination_dir / installer_name
    partial = destination.with_suffix(destination.suffix + ".part")
    for stale in (partial,):
        try:
            stale.unlink()
        except FileNotFoundError:
            pass

    url = str(installer.get("browser_download_url") or "")
    if not _https_github_url(url):
        raise UpdateError("Installer URL is not an approved GitHub HTTPS address.")
    expected_size = int(installer.get("size") or 0)
    downloaded = 0
    try:
        with urllib.request.urlopen(_request(url, accept="application/octet-stream"), timeout=timeout) as response, partial.open("wb") as out:
            total = expected_size or int(response.headers.get("Content-Length") or 0)
            while True:
                block = response.read(1024 * 1024)
                if not block:
                    break
                out.write(block)
                downloaded += len(block)
                if progress:
                    progress(downloaded, total)
    except (urllib.error.URLError, OSError) as exc:
        try:
            partial.unlink()
        except FileNotFoundError:
            pass
        raise UpdateError(f"Could not download the update installer: {exc}") from exc

    if expected_size and downloaded != expected_size:
        try:
            partial.unlink()
        except FileNotFoundError:
            pass
        raise UpdateError(f"Installer download size mismatch ({downloaded:,} bytes received, {expected_size:,} expected).")

    actual = sha256_file(partial)
    if actual.lower() != expected.lower():
        try:
            partial.unlink()
        except FileNotFoundError:
            pass
        raise UpdateError("Installer SHA-256 verification failed. The downloaded update will not be run.")

    os.replace(partial, destination)
    return {
        "installer_path": destination,
        "installer_name": installer_name,
        "sha256": actual,
        "checksum_asset": str(checksum_asset.get("name") or ""),
        "downloaded_bytes": downloaded,
    }


def release_summary(release: Dict[str, object]) -> Dict[str, str]:
    return {
        "tag": str(release.get("tag_name") or "").strip(),
        "name": str(release.get("name") or release.get("tag_name") or "CAS Curator update").strip(),
        "body": str(release.get("body") or "").strip(),
        "published_at": str(release.get("published_at") or "").strip(),
        "html_url": str(release.get("html_url") or "").strip(),
    }
