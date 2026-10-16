"""Safe metadata helpers for DEV-only downloadable CAS Curator palettes.

No PySide6 dependency: the metadata/parser rules can be tested without the GUI.
A site's generic browser title is deliberately *not* considered an original
palette name, and missing website metadata never silently becomes one.
"""
from __future__ import annotations

import colorsys
import html
import json
import re
from pathlib import Path
from typing import Dict, List, Any
from urllib.parse import urlsplit
from urllib.request import Request, urlopen

SUPPORTED_SITES = {"coolors.co": "Coolors", "lospec.com": "Lospec"}
GENERIC_NAMES = {
    "", "palette", "color palette", "colour palette", "colors", "colours",
    "download", "export", "coolors", "lospec", "new palette", "my palette",
    "color palettes", "colour palettes", "palette generator",
    "untitled", "untitled palette", "gimp palette",
}


def palette_source(url: str) -> str:
    try:
        parsed = urlsplit(url.strip())
        host = (parsed.hostname or "").lower()
        if parsed.scheme not in {"http", "https"}:
            return ""
        for allowed, label in SUPPORTED_SITES.items():
            if host == allowed or host == "www." + allowed:
                return label
    except ValueError:
        pass
    return ""


def find_palette_url(raw: str) -> str:
    for candidate in re.findall(r"https?://[^\s<>\"']+", raw or "", re.I):
        url = candidate.rstrip(",.;)]}")
        if palette_source(url):
            return url
    return ""


def _usable_name(value: object) -> str:
    if not isinstance(value, str):
        return ""
    value = re.sub(r"\s+", " ", html.unescape(value)).strip().strip("-_:")[:100]
    if value.casefold() in GENERIC_NAMES or not value:
        return ""
    if re.fullmatch(r"[0-9a-fA-F#_\- ]{10,}", value):
        return ""  # palette colour-code strings aren't human titles
    if re.search(r"(super fast color palettes generator|color palettes generator|color palette generator|palette list - lospec)", value, re.I):
        return ""
    return value


def embedded_palette_name(raw: str) -> str:
    """Read explicitly named palettes in JSON, GPL, or clear text headers."""
    try:
        data = json.loads(raw)
        if isinstance(data, dict):
            for key in ("original_name", "palette_name", "paletteName", "title", "name"):
                value = _usable_name(data.get(key))
                if value:
                    return value
            for key in ("palette", "data"):
                sub = data.get(key)
                if isinstance(sub, dict):
                    for name_key in ("palette_name", "title", "name"):
                        value = _usable_name(sub.get(name_key))
                        if value:
                            return value
    except (ValueError, TypeError):
        pass
    for line in raw.splitlines()[:16]:
        match = re.match(r"^\s*(?:#\s*)?(?:Name|Palette Name|Title)\s*:\s*(.+)$", line, re.I)
        if match:
            return _usable_name(match.group(1))
    return ""


def website_palette_name(url: str) -> str:
    """Best-effort supported-site title lookup. No media/package downloads.

    Restricted to the two known palette sites. Only a non-generic OG/browser
    title is accepted. Network failure just means use a suggested name.
    """
    if not palette_source(url):
        return ""
    try:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0 (CASCurator DEV Palette Name Lookup)"})
        with urlopen(req, timeout=3.0) as response:
            final_url = response.geturl()
            if not palette_source(final_url):
                return ""  # refuse titles after off-site redirect
            blob = response.read(384 * 1024).decode("utf-8", errors="ignore")
    except (OSError, ValueError, TimeoutError):
        return ""
    meta = re.search(r"<meta\b(?=[^>]*(?:property|name)\s*=\s*['\"](?:og:title|twitter:title)['\"])[^>]*>", blob, re.I)
    title = ""
    if meta:
        content = re.search(r"\bcontent\s*=\s*(['\"])(.*?)\1", meta.group(0), re.I | re.S)
        if content:
            title = content.group(2)
    if not title:
        match = re.search(r"<title[^>]*>(.*?)</title>", blob, re.I | re.S)
        if match:
            title = re.sub(r"<[^>]+>", "", match.group(1))
    title = html.unescape(title).strip()
    # Only remove a branded suffix from an actual palette-specific page title.
    title = re.sub(r"\s*[|\-–]\s*(?:Lospec|Coolors)(?:\s.*)?$", "", title, flags=re.I)
    return _usable_name(title)


def creative_palette_names(colours: List[str]) -> List[str]:
    """Stable nature/food/atmosphere suggestions for an *unnamed* palette.

    These are explicitly creative guesses, never a website's original title.
    Use the palette as a whole rather than merely joining two hue labels.
    Nothing is fetched over the network and the same colours give the same names.
    """
    parsed = []
    for value in colours[:48]:
        match = re.fullmatch(r"#?([0-9a-fA-F]{6})", str(value).strip())
        if not match:
            continue
        hex_value = match.group(1).upper()
        red, green, blue = (int(hex_value[i:i+2], 16) / 255 for i in (0, 2, 4))
        hue, sat, brightness = colorsys.rgb_to_hsv(red, green, blue)
        parsed.append((hex_value, hue, sat, brightness, red, green, blue))
    if not parsed:
        return ["Secret Garden", "Sunday Morning", "Cloudberry Daydream"]

    counts = dict.fromkeys(("earth", "leaf", "rose", "peach", "sun", "water", "blue", "violet", "neutral"), 0)
    pastel = dusky = dark = 0
    for _, hue, sat, val, r, g, b in parsed:
        pastel += val > .74 and sat < .41
        dusky += .23 < val < .79 and sat < .39
        dark += val < .36
        # Treat warm beige, terracotta and muted brown as natural/earthy rather
        # than naming everything Gold or Rose based on the hue alone.
        if (.045 < hue < .17 and .16 < sat < .65 and val < .82) or (
                .045 < hue < .17 and .055 < sat < .34 and val >= .82):
            counts["earth"] += 1
        elif sat < .11:
            counts["neutral"] += 1
        elif .19 <= hue < .46:
            counts["leaf"] += 1
        elif .46 <= hue < .57:
            counts["water"] += 1
        elif .57 <= hue < .71:
            counts["blue"] += 1
        elif .71 <= hue < .89:
            counts["violet"] += 1
        elif .89 <= hue or hue < .028:
            counts["rose"] += 1
        elif .028 <= hue < .092:
            counts["peach"] += 1
        elif .092 <= hue < .19:
            counts["sun"] += 1
        else:
            counts["neutral"] += 1
    # Context-sensitive themes: an earthy palette with greens should sound
    # like a forest/cottage, not a pair of literal colour names.
    if counts["earth"] and counts["leaf"]:
        mood = "woodland"
    elif counts["earth"] >= max(2, len(parsed) // 3):
        mood = "earth"
    elif counts["rose"] and (counts["peach"] or pastel >= len(parsed) // 2):
        mood = "blossom"
    elif counts["blue"] and (counts["water"] or counts["leaf"]):
        mood = "lagoon"
    elif counts["violet"] and counts["blue"]:
        mood = "twilight"
    elif counts["leaf"] and counts["water"]:
        mood = "garden"
    else:
        priority = ("rose", "peach", "sun", "leaf", "water", "blue", "violet", "earth", "neutral")
        main = max(priority, key=lambda key: (counts[key], -priority.index(key)))
        mood = {
            "rose": "blossom", "peach": "orchard", "sun": "honey", "leaf": "garden",
            "water": "lagoon", "blue": "lagoon", "violet": "twilight", "earth": "earth",
            "neutral": "stone",
        }[main]
    if dark >= len(parsed) // 2 + 1 and mood not in {"woodland", "earth"}:
        mood = "midnight"
    elif (dusky >= len(parsed) // 2 + 1 and mood == "lagoon"):
        mood = "mist"
    elif pastel >= len(parsed) // 2 + 1 and mood == "twilight":
        mood = "lilac"
    names = {
        "woodland": ["Mossy Cottage", "Woodland Linen", "Sage & Biscuit", "Forest Chai",
                     "Fern & Honey", "Cedar Morning", "Autumn Herb Garden", "Hazelnut Grove",
                     "Meadow Hearth", "Willow Tea"],
        "earth": ["Spiced Oatmilk", "Cinnamon Orchard", "Honeyed Autumn", "Terracotta Hearth",
                  "Warm Earthlight", "Maple Cottage", "Amber & Oats", "Toasted Hazelnut",
                  "Autumn Market", "Clove & Cream"],
        "blossom": ["Cherry Blossom Tea", "Rosewater Cream", "Petal Meringue", "Peony Morning",
                     "Strawberry Cloud", "Blushing Orchard", "Raspberry Macaron", "Spring Patisserie",
                     "Rose Petal Rain", "Cotton Candy Garden"],
        "orchard": ["Peach Sorbet", "Apricot Orchard", "Papaya Sunrise", "Peaches & Cream",
                    "Golden Nectar", "Summer Fruit Tart", "Coral Sunset", "Peach Tea Garden",
                    "Sweet Persimmon", "Honey Apricot"],
        "honey": ["Honeycomb Morning", "Golden Hour", "Lemon Biscuit", "Sunflower Fields",
                  "Butterscotch Garden", "Warm Brioche", "Dandelion Tea", "Amber Sunbeam",
                  "Marigold Honey", "Vanilla Harvest"],
        "garden": ["Morning Fern", "Wildflower Meadow", "Greenhouse Tea", "Fern & Dew",
                   "Spring Herb Garden", "Jasmine Leaves", "Bamboo Morning", "Water Lily Garden",
                   "Moss & Mint", "Summer Canopy"],
        "lagoon": ["Misty Lagoon", "Sea Glass Garden", "Coastal Drift", "Ocean Daydream",
                   "Blueberry Tide", "Seaside Hydrangea", "Saltwater Morning", "Aqua Rain",
                   "Lagoon Breeze", "Blue Iris Bay"],
        "mist": ["Rainy Coast", "Misty Harbour", "Foggy Seashore", "Stormglass Bay",
                 "Washed Sea Glass", "Blue Moon Tide", "Drifting Clouds", "Frosted Coast",
                 "Rain on the Pier", "Moonlit Waves"],
        "twilight": ["Twilight Orchid", "Blackberry Velvet", "Moonlit Iris", "Evening Wisteria",
                     "Velvet Dusk", "Plum Nightfall", "Lavender Cosmos", "Moonflower Garden",
                     "Dusky Bluebell", "Midnight Lavender"],
        "lilac": ["Lilac Meringue", "Lavender Milk", "Wisteria Morning", "Violet Mousse",
                  "Lavender Haze", "Blueberry Blossom", "Pastel Moonflower", "Lilac Rain",
                  "Sugarplum Sky", "Lavender Cloud"],
        "stone": ["Rainwashed Stone", "Cloudy Sunday", "Misty Pebble", "Moonstone Haze",
                  "Frosted Linen", "Silver Willow", "Winter Porcelain", "Dove Morning",
                  "Rain on Slate", "Foggy Orchard"],
        "midnight": ["Midnight Garden", "Moonlit Velvet", "After Dark", "Nocturne Blossom",
                     "Velvet Shadows", "Blackberry Night", "Nightfall Meadow", "Ink & Moonlight",
                     "Starlit Orchard", "Silent Midnight"],
    }[mood]
    # Stable rotation makes similarly themed but distinct palettes less likely
    # to receive the identical first suggestion. Do not use Python hash() here:
    # it changes between processes, breaking persistence/reproducibility.
    import zlib
    seed = zlib.crc32("-".join(item[0] for item in parsed).encode("ascii"))
    start = seed % len(names)
    return names[start:] + names[:start]


def suggested_palette_name(colours: List[str], filename: str = "", source_url: str = "") -> str:
    """Stable, descriptive fallback, not misrepresented as an original name."""
    stem = Path(filename).stem if filename else ""
    stem = re.sub(r"[_-]+", " ", stem).strip()
    generic_export = bool(re.fullmatch(r"(?i)(?:download(?:ed)?|export(?:ed)?|palette|coolors|lospec)(?:[ _-]*[0-9]*)?", stem))
    if (_usable_name(stem) and not generic_export
            and not re.fullmatch(r"(?:[0-9a-fA-F]{3,8}\s+){1,}[0-9a-fA-F]{3,8}", stem)):
        return stem
    if palette_source(source_url) == "Lospec":
        try:
            slug = urlsplit(source_url).path.split("/palette-list/")[1].split("/")[0].split(".")[0]
            candidate = re.sub(r"[-_]+", " ", slug).strip().title()
            if _usable_name(candidate):
                return candidate
        except (IndexError, ValueError):
            pass
    return creative_palette_names(colours)[0]


def palette_metadata(raw: str, colours: List[str], filename: str = "", online_lookup: bool = True) -> Dict[str, Any]:
    url = find_palette_url(raw)
    site = palette_source(url)
    original = embedded_palette_name(raw)
    if not original and online_lookup and url:
        original = website_palette_name(url)
    suggested = original or suggested_palette_name(colours, filename, url)
    creative = creative_palette_names(colours)
    suggested_names = [suggested] + [n for n in creative if n.casefold() != suggested.casefold()]
    return {
        "original_name": original,
        "suggested_name": suggested,
        "suggested_names": suggested_names,
        "source_site": site or ("Palette file" if filename else "Pasted palette"),
        "source_url": url,
        "source_filename": Path(filename).name if filename else "",
        "name_source": "original" if original else "suggested",
    }
