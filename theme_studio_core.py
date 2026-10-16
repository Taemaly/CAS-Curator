"""Isolated, non-Qt colour role helpers for the private DEV Visual Theme Studio.

Editor changes are kept in memory until explicitly saved. All saved values
are standard HEX roles within the existing .s3theme payload. No MDB/scanner
code is imported or accessed by this module.
"""
from __future__ import annotations

import random
import re
from typing import Dict, Iterable, Mapping, Optional

MAX_IMPORT_COLOURS = 48

# Roles correspond to real QSS surfaces in app.py. The editor never offers
# roles from a built-in theme for mutation; the caller checks membership.
ROLE_GROUPS = (
    ("Library", (
        ("ui_top_fill", "Toolbar"),
        ("ui_filter_panel_fill", "Search / filter panel"),
        ("ui_header_fill", "Library header strip"),
        ("ui_table_header_fill", "Table column headers"),
        ("ui_main_fill", "Main table background"),
        ("ui_row_alt", "Alternating table rows"),
        ("ui_side_fill", "Side information panel"),
        ("ui_info_card_fill", "Overview card"),
        ("ui_filter_button", "Buttons / filter dropdowns"),
        ("ui_tab_fill", "Inactive detail tabs"),
        ("ui_tab_selected_fill", "Selected detail tab"),
        ("ui_thumb_fill", "Thumbnail mat"),
        ("ui_info_card_border", "Structural borders"),
        ("ui_scrollbar_color", "Scrollbar / accent"),
    )),
    ("Scan", (
        ("ui_scan_page_fill", "Scan page"),
        ("ui_scan_card_fill", "Scan section cards"),
        ("ui_scan_row_fill", "Joined setup rows"),
        ("ui_scan_input_fill", "Path fields / dropdowns"),
        ("ui_scan_button_fill", "Scan controls / buttons"),
        ("ui_scan_log_fill", "Scan log / output area"),
        ("ui_scan_progress_fill", "Progress bar background"),
    )),
    ("Guide", (
        ("ui_guide_page_fill", "Guide page"),
        ("ui_guide_card_fill", "Guide navigation cards"),
        ("ui_guide_browser_fill", "Guide text area"),
        ("ui_guide_border", "Guide borders"),
    )),
)
ROLES = dict((key, label) for _group, entries in ROLE_GROUPS for key, label in entries)


def normalise_palette(colours: Iterable[str]) -> list[str]:
    result = []
    for raw in colours:
        value = str(raw or "").strip().lstrip("#")
        if re.fullmatch(r"[0-9a-fA-F]{3}", value):
            value = "".join(c * 2 for c in value)
        if re.fullmatch(r"[0-9a-fA-F]{6}", value):
            hex_value = "#" + value.upper()
            if hex_value not in result:
                result.append(hex_value)
        if len(result) >= MAX_IMPORT_COLOURS:
            break
    return result


def add_custom_colour(palette: Iterable[str], colour: str) -> tuple[list[str], bool]:
    """Append a unique user colour without replacing existing imported colours.

    Returns (updated_palette, added). Duplicate selections are harmless.
    Raises if no free slot remains or the selected colour is invalid.
    """
    current = normalise_palette(palette)
    normalized = normalise_palette([colour])
    if not normalized:
        raise ValueError('Select a valid six-digit HEX colour.')
    chosen = normalized[0]
    if chosen in current:
        return current, False
    if len(current) >= MAX_IMPORT_COLOURS:
        raise ValueError(f'This palette already contains {MAX_IMPORT_COLOURS} colours.')
    return current + [chosen], True


def _rgb(h):
    return tuple(int(h.lstrip('#')[n:n+2], 16) for n in (0, 2, 4))


def mix(a: str, b: str, fraction: float) -> str:
    return "#" + "".join(f"{round(x * (1-fraction) + y * fraction):02X}" for x, y in zip(_rgb(a), _rgb(b)))


def text_for(color: str) -> str:
    r, g, b = _rgb(color)
    # Accessible enough for the miniature mock preview; the real application
    # also derives foreground colours when applying a custom theme.
    luminance = .2126*r + .7152*g + .0722*b
    return "#161B22" if luminance >= 146 else "#FFFFFF"


def resolve_preview_roles(theme: Mapping[str, object]) -> Dict[str, str]:
    """Read the effective colours of an imported theme, including fallbacks.

    These are a *miniature illustrative preview* of the same stylesheet roles,
    not a screenshot or a separate Qt rendering of the running main window.
    """
    def get(key, default):
        val = theme.get(key)
        return normalise_palette([val])[0] if isinstance(val, str) and normalise_palette([val]) else default

    accent = get('accent', '#537988')
    soft = get('accent_soft', '#CBDDD9')
    top = get('ui_top_fill', mix(soft, accent, .10))
    side = get('ui_side_fill', top)
    header = get('ui_header_fill', mix(top, '#F5F5F5', .42))
    info = get('ui_info_fill', mix(top, '#FFFFFF', .55))
    btn = get('ui_filter_button', info)
    scroll = get('ui_scrollbar_color', accent)
    dark = sum(_rgb(side))/3 < 145
    if dark:
        scan_row = mix(side, header, .42)
        scan_input = mix(side, header, .72)
        scan_log = mix(side, header, .34)
        scan_progress = mix(side, header, .52)
        scan_btn = mix(header, btn, .24)
    else:
        scan_row = mix(side, header, .48)
        scan_input = mix(info, scroll, .20)
        scan_log = mix(info, scroll, .12)
        scan_progress = mix(info, scroll, .14)
        scan_btn = mix(info, scroll, .10)
    defaults = {
        'ui_top_fill': top,
        'ui_filter_panel_fill': top,
        'ui_header_fill': header,
        'ui_table_header_fill': header,
        'ui_main_fill': '#FAFAFA',
        'ui_row_alt': '#EEEEEE',
        'ui_side_fill': side,
        'ui_info_card_fill': info,
        'ui_filter_button': btn,
        'ui_tab_fill': header,
        'ui_tab_selected_fill': info,
        'ui_thumb_fill': get('ui_thumb_fill', mix(get('thumb','#70757A'), '#686868', .3)),
        'ui_info_card_border': get('ui_info_card_border','#777777'),
        'ui_scrollbar_color': scroll,
        'ui_scan_page_fill': mix(side, info, .7),
        'ui_scan_card_fill': side,
        'ui_scan_row_fill': scan_row,
        'ui_scan_input_fill': scan_input,
        'ui_scan_button_fill': scan_btn,
        'ui_scan_log_fill': scan_log,
        'ui_scan_progress_fill': scan_progress,
        'ui_guide_page_fill': mix(soft, '#FFFFFF', .68),
        'ui_guide_card_fill': mix(soft, '#FFFFFF', .48),
        'ui_guide_browser_fill': mix(soft, '#FFFFFF', .78),
        'ui_guide_border': mix(accent, '#B8B8B8', .48),
    }
    return {role: get(role, fallback) for role, fallback in defaults.items()}


def edited_theme(base_theme: dict, assignments: Mapping[str, Optional[str]]) -> dict:
    """Apply only explicitly edited roles; untouched properties stay bytewise equal."""
    result = dict(base_theme)
    for role, choice in assignments.items():
        if role not in ROLES:
            raise ValueError(f"Unrecognized theme section: {role}")
        if choice is not None:
            palette = normalise_palette([choice])
            if not palette:
                raise ValueError(f"Invalid colour for {role}: {choice}")
            result[role] = palette[0]
    result['_custom'] = True
    return result


def shuffle_roles(assignments: Mapping[str, Optional[str]], locked: Mapping[str, bool],
                  palette: Iterable[str], rng: Optional[random.Random] = None) -> dict:
    choices = normalise_palette(palette)
    if len(choices) < 2:
        return dict(assignments)
    rng = rng or random.Random()
    output = dict(assignments)
    previous = []
    for role in ROLES:
        if locked.get(role):
            continue
        old = output.get(role)
        options = [item for item in choices if item != old]
        options = options or choices
        # Choose from shuffled palette, cycling over all colours rather than
        # weighting the first two swatches more heavily.
        if not previous:
            previous = choices.copy()
            rng.shuffle(previous)
        candidate = previous.pop()
        if candidate == old and options:
            candidate = rng.choice(options)
        output[role] = candidate
    return output
