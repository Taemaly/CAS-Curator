"""Isolated, non-Qt colour role helpers for Theme Colour Studio.

Editor changes are kept in memory until explicitly saved. All saved values
are standard HEX roles within the existing .s3theme payload. No MDB/scanner
code is imported or accessed by this module.
"""
from __future__ import annotations

import random
import re
from typing import Dict, Iterable, Mapping, Optional

MAX_IMPORT_COLOURS = 48

# Text is intentionally not a user-editable theme role. CAS Curator chooses
# one of these fixed readable defaults from the actual surface underneath it.
AUTO_MAIN_DARK = "#171717"
AUTO_MAIN_LIGHT = "#F7F7F7"
AUTO_MUTED_DARK = "#5A5A5A"
AUTO_MUTED_LIGHT = "#DADADA"

# Roles correspond to real QSS surfaces in app.py. The editor never offers
# roles from a built-in theme for mutation; the caller checks membership.
ROLE_GROUPS = (
    ("Library", (
        ("ui_top_fill", "Toolbar + search/filter panels"),
        ("ui_header_fill", "Library header strip"),
        ("ui_table_header_fill", "Table column headers"),
        ("ui_main_fill", "Main table background"),
        ("ui_row_alt", "Alternating table rows"),
        ("ui_side_fill", "Side information panel"),
        ("ui_info_card_fill", "Overview card"),
        ("ui_filter_button", "Toolbar / filter buttons"),
        ("ui_tab_fill", "Inactive detail tabs"),
        ("ui_tab_selected_fill", "Selected detail tab"),
        ("ui_thumb_fill", "Thumbnail mat"),
        ("ui_info_card_border", "Structural borders"),
        ("ui_mdb_path_border", "Master Database path border"),
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
        ("ui_guide_section_fill", "Guide Section"),
        ("ui_guide_highlight_1", "Guide Highlight 1"),
        ("ui_guide_highlight_2", "Guide Highlight 2"),
        ("ui_guide_highlight_3", "Guide Highlight 3"),
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
    # Legacy neutral preview helper retained for non-adaptive saved themes.
    luminance = .2126*r + .7152*g + .0722*b
    return "#161B22" if luminance >= 146 else "#FFFFFF"


def adaptive_text_variants(base: str) -> tuple[str, str]:
    """Legacy hue-preserving helper retained for older compatible theme data.

    Current Theme Colour Studio does not expose text colours. New runtime text
    uses ``automatic_text_for`` instead.
    """
    parsed = normalise_palette([base])
    if not parsed:
        base = '#202020'
    else:
        base = parsed[0]
    # Mixing toward black/white retains the selected hue family while giving
    # enough separation for contrast selection on very different surfaces.
    darker = mix(base, '#000000', .46)
    lighter = mix(base, '#FFFFFF', .56)
    return darker, lighter


def adaptive_text_for(base: str, background: str) -> str:
    """Pick the derived variant with the better WCAG contrast on background."""
    dark, light = adaptive_text_variants(base)
    return max((dark, light), key=lambda colour: contrast_ratio(colour, background))


def automatic_text_for(background: str, muted: bool = False) -> str:
    """Return CAS Curator's fixed automatic text colour for a surface.

    Text colour is not theme-editable. The app chooses between approved dark
    and light defaults using real contrast against the current background.
    """
    dark, light = ((AUTO_MUTED_DARK, AUTO_MUTED_LIGHT) if muted
                   else (AUTO_MAIN_DARK, AUTO_MAIN_LIGHT))
    return max((dark, light), key=lambda colour: contrast_ratio(colour, background))


def _luma(color: str) -> float:
    r, g, b = _rgb(color)
    return .2126*r + .7152*g + .0722*b


def guide_section_tint(color: str) -> str:
    """Turn any selected hue into a Guide-safe section fill.

    Guide sections are deliberately softer than ordinary controls so large
    blocks do not overpower the instructional text.  The hue remains visible,
    but dark/strong selections are lifted toward white automatically.
    """
    parsed = normalise_palette([color])
    base = parsed[0] if parsed else '#6F8FA5'
    luma = _luma(base)
    if luma < 150:
        return mix(base, '#FFFFFF', .62)
    if luma < 205:
        return mix(base, '#FFFFFF', .42)
    if luma > 242:
        return mix(base, '#7F7F7F', .10)
    return mix(base, '#FFFFFF', .18)


def guide_highlight_tint(color: str) -> str:
    """Turn any selected hue into a soft callout/highlight background."""
    parsed = normalise_palette([color])
    base = parsed[0] if parsed else '#C8D8E6'
    luma = _luma(base)
    if luma < 160:
        return mix(base, '#FFFFFF', .78)
    if luma < 215:
        return mix(base, '#FFFFFF', .60)
    if luma > 246:
        return mix(base, '#8A8A8A', .08)
    return mix(base, '#FFFFFF', .34)


def ensure_guide_highlight_separation(fill: str, background: str, min_ratio: float = 1.35) -> str:
    """Keep a Guide highlight visibly distinct from the Guide reading surface.

    Text contrast and highlight visibility are separate concerns: a callout can
    contain perfectly readable text while its fill is almost indistinguishable
    from the surrounding Guide background. Preserve the selected hue as much as
    possible, then gently darken on light surfaces (or lighten on dark surfaces)
    until the callout reaches the minimum surface-to-surface contrast.
    """
    parsed_fill = normalise_palette([fill])
    parsed_bg = normalise_palette([background])
    if not parsed_fill or not parsed_bg:
        return fill
    current = parsed_fill[0]
    bg = parsed_bg[0]
    if contrast_ratio(current, bg) >= min_ratio:
        return current

    target = '#3F3F3F' if _luma(bg) >= 145 else '#FFFFFF'
    best = current
    best_ratio = contrast_ratio(current, bg)
    # Small increments keep the original hue much more faithfully than a hard
    # fallback colour while still guaranteeing visible callout separation.
    for step in range(6, 61, 2):
        candidate = mix(current, target, step / 100.0)
        ratio = contrast_ratio(candidate, bg)
        if ratio > best_ratio:
            best, best_ratio = candidate, ratio
        if ratio >= min_ratio:
            return candidate
    return best


def guide_highlight_border(fill: str) -> str:
    """Return a stronger edge colour for a flat Guide highlight block."""
    parsed = normalise_palette([fill])
    base = parsed[0] if parsed else '#D8D8D8'
    target = '#4B4B4B' if _luma(base) >= 145 else '#FFFFFF'
    return mix(base, target, .34)




def _relative_luminance(color: str) -> float:
    """WCAG relative luminance for a normalized HEX colour."""
    channels = []
    for value in _rgb(color):
        channel = value / 255.0
        channels.append(channel / 12.92 if channel <= 0.04045 else ((channel + 0.055) / 1.055) ** 2.4)
    r, g, b = channels
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_ratio(foreground: str, background: str) -> float:
    """Return the WCAG contrast ratio between two HEX colours (1.0..21.0)."""
    fg = normalise_palette([foreground])
    bg = normalise_palette([background])
    if not fg or not bg:
        raise ValueError('Contrast colours must be valid HEX values.')
    a = _relative_luminance(fg[0])
    b = _relative_luminance(bg[0])
    lighter, darker = max(a, b), min(a, b)
    return (lighter + 0.05) / (darker + 0.05)


def _app_text_for(background: str, dark: str = '#171717', light: str = '#F7F7F7', threshold: float = 148) -> str:
    """Mirror CAS Curator's runtime foreground-selection rule for custom themes."""
    return light if _luma(background) < threshold else dark


def readability_report(theme: Mapping[str, object]) -> list[dict]:
    """Advisory contrast report using CAS Curator's automatic text colours.

    Users choose surfaces and accents only. Main, muted, instructional and
    Guide text are selected automatically from fixed readable defaults.
    """
    resolved = resolve_preview_roles(theme)

    def main_on(background: str) -> str:
        return automatic_text_for(background, muted=False)

    def muted_on(background: str) -> str:
        return automatic_text_for(background, muted=True)

    checks = [
        ('Library', 'ui_top_fill', 'Toolbar + search/filter panels', 'Automatic text', main_on(resolved['ui_top_fill'])),
        ('Library', 'ui_header_fill', 'Library header strip', 'Automatic text', main_on(resolved['ui_header_fill'])),
        ('Library', 'ui_table_header_fill', 'Table column headers', 'Automatic text', main_on(resolved['ui_table_header_fill'])),
        ('Library', 'ui_main_fill', 'Main table background', 'Automatic text', main_on(resolved['ui_main_fill'])),
        ('Library', 'ui_row_alt', 'Alternating table rows', 'Automatic text', main_on(resolved['ui_row_alt'])),
        ('Library', 'ui_side_fill', 'Side information panel', 'Automatic text', main_on(resolved['ui_side_fill'])),
        ('Library', 'ui_info_card_fill', 'Overview card', 'Automatic text', main_on(resolved['ui_info_card_fill'])),
        ('Library', 'ui_filter_button', 'Toolbar / filter buttons', 'Automatic text', main_on(resolved['ui_filter_button'])),
        ('Library', 'ui_tab_fill', 'Inactive detail tabs', 'Automatic text', main_on(resolved['ui_tab_fill'])),
        ('Library', 'ui_tab_selected_fill', 'Selected detail tab', 'Automatic text', main_on(resolved['ui_tab_selected_fill'])),
        ('Library', 'ui_thumb_fill', 'Thumbnail mat', 'Automatic muted text', muted_on(resolved['ui_thumb_fill'])),
        ('Scan', 'ui_scan_page_fill', 'Scan page', 'Automatic text', main_on(resolved['ui_scan_page_fill'])),
        ('Scan', 'ui_scan_card_fill', 'Scan section cards', 'Automatic text', main_on(resolved['ui_scan_card_fill'])),
        ('Scan', 'ui_scan_row_fill', 'Joined setup rows', 'Automatic text', main_on(resolved['ui_scan_row_fill'])),
        ('Scan', 'ui_scan_input_fill', 'Path fields / dropdowns', 'Automatic text', main_on(resolved['ui_scan_input_fill'])),
        ('Scan', 'ui_scan_input_fill', 'Input placeholders / hints', 'Automatic muted text', muted_on(resolved['ui_scan_input_fill'])),
        ('Scan', 'ui_scan_button_fill', 'Scan controls / buttons', 'Automatic text', main_on(resolved['ui_scan_button_fill'])),
        ('Scan', 'ui_scan_log_fill', 'Scan log / output area', 'Automatic text', main_on(resolved['ui_scan_log_fill'])),
        ('Scan', 'ui_scan_progress_fill', 'Progress bar background', 'Automatic text', main_on(resolved['ui_scan_progress_fill'])),
        ('Guide', 'ui_guide_page_fill', 'Guide page', 'Automatic text', main_on(resolved['ui_guide_page_fill'])),
        ('Guide', 'ui_guide_card_fill', 'Guide navigation cards', 'Automatic text', main_on(resolved['ui_guide_card_fill'])),
        ('Guide', 'ui_guide_browser_fill', 'Guide text area', 'Automatic text', main_on(resolved['ui_guide_browser_fill'])),
        ('Guide', 'ui_guide_section_fill', 'Guide Section', 'Automatic text', main_on(resolved['ui_guide_section_fill'])),
        ('Guide', 'ui_guide_highlight_1', 'Guide Highlight 1', 'Automatic text', main_on(resolved['ui_guide_highlight_1'])),
        ('Guide', 'ui_guide_highlight_2', 'Guide Highlight 2', 'Automatic text', main_on(resolved['ui_guide_highlight_2'])),
        ('Guide', 'ui_guide_highlight_3', 'Guide Highlight 3', 'Automatic text', main_on(resolved['ui_guide_highlight_3'])),
    ]

    output = []
    for group, role, label, text_role, foreground in checks:
        background = str(resolved[role])
        ratio = contrast_ratio(foreground, background)
        status = 'poor' if ratio < 3.0 else ('borderline' if ratio < 4.5 else 'good')
        output.append({
            'group': group, 'role': role, 'label': label, 'text_role': text_role,
            'foreground': foreground, 'background': background,
            'ratio': ratio, 'status': status,
        })
    return output

def resolve_preview_roles(theme: Mapping[str, object]) -> Dict[str, str]:
    """Resolve the editable Theme Studio roles exactly as a custom theme does.

    Theme Studio only edits imported/custom themes.  A palette import may omit
    many direct ``ui_*`` keys, so the application derives those surfaces at
    runtime.  Keep the preview on the same fallback ladder as ``app.py`` rather
    than inventing a separate approximation.
    """
    def get(key, default):
        val = theme.get(key)
        parsed = normalise_palette([val]) if isinstance(val, str) else []
        return parsed[0] if parsed else default

    accent = get('accent', '#537988')
    soft = get('accent_soft', '#CBDDD9')

    # General custom-theme surfaces (mirrors MainWindow.apply_theme).
    top = get('ui_top_fill', mix(soft, accent, .10))
    side = get('ui_side_fill', top)
    header = get('ui_header_fill', mix(top, '#F5F5F5', .42))
    info = get('ui_info_fill', mix(top, '#FFFFFF', .55))
    button = get('ui_filter_button', info)
    scroll = get('ui_scrollbar_color', accent)
    # Toolbar panel and Search / filter panel are intentionally one shared surface.
    # Keep legacy ui_filter_panel_fill data readable in old files, but do not let it
    # diverge from ui_top_fill in the current editor/runtime model.
    filter_panel = top
    table_header = get('ui_table_header_fill', header)
    main = get('ui_main_fill', '#FAFAFA')
    row_alt = get('ui_row_alt', '#EEEEEE')
    info_card = get('ui_info_card_fill', info)
    tab = get('ui_tab_fill', header)
    # The real app's default selected detail tab is white, not info_fill.
    tab_selected = get('ui_tab_selected_fill', '#FFFFFF')

    # Custom themes which do not explicitly save ui_info_card_border fall back
    # to the same structural border calculation used by app.py.
    structural = get('ui_info_card_border', mix(accent, '#777777', .58))
    mdb_path_border = get('ui_mdb_path_border', '#C4C4C4')

    # Thumbnail mat has its own fixed fallback until a custom role is supplied.
    thumb_default = '#F7F7F7' if _luma(side) < 145 else '#FCFCFC'
    thumb = get('ui_thumb_fill', thumb_default)

    # Scanner fallbacks use the resolved custom Library surfaces.
    scan_page = mix(side, info, .70)
    if _luma(side) < 145:
        scan_row = mix(side, header, .42)
        scan_input = mix(side, header, .72)
        scan_log = mix(side, header, .34)
        scan_progress = mix(side, header, .52)
        scan_button = mix(header, button, .24)
    else:
        scan_row = mix(side, header, .48)
        scan_input = mix(info, scroll, .20)
        scan_log = mix(info, scroll, .12)
        scan_progress = mix(info, scroll, .14)
        scan_button = mix(info, scroll, .10)

    # Guide custom fallbacks begin from accent_soft, but their border falls
    # through the app's structural-border policy unless explicitly overridden.
    guide_page = mix(soft, '#FFFFFF', .68)
    guide_card = mix(soft, '#FFFFFF', .48)
    guide_browser = get('ui_guide_browser_fill', mix(soft, '#FFFFFF', .78))
    guide_border = get('ui_guide_border', structural)
    guide_section = guide_section_tint(get('ui_guide_section_fill', accent))
    guide_highlight_1 = ensure_guide_highlight_separation(guide_highlight_tint(get('ui_guide_highlight_1', soft)), guide_browser)
    guide_highlight_2 = ensure_guide_highlight_separation(guide_highlight_tint(get('ui_guide_highlight_2', scroll)), guide_browser)
    guide_highlight_3 = ensure_guide_highlight_separation(guide_highlight_tint(get('ui_guide_highlight_3', accent)), guide_browser)

    defaults = {
        'text': AUTO_MAIN_DARK,
        'muted': AUTO_MUTED_DARK,
        'ui_top_fill': top,
        'ui_filter_panel_fill': filter_panel,
        'ui_header_fill': header,
        'ui_table_header_fill': table_header,
        'ui_main_fill': main,
        'ui_row_alt': row_alt,
        'ui_side_fill': side,
        'ui_info_card_fill': info_card,
        'ui_filter_button': button,
        'ui_tab_fill': tab,
        'ui_tab_selected_fill': tab_selected,
        'ui_thumb_fill': thumb,
        'ui_info_card_border': structural,
        'ui_mdb_path_border': mdb_path_border,
        'ui_scrollbar_color': scroll,
        'ui_scan_page_fill': get('ui_scan_page_fill', scan_page),
        'ui_scan_card_fill': get('ui_scan_card_fill', side),
        'ui_scan_row_fill': get('ui_scan_row_fill', scan_row),
        'ui_scan_input_fill': get('ui_scan_input_fill', scan_input),
        'ui_scan_button_fill': get('ui_scan_button_fill', scan_button),
        'ui_scan_log_fill': get('ui_scan_log_fill', scan_log),
        'ui_scan_progress_fill': get('ui_scan_progress_fill', scan_progress),
        'ui_guide_page_fill': get('ui_guide_page_fill', guide_page),
        'ui_guide_card_fill': get('ui_guide_card_fill', guide_card),
        'ui_guide_browser_fill': guide_browser,
        'ui_guide_section_fill': guide_section,
        'ui_guide_highlight_1': guide_highlight_1,
        'ui_guide_highlight_2': guide_highlight_2,
        'ui_guide_highlight_3': guide_highlight_3,
        'ui_guide_border': guide_border,
    }
    return defaults


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
    # Text is automatic rather than user-editable. Keep standard values in the
    # portable payload for compatibility with older .s3theme readers.
    result['text'] = AUTO_MAIN_DARK
    result['muted'] = AUTO_MUTED_DARK
    result['_adaptive_text'] = True
    result['_auto_text'] = True
    return result


def shuffle_roles(assignments: Mapping[str, Optional[str]], locked: Mapping[str, bool],
                  palette: Iterable[str], rng: Optional[random.Random] = None,
                  protected_colours: Optional[Iterable[str]] = None,
                  current_colours: Optional[Mapping[str, str]] = None) -> dict:
    """Shuffle unlocked roles while preserving two independent lock concepts.

    ``locked`` is the Section Lock map: those exact roles never move.
    ``protected_colours`` is the Shuffle Lock set: those colours must still be
    present after each shuffle, but may move to any unlocked role.  The optional
    ``current_colours`` mapping supplies effective colours for roles whose
    assignment is ``None`` (meaning "keep saved colour").
    """
    choices = normalise_palette(palette)
    protected = normalise_palette(protected_colours or [])
    if len(choices) < 2 and not protected:
        return dict(assignments)
    rng = rng or random.Random()
    output = dict(assignments)
    current = dict(current_colours or {})
    unlocked_roles = [role for role in ROLES if not locked.get(role)]
    if not unlocked_roles:
        return output

    fixed_colours = {
        current.get(role) or output.get(role)
        for role in ROLES
        if locked.get(role)
    }
    movable_protected = [colour for colour in protected if colour not in fixed_colours]
    if len(movable_protected) > len(unlocked_roles):
        raise ValueError('More Shuffle Lock colours are selected than unlocked sections can hold.')

    # Reserve random unlocked slots for every Shuffle Lock colour. This guarantees
    # survival without pinning a protected colour to the section where it started.
    reserve_roles = list(unlocked_roles)
    rng.shuffle(reserve_roles)
    protected_order = list(movable_protected)
    rng.shuffle(protected_order)
    reserved = set()
    for role, colour in zip(reserve_roles, protected_order):
        output[role] = colour
        reserved.add(role)

    pool = choices or protected_order
    cycle: list[str] = []
    for role in unlocked_roles:
        if role in reserved:
            continue
        old = current.get(role) or output.get(role)
        options = [item for item in pool if item != old]
        options = options or list(pool)
        if not cycle:
            cycle = list(pool)
            rng.shuffle(cycle)
        candidate = cycle.pop()
        if candidate == old and options:
            candidate = rng.choice(options)
        output[role] = candidate
    return output
