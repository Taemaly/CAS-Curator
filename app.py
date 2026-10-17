from __future__ import annotations

import colorsys
import html
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
from urllib.parse import quote_plus

from PySide6.QtCore import (
    QAbstractTableModel,
    QEvent,
    QModelIndex,
    QItemSelectionModel,
    QObject,
    QSettings,
    QSortFilterProxyModel,
    Qt,
    QThread,
    QProcess,
    QTimer,
    Signal,
    QPoint,
    QUrl,
    QSize,
    QRect,
)
from PySide6.QtGui import QAction, QActionGroup, QColor, QDesktopServices, QFontMetrics, QIcon, QPainter, QPen, QPixmap, QPolygon
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QListView,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QProgressBar,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStyledItemDelegate,
    QStyle,
    QStyleFactory,
    QStackedWidget,
    QTabWidget,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from mdb_reader import MDBData, MDBWorkbook, WorkbookChangedError, target_to_local_path
from mdb_compat import inspect_mdb_compatibility, create_compatible_copy
from custom_theme_tools import palette_metadata, palette_source
from random_sim_core import (
    category_label,
    default_count_for_category,
    generate_selection,
    group_eligible_records,
    record_key,
    subtype_label,
    type_rule_key,
    validate_probability_rules,
)
from random_sim_exporter import (
    ExportPreflightError,
    build_export_plan,
    export_collection,
    package_identity,
)
from random_sim_usage import (
    default_usage_registry_path,
    dependency_usages_for_record,
    load_usage_registry,
    register_export as register_random_sim_export,
    save_usage_registry,
    selected_usages_for_record,
    sim_history,
    usages_for_record,
    usages_for_sim,
)

APP_NAME = "CAS Curator"
_DEFAULT_EDITION = "PUBLIC" if (Path(__file__).resolve().parent / "PUBLIC_EDITION").exists() else "DEV"
APP_EDITION = str(os.environ.get("CAS_CURATOR_EDITION", _DEFAULT_EDITION) or _DEFAULT_EDITION).strip().upper()
if APP_EDITION not in {"DEV", "PUBLIC"}:
    APP_EDITION = "DEV"
IS_PUBLIC_BUILD = APP_EDITION == "PUBLIC"
IS_DEV_BUILD = not IS_PUBLIC_BUILD
VERSION = (
    "v1.0.1"
    if IS_PUBLIC_BUILD else
    "v0.4.63 - DEV Scanner Log + Table Panel Resize TEST"
)



from zip_tools import (
    _zip_preview_plan,
    _zip_conflicts,
    _extract_zip_plan,
    _preflight_scan_zip_plans,
    _extract_scan_zip_plans,
)
from sims3pack_tools import (
    inspect_sims3pack,
    package_hashes as _sims3pack_package_hashes,
    annotate_package_alternatives,
    preflight_missing_cas_batch,
    extract_missing_cas_batch,
)

def _asset_path(name: str) -> Path:
    """Return a bundled asset path in source and PyInstaller onedir builds."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / "assets" / name

PUBLIC_THEME_ORDER = [
    "Cherry Blossom Pink",
    "Strawberry Milk",
    "Plum Ice",
    "Berry Cloud",
    "Aqua Depth",
    "Navy Bloom",
    "Blue Grey Rose",
    "Midnight Steel",
    "Carbon Blue",
    "Deep Ocean",
    "Storm Purple",
    "Retro Terminal",
    "Black Cherry",
    "Dark Teal",
    "Night Plum",
]
PUBLIC_THEME_NAMES = set(PUBLIC_THEME_ORDER)

EDITABLE_FIELDS = {"Image", "Status", "Creator", "File Location", "CC Magic Profiles", "Notes"}
# Public Status choices are deliberately separate from the legacy Settings-sheet
# values. The workbook is not migrated here; existing values remain untouched
# until a previewed, backed-up migration is implemented.
PUBLIC_STATUS_OPTIONS = ["In Game Use", "To Keep", "To Review"]

# The normal interface stays deliberately compact. Advanced and Experimental are
# opt-in display layers only: they never change scanner logic, Routing decisions,
# or package files. The workbook data remains intact when either layer is off.
PUBLIC_BASIC_FIELDS = {
    "Package",
    "Status",
    "Creator",
    "Health Status",
    "CC Type",
    "Risk Score",
    "Highest Polys",
    "Age",
    "Gender",
    "File Location",
    "CC Magic Import",
    "CC Magic Profiles",
    "Merge Profiles",
    "Notes",
}
PUBLIC_ADVANCED_FIELDS = {
    "Health Issues",
    "CASP Instance IDs",
    "Risk",
    "HQ Warning",
    "Poly Warning",
    "Biggest Texture",
    "Duplicate Files",
    "Last Scanned",
    "Package Hash",
}
PUBLIC_EXPERIMENTAL_FIELDS = {
    "Subtype",
    "Skin Type",
    "EA Resource Status",
    "Outfit Category",
}
# Package Format is deliberately normal, understandable side-panel metadata,
# but it is too sparse to spend a Table column on. Keep it out of the column
# visibility layer and surface it in Overview when a value exists.
PUBLIC_OVERVIEW_ONLY_FIELDS = {"Package Format"}
# Live display-only fields that are derived from the current filesystem/settings,
# never written to the MDB.
VIRTUAL_LIBRARY_FIELDS = {"CC Magic Import"}
_ADVANCED_FEATURES_ENABLED = False
_EXPERIMENTAL_FEATURES_ENABLED = False
DIRTY_ROLE = 257  # Qt.UserRole + 1

def _public_field_visible(field: str) -> bool:
    if field == "Library Location":
        return False
    if field in ("#", "Thumbnail"):
        return True
    if IS_PUBLIC_BUILD:
        if field in PUBLIC_BASIC_FIELDS:
            return True
        if _ADVANCED_FEATURES_ENABLED and field in PUBLIC_ADVANCED_FIELDS:
            return True
        if _EXPERIMENTAL_FEATURES_ENABLED and field in PUBLIC_EXPERIMENTAL_FIELDS:
            return True
        return False
    return True

def _available_public_columns_for_mdb(mdb) -> List[str]:
    """Return stored + live virtual Library columns in a stable user-facing order."""
    stored = list(getattr(mdb, "headers", []) or [])
    result = []
    for field in stored + ["CC Magic Import"]:
        if field not in result and _public_field_visible(field):
            result.append(field)
    return result


_BASE_DEFAULT_COLUMNS = [
    "#", "Package", "Thumbnail", "Creator", "Status", "Health Status", "Health Issues",
    "CC Type", "Subtype", "CASP Instance IDs", "Risk", "Risk Score", "Highest Polys",
    "File Location", "CC Magic Import", "CC Magic Profiles", "Merge Profiles", "Notes", "Last Scanned",
]
DEFAULT_COLUMNS = [field for field in _BASE_DEFAULT_COLUMNS if _public_field_visible(field)]

_BASE_SEARCH_FIELDS = [
    "Package", "Status", "Health Status", "Health Issues", "CC Type",
    "Subtype", "Age", "Gender", "Outfit Category", "CASP Instance IDs", "Risk", "Risk Score", "EA Resource Status",
    "Creator", "File Location", "CC Magic Import", "CC Magic Profiles", "Merge Profiles", "Notes", "Package Hash", "Normalized Package",
    "Original Source Path",
]
SEARCH_FIELDS = [
    field for field in _BASE_SEARCH_FIELDS
    if _public_field_visible(field)
]

_BASE_SUMMARY_FIELDS = [
    # Normal package summary first. CC Type / Highest Polys are skipped below
    # because they already live in the top at-a-glance badges.
    "Health Status", "Risk Score", "Age", "Gender", "Package Format",
    "Creator", "File Location", "CC Magic Import", "CC Magic Profiles", "Merge Profiles", "Notes",
    # Opt-in technical / provisional details follow when their layer is enabled.
    "Health Issues", "Risk", "HQ Warning", "Poly Warning", "Biggest Texture",
    "CASP Instance IDs", "Duplicate Files", "Last Scanned", "Package Hash",
    "CC Type", "Highest Polys", "Subtype", "EA Resource Status", "Skin Type", "Outfit Category",
    "Original Source Path",
]
SUMMARY_FIELDS = [field for field in _BASE_SUMMARY_FIELDS if _public_field_visible(field)]

# Permanent quick filters stay deliberately small. More specific metadata is
# available through + Add Filter… so the top Library panel does not grow a new
# control for every MDB column.
_BASE_FILTER_FIELDS = [
    ("Status", "Status"),
    ("CC Type", "CC Type"),
    ("Creator", "Creator"),
    ("CC Magic Profiles", "Profile"),
    # Kept instantiated for DEV/backwards-compatible internal use, but PUBLIC
    # does not place these on the permanent filter bar.
    ("Health Status", "Health"),
    ("Risk", "Risk"),
    ("File Location", "File Location"),
    ("CC Magic Import", "CC Magic"),
]
FILTER_FIELDS = [item for item in _BASE_FILTER_FIELDS if _public_field_visible(item[0])]

# Quick table column folds. These are only visibility shortcuts: they do not
# alter workbook data, scanner output or the user's fine-grained Columns menu.
COLUMN_FOLD_GROUPS = {
    "CAS": [
        "Age", "Gender", "Outfit Category", "Subtype", "Skin Type",
        "EA Resource Status", "CASP Instance IDs",
    ],
    "Health": [
        "Health Status", "Health Issues", "Risk", "Risk Score",
        "HQ Warning", "Poly Warning", "Biggest Texture", "Highest Polys",
    ],
    "Organise": [
        "Creator", "File Location", "CC Magic Import", "CC Magic Profiles", "Merge Profiles",
        "Notes", "Last Scanned", "Package Hash",
    ],
}


DISPLAY_FIELD_NAMES = {"CC Magic Profiles": "Profiles", "Merge Profiles": "Merge Profiles", "CC Magic Import": "CC Magic Import"}


def _display_field_name(field: str) -> str:
    return DISPLAY_FIELD_NAMES.get(field, field)


def _subtype_basis_label(record: Dict[str, str]) -> str:
    """Explain the evidence level behind the current stored Experimental Subtype.

    The shared scanner core currently derives Build/Buy subtype labels from the
    filename, while Scripted Object subtypes come from package/script evidence.
    This helper is display-only and never changes stored workbook values.
    """
    subtype = str(record.get("Subtype", "") or "").strip()
    if not subtype:
        return ""
    if subtype in {"Embedded Script", "External Script"}:
        return "package/script evidence"
    return "filename hint"


def _skin_type_basis_label(record: Dict[str, str]) -> str:
    """Explain the evidence behind the stored Experimental Skin Type.

    A TONE resource directly supports CC Type = Skin, but the current scanner's
    Default Replacement / Non-Default Skin Type label is still derived from the
    filename.  When CC Type is not Skin, the value comes from the scanner's
    broader possible-skin heuristic plus the same filename clues.  This helper
    is display-only and never changes stored workbook values or Routing.
    """
    skin_type = str(record.get("Skin Type", "") or "").strip()
    if not skin_type:
        return ""
    cc_type = str(record.get("CC Type", "") or "").strip()
    if cc_type == "Skin":
        return "TONE confirms Skin; default/non-default label from filename"
    return "possible-skin heuristic; type label from filename"


def _outfit_category_basis_label(record: Dict[str, str]) -> str:
    """Explain the evidence behind Experimental Outfit Category."""
    if not str(record.get("Outfit Category", "") or "").strip():
        return ""
    return "direct CASP ClothingCategory flags"


def _split_cc_magic_profiles(value: str) -> List[str]:
    """Return a clean, case-insensitively de-duplicated list of profiles."""
    text = str(value or "").replace("\r", "\n")
    raw_parts = []
    for line in text.split("\n"):
        raw_parts.extend(line.split(";"))
    result = []
    seen = set()
    for part in raw_parts:
        clean = part.strip()
        if not clean:
            continue
        key = clean.casefold()
        if key in seen:
            continue
        seen.add(key)
        result.append(clean)
    return result


def _normalise_cc_magic_profiles(value: str) -> str:
    """Store an editable multi-profile list as a clean semicolon-separated string."""
    return "; ".join(_split_cc_magic_profiles(value))


def _unique_cc_magic_profiles(records) -> List[str]:
    """Collect individual profile names from MDB records, not combined profile strings."""
    result = []
    seen = set()
    for rec in records:
        for profile in _split_cc_magic_profiles(rec.get("CC Magic Profiles", "")):
            key = profile.casefold()
            if key in seen:
                continue
            seen.add(key)
            result.append(profile)
    return sorted(result, key=str.casefold)


def _cc_type_key(value: str) -> str:
    """Case-insensitive key for CC Type filters without changing MDB data."""
    return str(value or "").strip().casefold()


def _cc_type_label(value: str) -> str:
    """Clean display label for CC Type values used only by filter controls."""
    clean = str(value or "").strip()
    if clean.casefold() == "unknown":
        return "Unknown"
    return clean


def _unique_cc_type_values(records) -> list[str]:
    """Collapse case-only CC Type duplicates, e.g. UNKNOWN + unknown -> Unknown."""
    by_key = {}
    for rec in records:
        raw = rec.get("CC Type", "")
        key = _cc_type_key(raw)
        if not key:
            continue
        label = _cc_type_label(raw)
        # Unknown gets a fixed friendly label. Other values keep the first MDB spelling.
        if key not in by_key or key == "unknown":
            by_key[key] = label
    return sorted(by_key.values(), key=str.casefold)

COLUMN_WIDTHS = {
    "#": 54,
    "Package": 340,
    "Creator": 160,
    "Thumbnail": 86,
    "Status": 100,
    "Health Status": 112,
    "Health Issues": 300,
    "CC Type": 145,
    "Subtype": 145,
    "EA Resource Status": 250,
    "Outfit Category": 260,
    "CASP Instance IDs": 240,
    "Risk": 95,
    "Risk Score": 90,
    "File Location": 190,
    "CC Magic Import": 140,
    "CC Magic Profiles": 220,
    "Merge Profiles": 220,
    "Notes": 220,
    "Last Scanned": 175,
}

THEMES = {
    "Shades of Grey": {
        "accent": "#8F8F8F", "accent_soft": "#ECECEC", "selection": "#7E93A5",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F1F1F1",
        "thumb": "#747474", "thumb_border": "#E2E2E2", "info_bg": "#ECECEC",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#303030", "muted": "#666666",
    },
    "Rose Grey": {
        "accent": "#D89DBA", "accent_soft": "#F5E3EA", "selection": "#7FA2C0",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F3F3F3",
        "thumb": "#777777", "thumb_border": "#E2E2E2", "info_bg": "#ECECEC",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#30343A", "muted": "#69717A",
    },
    "Cherry Blossom Pink": {
        "accent": "#C97591", "accent_soft": "#F2AEBC", "selection": "#86A7C2",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F6DFE6",
        "thumb": "#777777", "thumb_border": "#E2E2E2", "info_bg": "#DFA7B6",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#352F31", "muted": "#786C70",
    },
    "Strawberry Milk": {
        "accent": "#D96F91", "accent_soft": "#F5C4D2", "selection": "#839FB7",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#FFF1E8",
        "thumb": "#767676", "thumb_border": "#E2E2E2", "info_bg": "#E7B2C0",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#352F32", "muted": "#74686D",
    },
    "Lavender Cream": {
        "accent": "#B492D3", "accent_soft": "#E5D4F1", "selection": "#789BB7",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F3EEF7",
        "thumb": "#757575", "thumb_border": "#E2E2E2", "info_bg": "#EEE8F2",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#332F37", "muted": "#6E6874",
    },
    "Jellyfish Glow": {
        "accent": "#7997E6", "accent_soft": "#F0B7CF", "selection": "#206ABC",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F1ECFA",
        "thumb": "#707070", "thumb_border": "#E2E2E2", "info_bg": "#F7E4EB",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#2F3150", "muted": "#666A8A",
    },
    "Sunset Palace": {
        "accent": "#AA74A0", "accent_soft": "#D8C8E6", "selection": "#6978B3",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F2EEF8",
        "thumb": "#747474", "thumb_border": "#E2E2E2", "info_bg": "#F1EDF4",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#343552", "muted": "#6B687D",
    },
    "Celestial Lilac": {
        "accent": "#4E7AB1", "accent_soft": "#CEB5D4", "selection": "#5E789B",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EDF2F8",
        "thumb": "#747474", "thumb_border": "#E2E2E2", "info_bg": "#E9EFF5",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#2E3850", "muted": "#667287",
    },
    "Plum Ice": {
        "accent": "#5B1E32", "accent_soft": "#D9B8C4", "selection": "#88A8C2",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F2ECEE",
        "thumb": "#737373", "thumb_border": "#E2E2E2", "info_bg": "#EEE9EB",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#3D2730", "muted": "#75646B",
    },
    "Lotus Pond": {
        "accent": "#839958", "accent_soft": "#E3E7C7", "selection": "#5B8190",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F7F4D5",
        "thumb": "#737373", "thumb_border": "#E2E2E2", "info_bg": "#F2D8DE",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#314238", "muted": "#6A7468",
    },
    "Sakura Tea": {
        "accent": "#EC9C9D", "accent_soft": "#F2CFCA", "selection": "#8A98A9",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F5ECE6",
        "thumb": "#737373", "thumb_border": "#E2E2E2", "info_bg": "#F0E6E1",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#4A382E", "muted": "#766A62",
    },
    "Cafe Rose": {
        "accent": "#B46C72", "accent_soft": "#FFC8CB", "selection": "#6E8AA7",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F5E6D1",
        "thumb": "#737373", "thumb_border": "#E2E2E2", "info_bg": "#F0E7DD",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#4A3134", "muted": "#76656A",
    },
    "Berry Cloud": {
        "accent": "#B35A8A", "accent_soft": "#F6B7C7", "selection": "#7880AE",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EEEAF5",
        "thumb": "#737373", "thumb_border": "#E2E2E2", "info_bg": "#EAE6F0",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#403149", "muted": "#6F6575",
    },
    "Soft Blush Slate": {
        "accent": "#DB7F8E", "accent_soft": "#FFDDDA", "selection": "#7C94A5",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EEEAEA",
        "thumb": "#737373", "thumb_border": "#E2E2E2", "info_bg": "#EAE6E7",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#40383A", "muted": "#71696B",
    },
    "Wisteria Aqua": {
        "accent": "#80A1D4", "accent_soft": "#C0B9DD", "selection": "#75AFC1",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F7F4EA",
        "thumb": "#737373", "thumb_border": "#E2E2E2", "info_bg": "#EEEAF4",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#334256", "muted": "#68778A",
    },
    "Sea Glass": {
        "accent": "#4682B4", "accent_soft": "#B0E0E6", "selection": "#5F9EA0",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#ECF7F8",
        "thumb": "#70777A", "thumb_border": "#E2E2E2", "info_bg": "#D8F3F6",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#28434D", "muted": "#627A83",
    },
    "Blue Coral": {
        "accent": "#42899B", "accent_soft": "#94C4C1", "selection": "#E8877B",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EDF7F5",
        "thumb": "#707575", "thumb_border": "#E2E2E2", "info_bg": "#F9C7BE",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#28434A", "muted": "#65777A",
    },
    "Aqua Depth": {
        "accent": "#006C83", "accent_soft": "#C7EEF1", "selection": "#00A7B5",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#E7F7F8",
        "thumb": "#687477", "thumb_border": "#E2E2E2", "info_bg": "#C7EEF1",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#173E48", "muted": "#5E747A",
    },
    "Periwinkle Mist": {
        "accent": "#6C7EB8", "accent_soft": "#C8BFE7", "selection": "#8EA9DB",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F0EDF8",
        "thumb": "#716F78", "thumb_border": "#E2E2E2", "info_bg": "#D9D8F0",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#343B5A", "muted": "#6C6F86",
    },
    "Atlantic": {
        "accent": "#174E8A", "accent_soft": "#C7D8F0", "selection": "#4682B4",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#E6EEF7",
        "thumb": "#6A7179", "thumb_border": "#E2E2E2", "info_bg": "#D5E2F2",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#263B55", "muted": "#63748A",
    },
    "Ice Lagoon": {
        "accent": "#3A8795", "accent_soft": "#D8F0F3", "selection": "#5DA9B3",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F2FBFC",
        "thumb": "#697577", "thumb_border": "#E2E2E2", "info_bg": "#D8F0F3",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#26454B", "muted": "#637B80",
    },
    "Navy Bloom": {
        "accent": "#5D84B8", "accent_soft": "#BFD0E4", "selection": "#7EA6D9",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#E8EEF6",
        "thumb": "#515A66", "thumb_border": "#D7DDE5", "info_bg": "#2C4F73",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#253344", "muted": "#69798D",
    },
    "Jelly Blue Pink": {
        "accent": "#206ABC", "accent_soft": "#CAA9F3", "selection": "#D989AE",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F1EEFA",
        "thumb": "#6C6E77", "thumb_border": "#E2E2E2", "info_bg": "#F2C2DA",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#303954", "muted": "#6B718A",
    },
    "Seafoam Pearl": {
        "accent": "#3E7C7B", "accent_soft": "#BFE8E3", "selection": "#20B2AA",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EAF7F5",
        "thumb": "#6C7472", "thumb_border": "#E2E2E2", "info_bg": "#D2ECEA",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#294542", "muted": "#647B77",
    },
    "Blue Grey Rose": {
        "accent": "#6E89A3", "accent_soft": "#CFD9E3", "selection": "#D391AC",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EFF2F5",
        "thumb": "#6D7278", "thumb_border": "#E2E2E2", "info_bg": "#F2CEDA",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#33404D", "muted": "#697683",
    },
    "Blue Porcelain": {
        "accent": "#5E8FC3", "accent_soft": "#DCEAF6", "selection": "#7AA6D0",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F7FBFE",
        "thumb": "#70767B", "thumb_border": "#E2E2E2", "info_bg": "#EDF5FB",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#29435F", "muted": "#667B90",
    },
    "Midnight Steel": {
        "accent": "#172033", "accent_soft": "#DDE3EA", "selection": "#52657A",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EFF2F5",
        "thumb": "#70757A", "thumb_border": "#E2E2E2", "info_bg": "#E9EDF1",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#252A31", "muted": "#65707B",
    },
    "Carbon Blue": {
        "accent": "#3D566E", "accent_soft": "#DCE4EA", "selection": "#657F97",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EFF3F6",
        "thumb": "#6F7377", "thumb_border": "#E2E2E2", "info_bg": "#E9EEF2",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#2A3035", "muted": "#66727B",
    },
    "Espresso Copper": {
        "accent": "#855C45", "accent_soft": "#E9DDD5", "selection": "#9A745D",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F4EFEB",
        "thumb": "#74716F", "thumb_border": "#E2E2E2", "info_bg": "#EEE8E3",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#3A302B", "muted": "#746961",
    },
    "Deep Ocean": {
        "accent": "#176071", "accent_soft": "#D7E6E9", "selection": "#4F8792",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EDF4F5",
        "thumb": "#6F7475", "thumb_border": "#E2E2E2", "info_bg": "#E7EFF0",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#263A3F", "muted": "#62777B",
    },
    "Burgundy Ink": {
        "accent": "#653944", "accent_soft": "#E8DDE0", "selection": "#95606C",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F4EFF1",
        "thumb": "#737071", "thumb_border": "#E2E2E2", "info_bg": "#EEE8EA",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#3B2E32", "muted": "#75666B",
    },
    "Storm Purple": {
        "accent": "#50506D", "accent_soft": "#E0DFE8", "selection": "#777994",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F1F0F5",
        "thumb": "#707073", "thumb_border": "#E2E2E2", "info_bg": "#EAE9EF",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#31313F", "muted": "#6D6D7C",
    },
    "Retro Terminal": {
        "accent": "#345C50", "accent_soft": "#DCE7E2", "selection": "#76A18E",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EEF3F0",
        "thumb": "#6E7371", "thumb_border": "#E2E2E2", "info_bg": "#E7EDEA",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#283530", "muted": "#65736D",
    },

    "Obsidian Blue": {
        "accent": "#637C99", "accent_soft": "#C9D3DE", "selection": "#8197B0",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EEF1F5",
        "thumb": "#686E76", "thumb_border": "#E2E2E2", "info_bg": "#3D5068",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#29323D", "muted": "#697788",
    },
    "Black Cherry": {
        "accent": "#8A5368", "accent_soft": "#D8C2CB", "selection": "#A66D83",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F2ECEF",
        "thumb": "#6D686A", "thumb_border": "#E2E2E2", "info_bg": "#63394A",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#382B30", "muted": "#76666D",
    },
    "Dark Teal": {
        # Release-sync: keep deep teal surfaces but give controls a cooler aqua identity.
        "accent": "#4F8E91", "accent_soft": "#C9DEDD", "selection": "#73AAA8",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#EDF3F2",
        "thumb": "#676F6E", "thumb_border": "#E2E2E2", "info_bg": "#2F6663",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#273735", "muted": "#667571",
    },
    "Charcoal Copper": {
        "accent": "#A06F52", "accent_soft": "#DDC9BD", "selection": "#BE8765",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F2EEEC",
        "thumb": "#6D6967", "thumb_border": "#E2E2E2", "info_bg": "#5A4941",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#382F2B", "muted": "#746B66",
    },
    "Night Plum": {
        "accent": "#826E93", "accent_soft": "#D1C7D8", "selection": "#9B86AA",
        "line": "#E3E3E3", "line_alt": "#F6F6F6", "panel": "#FFFFFF", "panel_alt": "#F0EDF2",
        "thumb": "#6B686E", "thumb_border": "#E2E2E2", "info_bg": "#574863",
        "border": "#C8C8C8", "header_border": "#B6B6B6", "text": "#342D39", "muted": "#716A77",
    },
}

THEME_GROUPS = {
    "Neutral": ["Shades of Grey", "Carbon Blue"],
    "Pink & Peach": ["Rose Grey", "Cherry Blossom Pink", "Strawberry Milk", "Berry Cloud", "Sakura Tea", "Cafe Rose", "Soft Blush Slate"],
    "Blue & Sea": ["Sea Glass", "Aqua Depth", "Atlantic", "Ice Lagoon", "Seafoam Pearl", "Blue Porcelain"],
    "Blue + Pink": ["Blue Coral", "Jelly Blue Pink", "Blue Grey Rose"],
    "Blue & Purple": ["Lavender Cream", "Jellyfish Glow", "Sunset Palace", "Celestial Lilac", "Wisteria Aqua", "Periwinkle Mist"],
    "Nature": ["Lotus Pond"],
    "Dark & Cool": ["Storm Purple", "Retro Terminal", "Midnight Steel", "Deep Ocean", "Burgundy Ink", "Obsidian Blue", "Black Cherry", "Dark Teal", "Night Plum", "Plum Ice", "Navy Bloom"],
    "Warm & Bold": ["Espresso Copper", "Charcoal Copper"],
}


def _mix_hex(a: str, b: str, amount: float) -> str:
    """Blend colour a toward b by amount (0..1)."""
    amount = max(0.0, min(1.0, amount))

    def rgb(value: str):
        value = value.lstrip("#")
        return tuple(int(value[i:i+2], 16) for i in (0, 2, 4))

    ar, ag, ab = rgb(a)
    br, bg, bb = rgb(b)
    vals = [round(x + (y - x) * amount) for x, y in ((ar, br), (ag, bg), (ab, bb))]
    return "#%02X%02X%02X" % tuple(vals)


def _luma_hex(value: str) -> float:
    value = value.lstrip("#")
    r, g, b = (int(value[i:i+2], 16) for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _button_text_for(bg: str, dark: str = "#2F2F2F", light: str = "#FFFFFF") -> str:
    return light if _luma_hex(bg) < 148 else dark


def _normalize_hex(value: str) -> Optional[str]:
    value = value.strip().lstrip("#")
    if re.fullmatch(r"[0-9A-Fa-f]{6}", value):
        return "#" + value.upper()
    if re.fullmatch(r"[0-9A-Fa-f]{3}", value):
        return "#" + "".join(ch * 2 for ch in value).upper()
    return None


def _hex_rgb(value: str):
    value = value.lstrip("#")
    return tuple(int(value[i:i+2], 16) for i in (0, 2, 4))


def _colour_saturation(value: str) -> float:
    r, g, b = (x / 255.0 for x in _hex_rgb(value))
    return colorsys.rgb_to_hsv(r, g, b)[1]


def _colour_distance(a: str, b: str) -> float:
    ar, ag, ab = _hex_rgb(a)
    br, bg, bb = _hex_rgb(b)
    return ((ar-br)**2 + (ag-bg)**2 + (ab-bb)**2) ** 0.5


def parse_palette_text(raw: str) -> List[str]:
    """Extract colours from HEX, GIMP GPL / PAL decimal RGB, and Coolors URLs.

    Parse decimal RGB lines *before* generic HEX: a GPL line such as
    ``201 88 126 Cherry`` must not become fake 3-digit hex colours.
    """
    colours: List[str] = []
    hex_pattern = re.compile(
        r"(?<![0-9A-Za-z_])(?:#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{3})|([0-9A-Fa-f]{6}))(?![0-9A-Za-z_])"
    )
    for line in raw.splitlines():
        rgb_line = re.match(r"^\s*(\d{1,3})\s+(\d{1,3})\s+(\d{1,3})(?:\s+.*)?$", line)
        if rgb_line:
            rgb = tuple(int(value) for value in rgb_line.groups()[:3])
            if all(0 <= value <= 255 for value in rgb):
                colour = "#%02X%02X%02X" % rgb
                if colour not in colours:
                    colours.append(colour)
                continue
        for match in hex_pattern.finditer(line):
            token = match.group(1) or match.group(2)
            colour = _normalize_hex(token)
            if colour and colour not in colours:
                colours.append(colour)
    return colours[:48]


def build_theme_from_palette(name: str, colours: List[str]) -> Dict[str, object]:
    """Turn a downloaded colour palette into a usable private DEV theme."""
    colours = [c for c in (_normalize_hex(x) for x in colours) if c]
    colours = list(dict.fromkeys(colours))
    if len(colours) < 2:
        raise ValueError("A palette needs at least two colours.")

    ranked = sorted(colours, key=_luma_hex)
    darkest, lightest = ranked[0], ranked[-1]
    average_luma = sum(_luma_hex(c) for c in colours) / len(colours)
    dark_mode = average_luma < 130 and _luma_hex(lightest) < 225

    # Prefer a colourful accent rather than accidentally picking white/grey.
    accent_candidates = sorted(
        colours,
        key=lambda c: (_colour_saturation(c) * 180) + abs(_luma_hex(c) - 135),
        reverse=True,
    )
    accent = accent_candidates[0]
    accent2 = max((c for c in colours if c != accent), key=lambda c: _colour_distance(c, accent), default=accent)

    if dark_mode:
        top_fill = darkest
        side_fill = _mix_hex(darkest, accent, 0.20)
        header_fill = _mix_hex(side_fill, lightest, 0.24)
        info_fill = ranked[min(len(ranked)-1, max(1, len(ranked)//2))]
        if _luma_hex(info_fill) < 90:
            info_fill = _mix_hex(info_fill, lightest, 0.30)
        filter_button = lightest if _luma_hex(lightest) < 230 else _mix_hex(lightest, accent, 0.22)
        scrollbar = accent if 55 <= _luma_hex(accent) <= 205 else accent2
        thumb_fill = _mix_hex(darkest, "#555B63", 0.35)
        info_border = "#F0F0F0"
    else:
        mids = sorted(colours, key=lambda c: abs(_luma_hex(c) - 185))
        top_fill = mids[0]
        if _luma_hex(top_fill) < 125:
            top_fill = _mix_hex(top_fill, lightest, 0.48)
        side_fill = _mix_hex(top_fill, lightest, 0.42)
        header_fill = _mix_hex(side_fill, lightest, 0.28)
        info_candidates = sorted(colours, key=lambda c: abs(_luma_hex(c) - 210))
        info_fill = info_candidates[0]
        if _luma_hex(info_fill) < 150:
            info_fill = _mix_hex(info_fill, lightest, 0.50)
        filter_button = accent
        if _luma_hex(filter_button) < 75:
            filter_button = _mix_hex(filter_button, lightest, 0.32)
        elif _luma_hex(filter_button) > 225:
            filter_button = _mix_hex(filter_button, darkest, 0.20)
        scrollbar = accent2 if _colour_saturation(accent2) >= 0.18 else accent
        if _luma_hex(scrollbar) > 210:
            scrollbar = _mix_hex(scrollbar, darkest, 0.35)
        thumb_fill = side_fill
        info_border = "#FFFFFF" if _luma_hex(info_fill) < 225 else "#D7DCE1"

    scrollbar_hover = _mix_hex(scrollbar, lightest if _luma_hex(scrollbar) < 150 else darkest, 0.18)
    selection = accent2

    # Base roles keep the rest of the app readable. Direct UI roles below preserve
    # more of the imported palette than the normal built-in theme formula does.
    theme: Dict[str, object] = {
        "accent": scrollbar,
        "accent_soft": lightest,
        "selection": selection,
        "line": "#E3E3E3",
        "line_alt": "#F6F6F6",
        "panel": "#FFFFFF",
        "panel_alt": _mix_hex(lightest, "#FFFFFF", 0.45),
        "thumb": "#70757A",
        "thumb_border": "#E2E2E2",
        "info_bg": info_fill,
        "border": "#C8C8C8",
        "header_border": "#B6B6B6",
        "text": "#28323A",
        "muted": "#66727B",
        "_custom": True,
        "_palette": colours,
        "ui_top_fill": top_fill,
        "ui_side_fill": side_fill,
        "ui_header_fill": header_fill,
        "ui_info_fill": info_fill,
        "ui_filter_button": filter_button,
        "ui_scrollbar_color": scrollbar,
        "ui_scrollbar_hover": scrollbar_hover,
        "ui_thumb_fill": thumb_fill,
        "ui_info_card_border": info_border,
    }
    return theme


class LoadWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, path: Path):
        super().__init__()
        self.path = path

    def run(self):
        try:
            self.finished.emit(MDBWorkbook().load(self.path))
        except Exception as exc:
            self.failed.emit(str(exc))


class LibraryTableModel(QAbstractTableModel):
    manualFieldEdited = Signal(int, str, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.records: List[Dict[str, str]] = []
        self.columns: List[str] = list(DEFAULT_COLUMNS)
        self.pending_changes: Dict[int, Dict[str, str]] = {}
        self.pending_deletions: set[int] = set()

    def set_pending_changes(self, pending_changes: Dict[int, Dict[str, str]]):
        # Keep the same shared dict as MainWindow so the table is only an editor/view,
        # never a second source of truth.
        self.pending_changes = pending_changes

    def set_pending_deletions(self, pending_deletions: set[int]):
        self.pending_deletions = pending_deletions

    def set_records(self, records: List[Dict[str, str]]):
        self.beginResetModel()
        self.records = records
        self.endResetModel()

    def set_columns(self, columns: List[str]):
        self.beginResetModel()
        self.columns = columns
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.records)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.columns)

    @staticmethod
    def _excel_row(rec: Dict[str, str]) -> int:
        try:
            return int(rec.get("_ExcelRow", "0") or 0)
        except (TypeError, ValueError):
            return 0

    def effective_value(self, row: int, field: str) -> str:
        if not (0 <= row < len(self.records)):
            return ""
        rec = self.records[row]
        excel_row = self._excel_row(rec)
        if excel_row:
            pending = self.pending_changes.get(excel_row, {})
            if field in pending:
                return pending[field]
        return rec.get(field, "")

    def is_row_dirty(self, row: int) -> bool:
        if not (0 <= row < len(self.records)):
            return False
        excel_row = self._excel_row(self.records[row])
        return bool(excel_row and self.pending_changes.get(excel_row))

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid() or index.row() >= len(self.records):
            return None
        rec = self.records[index.row()]
        field = self.columns[index.column()]
        value = self.effective_value(index.row(), field)

        if role == Qt.DisplayRole:
            if field == "#":
                return str(index.row() + 1)
            if "\n" in value:
                parts = value.splitlines()
                return parts[0] + (f"  (+{len(parts)-1})" if len(parts) > 1 else "")
            return value
        if role == Qt.EditRole and field in EDITABLE_FIELDS:
            return value
        if role == Qt.ToolTipRole:
            if field == "#" and self.is_row_dirty(index.row()):
                return "Unsaved manual-field changes in this row"
            if field in EDITABLE_FIELDS:
                suffix = "\n\nDouble-click to edit. Changes stay unsaved until you use Save Changes."
                return (value or "(empty)") + suffix
            if field != "#" and value:
                return value
        if role == Qt.UserRole:
            return rec
        if role == DIRTY_ROLE:
            return self.is_row_dirty(index.row())
        if role == Qt.TextAlignmentRole and field in ("#", "Thumbnail", "Risk Score"):
            return int(Qt.AlignCenter)
        if role == Qt.ForegroundRole and field == "#" and self.is_row_dirty(index.row()):
            return QColor("#C77719")
        return None

    def flags(self, index: QModelIndex):
        flags = super().flags(index)
        if index.isValid() and self.columns[index.column()] in EDITABLE_FIELDS:
            rec = self.records[index.row()] if index.row() < len(self.records) else {}
            try:
                excel_row = int(rec.get("_ExcelRow", "0") or 0)
            except (TypeError, ValueError):
                excel_row = 0
            if excel_row not in self.pending_deletions:
                flags |= Qt.ItemIsEditable
        return flags

    def setData(self, index: QModelIndex, value, role=Qt.EditRole):
        if role != Qt.EditRole or not index.isValid():
            return False
        field = self.columns[index.column()]
        if field not in EDITABLE_FIELDS:
            return False
        row = index.row()
        if not (0 <= row < len(self.records)):
            return False
        rec = self.records[row]
        excel_row = self._excel_row(rec)
        if not excel_row or excel_row in self.pending_deletions:
            return False
        value = "" if value is None else str(value)
        row_pending = dict(self.pending_changes.get(excel_row, {}))
        if value == rec.get(field, ""):
            row_pending.pop(field, None)
        else:
            row_pending[field] = value
        if row_pending:
            self.pending_changes[excel_row] = row_pending
        else:
            self.pending_changes.pop(excel_row, None)
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole, Qt.ToolTipRole, DIRTY_ROLE])
        if self.columns:
            first = self.index(row, 0)
            last = self.index(row, len(self.columns) - 1)
            self.dataChanged.emit(first, last, [Qt.DisplayRole, Qt.ToolTipRole, DIRTY_ROLE, Qt.ForegroundRole])
        self.manualFieldEdited.emit(row, field, value)
        return True

    def stage_fill_value(self, source_rows: List[int], field: str, value: str) -> List[int]:
        """Stage one exact value into several source-model rows.

        The caller must map any proxy/filter ordering before invoking this method.
        No worksheet write occurs here; edits stay in the shared pending_changes
        dictionary until MainWindow.save_changes() runs.
        """
        if field not in EDITABLE_FIELDS:
            return []
        value = "" if value is None else str(value)
        changed_rows: List[int] = []
        seen = set()
        for row in source_rows:
            if row in seen or not (0 <= row < len(self.records)):
                continue
            seen.add(row)
            rec = self.records[row]
            excel_row = self._excel_row(rec)
            if not excel_row or excel_row in self.pending_deletions:
                continue
            before = self.effective_value(row, field)
            if before == value:
                continue
            row_pending = dict(self.pending_changes.get(excel_row, {}))
            if value == rec.get(field, ""):
                row_pending.pop(field, None)
            else:
                row_pending[field] = value
            if row_pending:
                self.pending_changes[excel_row] = row_pending
            else:
                self.pending_changes.pop(excel_row, None)
            changed_rows.append(row)
        return changed_rows

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and 0 <= section < len(self.columns):
            field = self.columns[section]
            if role == Qt.DisplayRole:
                return _display_field_name(field)
            if role == Qt.ToolTipRole and field in EDITABLE_FIELDS:
                label = _display_field_name(field)
                return f"{label} — double-click a cell to edit"
        return super().headerData(section, orientation, role)

    def record(self, row: int) -> Dict[str, str]:
        if 0 <= row < len(self.records):
            return self.records[row]
        return {}

    def notify_row_changed(self, row: int):
        if not (0 <= row < len(self.records)) or not self.columns:
            return
        left = self.index(row, 0)
        right = self.index(row, len(self.columns) - 1)
        self.dataChanged.emit(left, right, [Qt.DisplayRole, Qt.EditRole, Qt.ToolTipRole, Qt.UserRole, DIRTY_ROLE, Qt.ForegroundRole])

    def notify_all_changed(self):
        if not self.records or not self.columns:
            return
        self.dataChanged.emit(
            self.index(0, 0),
            self.index(len(self.records) - 1, len(self.columns) - 1),
            [Qt.DisplayRole, Qt.EditRole, Qt.ToolTipRole, DIRTY_ROLE, Qt.ForegroundRole],
        )


class LibraryFilterProxy(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.search_text = ""
        self.filters: Dict[str, str] = {}
        self.column_filters: List[Dict[str, str]] = []
        self.gallery_quick_status = ""
        self.gallery_quick_cc_type = ""
        self.gallery_thumbnails_only = False
        self.gallery_pinned_only = False
        self.gallery_dirty_only = False
        self.gallery_pinned_rows = set()
        self.setDynamicSortFilter(True)

    def set_search_text(self, text: str):
        self.beginFilterChange()
        self.search_text = text.strip().casefold()
        self.endFilterChange(QSortFilterProxyModel.Direction.Rows)

    def set_filters(self, filters: Dict[str, str]):
        self.beginFilterChange()
        self.filters = filters
        self.endFilterChange(QSortFilterProxyModel.Direction.Rows)

    def set_column_filters(self, filters: List[Dict[str, str]]):
        """Set user-added filters from the universal + Add Filter… control."""
        self.beginFilterChange()
        self.column_filters = [dict(item) for item in (filters or [])]
        self.endFilterChange(QSortFilterProxyModel.Direction.Rows)

    @staticmethod
    def _matches_column_filter(field: str, value: str, condition: str, expected: str) -> bool:
        value_text = str(value or "")
        expected_text = str(expected or "")
        value_key = value_text.strip().casefold()
        expected_key = expected_text.strip().casefold()

        if condition == "is blank":
            return not value_text.strip()
        if condition == "is not blank":
            return bool(value_text.strip())

        # Exact CC Type comparisons use the same alias-normalisation as the
        # permanent quick filter, so legacy/canonical labels behave consistently.
        if field == "CC Type" and condition in {"is", "is not"}:
            matched = _cc_type_key(value_text) == _cc_type_key(expected_text)
        # A Profile row can contain several assigned profiles. Exact profile
        # filtering means membership, matching the permanent Profile filter.
        elif field == "CC Magic Profiles" and condition in {"is", "is not"}:
            profile_keys = {p.casefold() for p in _split_cc_magic_profiles(value_text)}
            matched = expected_key in profile_keys
        elif condition in {"is", "is not"}:
            matched = value_key == expected_key
        elif condition == "contains":
            matched = expected_key in value_key
        elif condition == "does not contain":
            matched = expected_key not in value_key
        else:
            matched = True

        return (not matched) if condition == "is not" else matched

    def set_gallery_quick_filters(
        self,
        status: str = "",
        cc_type: str = "",
        thumbnails_only: bool = False,
        pinned_only: bool = False,
        dirty_only: bool = False,
        pinned_rows=None,
    ):
        self.beginFilterChange()
        self.gallery_quick_status = status or ""
        self.gallery_quick_cc_type = cc_type or ""
        self.gallery_thumbnails_only = bool(thumbnails_only)
        self.gallery_pinned_only = bool(pinned_only)
        self.gallery_dirty_only = bool(dirty_only)
        self.gallery_pinned_rows = set(pinned_rows or [])
        self.endFilterChange(QSortFilterProxyModel.Direction.Rows)

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()
        if model is None:
            return True
        rec = model.record(source_row)
        for field, expected in self.filters.items():
            value = model.effective_value(source_row, field) if hasattr(model, "effective_value") else rec.get(field, "")
            if expected and expected != "All":
                if field == "CC Type":
                    if _cc_type_key(value) != _cc_type_key(expected):
                        return False
                elif field == "CC Magic Profiles":
                    profile_keys = {p.casefold() for p in _split_cc_magic_profiles(value)}
                    if str(expected).strip().casefold() not in profile_keys:
                        return False
                elif value != expected:
                    return False

        for item in self.column_filters:
            field = str(item.get("field", "") or "")
            condition = str(item.get("condition", "is") or "is")
            expected = str(item.get("value", "") or "")
            if not field:
                continue
            value = model.effective_value(source_row, field) if hasattr(model, "effective_value") else rec.get(field, "")
            if not self._matches_column_filter(field, value, condition, expected):
                return False

        # Gallery-only quick filters. They stack with the normal search/filter row.
        if self.gallery_quick_status:
            status_value = model.effective_value(source_row, "Status") if hasattr(model, "effective_value") else rec.get("Status", "")
            if status_value != self.gallery_quick_status:
                return False
        if self.gallery_quick_cc_type:
            type_value = rec.get("CC Type", "")
            if _cc_type_key(type_value) != _cc_type_key(self.gallery_quick_cc_type):
                return False
        if self.gallery_thumbnails_only and rec.get("Thumbnail", "") != "Yes":
            return False
        if self.gallery_pinned_only and source_row not in self.gallery_pinned_rows:
            return False
        if self.gallery_dirty_only and hasattr(model, "is_row_dirty") and not model.is_row_dirty(source_row):
            return False

        if self.search_text:
            values = [model.effective_value(source_row, field) if hasattr(model, "effective_value") else rec.get(field, "") for field in SEARCH_FIELDS]
            haystack = "\n".join(values).casefold()
            if self.search_text not in haystack:
                return False
        return True

    def lessThan(self, left: QModelIndex, right: QModelIndex):
        lv = self.sourceModel().data(left, Qt.DisplayRole) or ""
        rv = self.sourceModel().data(right, Qt.DisplayRole) or ""
        try:
            return float(lv) < float(rv)
        except (TypeError, ValueError):
            return str(lv).casefold() < str(rv).casefold()


class ArrowComboBox(QComboBox):
    """Combo box that paints its own chevron so Windows/Qt styles cannot hide it."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._arrow_color = "#3F3F3F"

    def set_arrow_color(self, color: str):
        self._arrow_color = color
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(self._arrow_color))
        cx = self.width() - 20
        cy = self.height() // 2 + 1
        triangle = QPolygon([
            QPoint(cx - 6, cy - 3),
            QPoint(cx + 6, cy - 3),
            QPoint(cx, cy + 4),
        ])
        painter.drawPolygon(triangle)
        painter.end()


class RowOutlineDelegate(QStyledItemDelegate):
    """Neutral rows, outline-only selection, dirty marker, and safe manual-field editors."""

    def __init__(self, table: QTableView, theme: Dict[str, str], options_provider=None, parent=None):
        super().__init__(parent)
        self.table = table
        self.theme = theme
        self.options_provider = options_provider

    def set_theme(self, theme: Dict[str, str]):
        self.theme = theme
        self.table.viewport().update()

    def set_options_provider(self, provider):
        self.options_provider = provider

    def paint(self, painter: QPainter, option, index: QModelIndex):
        from PySide6.QtWidgets import QStyleOptionViewItem
        selected = bool(option.state & QStyle.State_Selected)
        opt = option
        if selected:
            from PySide6.QtWidgets import QStyleOptionViewItem
            opt = QStyleOptionViewItem(option)
            opt.state &= ~QStyle.State_Selected

        painter.save()
        row_bg = (self.theme["line_alt"] if index.row() % 2 else
                  self.theme.get("row_main", "#FFFFFF"))
        painter.fillRect(option.rect, QColor(row_bg))
        painter.restore()
        if self.theme.get("_custom"):
            from PySide6.QtGui import QPalette
            opt = QStyleOptionViewItem(opt) if opt is option else opt
            opt.palette.setColor(QPalette.Text, QColor(_button_text_for(
                row_bg, dark="#171717", light="#FAFAFA")))

        super().paint(painter, opt, index)

        painter.save()
        line_pen = QPen(QColor(self.theme["line"]))
        line_pen.setWidth(1)
        painter.setPen(line_pen)
        painter.drawLine(option.rect.bottomLeft(), option.rect.bottomRight())
        if index.column() == 0 and bool(index.data(DIRTY_ROLE)):
            marker = QColor(self.theme.get("dirty_marker", "#E09A32"))
            painter.fillRect(option.rect.left() + 1, option.rect.top() + 4, 3, max(4, option.rect.height() - 8), marker)
        painter.restore()

    def createEditor(self, parent, option, index):
        field = index.model().headerData(index.column(), Qt.Horizontal, Qt.DisplayRole)
        if field == "Profiles":
            field = "CC Magic Profiles"
        if field in ("Status", "File Location", "CC Magic Profiles"):
            editor = QComboBox(parent)
            editor.addItem("")
            options = self.options_provider(field) if self.options_provider else []
            editor.addItems([str(v) for v in options if str(v)])
            if field == "CC Magic Profiles":
                editor.setToolTip("Choose an existing Profile. Free typing is disabled in the table.")
            editor.activated.connect(lambda *_: self._commit_combo(editor))
            return editor
        if field == "Notes":
            editor = QLineEdit(parent)
            editor.setPlaceholderText("Quick note…")
            editor.setToolTip("Quick inline note edit. Use the Edit tab for longer/multiline notes.")
            return editor
        return super().createEditor(parent, option, index)

    def _commit_combo(self, editor):
        self.commitData.emit(editor)
        self.closeEditor.emit(editor)

    def setEditorData(self, editor, index):
        field = index.model().headerData(index.column(), Qt.Horizontal, Qt.DisplayRole)
        if field == "Profiles":
            field = "CC Magic Profiles"
        value = index.model().data(index, Qt.EditRole) or ""
        if field in ("Status", "File Location", "CC Magic Profiles") and isinstance(editor, QComboBox):
            # Preserve an existing multi-profile value visually without enabling
            # free text. Choosing a list item replaces the cell with that Profile;
            # fill-down can then propagate the exact assignment to displayed rows.
            if value and editor.findText(value) < 0:
                editor.insertItem(1, value)
            editor.setCurrentText(value)
            return
        if field == "Notes" and isinstance(editor, QLineEdit):
            editor.setText(value)
            editor.selectAll()
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        field = model.headerData(index.column(), Qt.Horizontal, Qt.DisplayRole)
        if field == "Profiles":
            field = "CC Magic Profiles"
        if field in ("Status", "File Location") and isinstance(editor, QComboBox):
            model.setData(index, editor.currentText(), Qt.EditRole)
            return
        if field == "CC Magic Profiles" and isinstance(editor, QComboBox):
            model.setData(index, _normalise_cc_magic_profiles(editor.currentText()), Qt.EditRole)
            return
        if field == "Notes" and isinstance(editor, QLineEdit):
            model.setData(index, editor.text(), Qt.EditRole)
            return
        super().setModelData(editor, model, index)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)



class GalleryDelegate(QStyledItemDelegate):
    """Lightweight thumbnail-card renderer backed by the same filtered table model."""

    def __init__(self, proxy: QSortFilterProxyModel, source_model: LibraryTableModel, parent=None):
        super().__init__(parent)
        self.proxy = proxy
        self.source_model = source_model
        self.theme = {
            "card": "#FFFFFF",
            "border": "#D0D0D0",
            "image": "#F7F7F7",
            "text": "#303030",
            "muted": "#6F6F6F",
            "selection": "#7E93A5",
            "badge": "#ECECEC",
            "badge_text": "#303030",
            "dirty": "#E09A32",
        }
        self._pix_cache = {}
        self._cache_order = []
        self._cache_limit = 256

    def set_theme(self, theme: Dict[str, str]):
        self.theme.update(theme)
        if self.parent():
            self.parent().viewport().update()

    def clear_cache(self):
        self._pix_cache.clear()
        self._cache_order.clear()

    def sizeHint(self, option, index):
        return QSize(196, 228)

    def _cached_pixmap(self, path: Path, size: QSize) -> QPixmap:
        key = f"{path}|{size.width()}x{size.height()}"
        cached = self._pix_cache.get(key)
        if cached is not None:
            return cached
        pix = QPixmap(str(path))
        if pix.isNull():
            scaled = QPixmap()
        else:
            scaled = pix.scaled(size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self._pix_cache[key] = scaled
        self._cache_order.append(key)
        while len(self._cache_order) > self._cache_limit:
            oldest = self._cache_order.pop(0)
            self._pix_cache.pop(oldest, None)
        return scaled

    def paint(self, painter: QPainter, option, index: QModelIndex):
        if not index.isValid():
            return
        source_index = self.proxy.mapToSource(index)
        rec = self.source_model.record(source_index.row())
        if not rec:
            return

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)

        outer = option.rect.adjusted(5, 5, -5, -5)
        selected = bool(option.state & QStyle.State_Selected)
        border = self.theme["selection"] if selected else self.theme["border"]
        pen = QPen(QColor(border))
        pen.setWidth(2 if selected else 1)
        painter.setPen(pen)
        painter.setBrush(QColor(self.theme["card"]))
        painter.drawRoundedRect(outer, 9, 9)

        image_rect = QRect(outer.left() + 9, outer.top() + 9, outer.width() - 18, 136)
        painter.setPen(QPen(QColor(self.theme["border"]), 1))
        painter.setBrush(QColor(self.theme["image"]))
        painter.drawRoundedRect(image_rect, 7, 7)

        target = rec.get("_ThumbnailTarget", "")
        path = target_to_local_path(target) if target else None
        if path is not None and path.exists():
            pix = self._cached_pixmap(path, QSize(image_rect.width() - 12, image_rect.height() - 12))
            if not pix.isNull():
                x = image_rect.center().x() - pix.width() // 2
                y = image_rect.center().y() - pix.height() // 2
                painter.drawPixmap(x, y, pix)
            else:
                painter.setPen(QColor(self.theme["muted"]))
                painter.drawText(image_rect, Qt.AlignCenter, "Thumbnail unavailable")
        else:
            painter.setPen(QColor(self.theme["muted"]))
            painter.drawText(image_rect, Qt.AlignCenter, "No thumbnail")

        package = rec.get("Package", "Package") or "Package"
        name_rect = QRect(outer.left() + 9, image_rect.bottom() + 7, outer.width() - 18, 39)
        painter.setPen(QColor(self.theme["text"]))
        font = painter.font()
        font.setBold(True)
        painter.setFont(font)
        metrics = painter.fontMetrics()
        # Keep filenames compact and readable in a two-line card title.
        max_chars = max(24, int(name_rect.width() / max(5, metrics.averageCharWidth())) * 2)
        display_name = package if len(package) <= max_chars else package[: max_chars - 1] + "…"
        painter.drawText(name_rect, Qt.TextWordWrap | Qt.AlignTop | Qt.AlignLeft, display_name)

        health = rec.get("Health Status", "") or "—"
        cc_type = rec.get("CC Type", "") or "—"
        badge_rect = QRect(outer.left() + 9, outer.bottom() - 27, outer.width() - 18, 20)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(self.theme["badge"]))
        painter.drawRoundedRect(badge_rect, 7, 7)
        painter.setPen(QColor(self.theme["badge_text"]))
        small = painter.font()
        small.setBold(False)
        if small.pointSize() > 0:
            small.setPointSize(max(8, small.pointSize() - 1))
        painter.setFont(small)
        label = f"{health}  •  {cc_type}"
        label = painter.fontMetrics().elidedText(label, Qt.ElideRight, badge_rect.width() - 12)
        painter.drawText(badge_rect.adjusted(6, 0, -6, 0), Qt.AlignVCenter | Qt.AlignLeft, label)

        if self.source_model.is_row_dirty(source_index.row()):
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(self.theme["dirty"]))
            painter.drawEllipse(QRect(outer.right() - 15, outer.top() + 7, 8, 8))

        painter.restore()


class OutlineTableView(QTableView):
    fillRequested = Signal(QModelIndex, QModelIndex)

    def __init__(self, theme: Dict[str, str], parent=None):
        super().__init__(parent)
        self.theme = theme
        self._fill_drag_source = QModelIndex()
        self._fill_drag_target = QModelIndex()
        self.setMouseTracking(True)

    def set_theme(self, theme: Dict[str, str]):
        self.theme = theme
        self.viewport().update()

    def _fill_handle_rect(self, index: QModelIndex) -> QRect:
        if not index.isValid() or not (index.flags() & Qt.ItemIsEditable):
            return QRect()
        rect = self.visualRect(index)
        if not rect.isValid() or rect.width() < 10 or rect.height() < 10:
            return QRect()
        size = 8
        return QRect(rect.right() - size + 1, rect.bottom() - size + 1, size, size)

    def _current_fill_index(self) -> QModelIndex:
        index = self.currentIndex()
        if index.isValid() and (index.flags() & Qt.ItemIsEditable):
            return index
        return QModelIndex()

    def mousePressEvent(self, event):
        current = self._current_fill_index()
        handle = self._fill_handle_rect(current)
        if event.button() == Qt.LeftButton and handle.isValid() and handle.contains(event.position().toPoint()):
            self._fill_drag_source = current
            self._fill_drag_target = current
            self.viewport().update()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        point = event.position().toPoint()
        if self._fill_drag_source.isValid():
            # Keep long fills usable without changing the source mapping: gently
            # scroll when the pointer reaches the lower edge of the viewport.
            if point.y() >= self.viewport().height() - 18:
                bar = self.verticalScrollBar()
                bar.setValue(min(bar.maximum(), bar.value() + 1))
                point.setY(max(0, self.viewport().height() - 4))
            hovered = self.indexAt(point)
            if hovered.isValid() and hovered.row() >= self._fill_drag_source.row():
                candidate = self.model().index(hovered.row(), self._fill_drag_source.column())
                if candidate.isValid():
                    self._fill_drag_target = candidate
            self.viewport().update()
            event.accept()
            return

        current = self._current_fill_index()
        handle = self._fill_handle_rect(current)
        if handle.isValid() and handle.contains(point):
            self.setCursor(Qt.CrossCursor)
        else:
            self.unsetCursor()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._fill_drag_source.isValid():
            source = self._fill_drag_source
            target = self._fill_drag_target
            self._fill_drag_source = QModelIndex()
            self._fill_drag_target = QModelIndex()
            self.unsetCursor()
            self.viewport().update()
            event.accept()
            if target.isValid() and target.row() > source.row():
                self.fillRequested.emit(source, target)
            return
        super().mouseReleaseEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self.viewport())

        selection = self.selectionModel()
        if selection:
            rows = selection.selectedRows()
            if rows:
                pen = QPen(QColor(self.theme["selection"]))
                pen.setWidth(2)
                painter.setPen(pen)
                x1 = 0
                x2 = self.viewport().width() - 1
                for idx in rows:
                    row_rect = self.visualRect(idx)
                    if not row_rect.isValid():
                        continue
                    top = row_rect.top() + 1
                    bottom = row_rect.bottom() - 1
                    painter.drawLine(x1, top, x2, top)
                    painter.drawLine(x1, bottom, x2, bottom)
                    painter.drawLine(x1, top, x1, bottom)
                    painter.drawLine(x2, top, x2, bottom)

        current = self._current_fill_index()
        if current.isValid():
            handle = self._fill_handle_rect(current)
            if handle.isValid():
                painter.fillRect(handle, QColor(self.theme["selection"]))
                painter.setPen(QPen(QColor(self.theme["text"]), 1))
                painter.drawRect(handle.adjusted(0, 0, -1, -1))

        if self._fill_drag_source.isValid() and self._fill_drag_target.isValid():
            first = self.visualRect(self._fill_drag_source)
            last = self.visualRect(self._fill_drag_target)
            if first.isValid() and last.isValid():
                preview = first.united(last)
                pen = QPen(QColor(self.theme["selection"]))
                pen.setWidth(2)
                painter.setPen(pen)
                painter.setBrush(Qt.NoBrush)
                painter.drawRect(preview.adjusted(1, 1, -2, -2))

        painter.end()

class ResponsiveThumbnailLabel(QLabel):
    """Draw the complete original image inside the *actual visible widget box*.

    QLabel.setPixmap() gives QLabel a pixmap-dependent size hint. When an
    unusually long package title steals vertical space, Qt can keep a larger
    cached pixmap or request an oversized label and only display its top.
    Drawing at paint time removes that feedback loop entirely: each frame
    fits the source to the rectangle that the layout really allocated.
    """

    def __init__(self, text: str = "", parent=None):
        super().__init__(text, parent)
        self._source_pixmap = QPixmap()
        # Never let the image dimensions affect layout's minimum/suggested size.
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)

    def sizeHint(self):
        return QSize(240, 200)

    def minimumSizeHint(self):
        return QSize(100, 120)

    def set_thumbnail(self, pixmap: QPixmap):
        self._source_pixmap = QPixmap(pixmap)
        self.setText("")
        self.update()

    def clear_thumbnail(self, text: str):
        self._source_pixmap = QPixmap()
        QLabel.setPixmap(self, QPixmap())
        self.setText(text)
        self.update()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update()

    def _refit_thumbnail(self):
        # Compatibility with existing deferred refresh calls. Fit happens in
        # paintEvent so the drawn pixmap cannot be larger than the real box.
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)  # Keep theme styling and empty-state text.
        if self._source_pixmap.isNull():
            return
        # Keep the full image inside the visible mat and centre it on both axes.
        # The source is never zoom-cropped or distorted; tall and wide previews
        # both use true aspect-fit/contain behaviour.
        viewport = self.contentsRect().adjusted(8, 7, -8, -7)
        if viewport.width() < 1 or viewport.height() < 1:
            return
        image = self._source_pixmap.scaled(
            viewport.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        x = viewport.x() + (viewport.width() - image.width()) // 2
        y = viewport.y() + (viewport.height() - image.height()) // 2
        painter = QPainter(self)
        painter.setClipRect(viewport)
        painter.drawPixmap(x, y, image)
        painter.end()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        settings_app = "CASCuratorPublicBeta" if IS_PUBLIC_BUILD else "CASCuratorDEV"
        self.settings = QSettings("CASCurator", settings_app)

        def _stored_bool(key: str, default: bool = False) -> bool:
            raw = self.settings.value(key, default)
            if isinstance(raw, bool):
                return raw
            return str(raw).strip().lower() not in {"0", "false", "no", "off", ""}

        self.advanced_features = _stored_bool("advanced_features", False) if IS_PUBLIC_BUILD else True
        self.experimental_features = _stored_bool("experimental_features", False) if IS_PUBLIC_BUILD else True
        global _ADVANCED_FEATURES_ENABLED, _EXPERIMENTAL_FEATURES_ENABLED
        _ADVANCED_FEATURES_ENABLED = bool(self.advanced_features)
        _EXPERIMENTAL_FEATURES_ENABLED = bool(self.experimental_features)

        self.mdb: Optional[MDBData] = None
        self.visible_columns = [field for field in DEFAULT_COLUMNS if _public_field_visible(field)]

        # DEV keeps the full theme collection + private palette tools. PUBLIC
        # exposes only the explicitly curated release set.
        if IS_PUBLIC_BUILD:
            # Keep PUBLIC strictly to the user-approved flat theme list and order.
            self.themes = {name: dict(THEMES[name]) for name in PUBLIC_THEME_ORDER if name in THEMES}
        else:
            self.themes = {name: dict(theme) for name, theme in THEMES.items()}
        self.custom_theme_dir = Path(__file__).resolve().parent / "custom_themes"
        self.custom_theme_files: Dict[str, Path] = {}
        self.custom_theme_info: Dict[str, dict] = {}
        if IS_DEV_BUILD:
            self._load_custom_themes()
        default_theme = "Carbon Blue" if IS_PUBLIC_BUILD else "Shades of Grey"
        self.theme_name = self.settings.value("theme", default_theme)
        if self.theme_name not in self.themes:
            self.theme_name = default_theme
        self.theme = self.themes[self.theme_name]
        self._load_thread: Optional[QThread] = None
        self._load_worker: Optional[LoadWorker] = None
        self.pending_changes: Dict[int, Dict[str, str]] = {}
        # Physical worksheet row numbers staged for deletion; the package file is never touched.
        self.pending_deletions: set[int] = set()
        self.profile_registry_saved: List[str] = []
        self.profile_registry_current: List[str] = []
        self.profile_registry_dirty = False
        self.status_registry_dirty = False
        self.file_location_registry_saved: List[tuple[str, str]] = []
        self.file_location_registry_current: List[tuple[str, str]] = []
        self.file_location_registry_dirty = False
        self.merge_profile_registry_saved: List[str] = []
        self.merge_profile_registry_current: List[str] = []
        self.merge_profile_registry_dirty = False
        self.current_source_row: Optional[int] = None
        self._loading_editor = False
        self.pinned_packages: List[Dict[str, str]] = []
        self.view_mode = str(self.settings.value("library_view", "table") or "table")
        if self.view_mode not in {"table", "gallery"}:
            self.view_mode = "table"
        self.app_page = "library"
        self.scan_source_path = str(self.settings.value("scan_source_folder", "") or "")
        self.scan_mdb_path = ""
        self.routing_preview_root = str(self.settings.value("routing_preview_root", "") or "")
        self.cc_magic_dropin_path = str(self.settings.value("cc_magic_dropin_folder", "") or "")
        self._cc_magic_import_last_signature = None
        # Extra column filters are session UI state only; they never alter MDB
        # data and are rebuilt as removable chips below the permanent filter bar.
        self.extra_column_filters: List[Dict[str, str]] = []
        self.scan_process: Optional[QProcess] = None
        self.scan_run_kind = "scan"
        self.scan_total_packages = 0
        self.scan_current_package = 0
        self._scan_output_buffer = ""
        self.last_scan_output_dir: Optional[Path] = None
        self.last_scan_log_path: Optional[Path] = None
        self.last_routing_preview_path: Optional[Path] = None
        self.last_sorted_root: Optional[Path] = None
        self.last_scan_backup_path: Optional[Path] = None
        self.last_move_manifest_path: Optional[Path] = None
        self.last_recovery_manifest_path: Optional[Path] = None
        self.last_routing_location_tracking: Dict[str, object] = {}
        self.scan_stats: Dict[str, object] = {}
        self.last_scan_summary_text = ""
        self._post_scan_refresh_target: Optional[Path] = None
        self._scan_log_dialog = None

        # DEV-only Random Sim Studio state. Exclusions live in private app
        # settings so the experimental generator never changes scanner-owned
        # MDB fields or scanner logic.
        self.random_sim_grouped = {}
        self.random_sim_selection = {}
        self.random_sim_result_rows = []
        self.random_sim_count_boxes = {}
        self.random_sim_category_chance_boxes = {}
        try:
            self.random_sim_category_chances = json.loads(str(self.settings.value("random_sim/category_chances_json", "{}") or "{}"))
            if not isinstance(self.random_sim_category_chances, dict):
                self.random_sim_category_chances = {}
        except Exception:
            self.random_sim_category_chances = {}
        try:
            self.random_sim_type_rules = json.loads(str(self.settings.value("random_sim/type_rules_json", "{}") or "{}"))
            if not isinstance(self.random_sim_type_rules, dict):
                self.random_sim_type_rules = {}
        except Exception:
            self.random_sim_type_rules = {}
        try:
            self.random_sim_item_rules = json.loads(str(self.settings.value("random_sim/item_rules_json", "{}") or "{}"))
            if not isinstance(self.random_sim_item_rules, dict):
                self.random_sim_item_rules = {}
        except Exception:
            self.random_sim_item_rules = {}
        stored_exclusions = self.settings.value("random_sim/excluded_keys", []) or []
        if isinstance(stored_exclusions, str):
            stored_exclusions = [stored_exclusions] if stored_exclusions else []
        self.random_sim_excluded_keys = {str(value) for value in stored_exclusions if str(value)}
        self.random_sim_export_root = str(self.settings.value("random_sim/export_root", "") or "")
        self.random_sim_last_export: Optional[Path] = None
        # DEV-only persistent relationship metadata. This is deliberately stored
        # outside the scanner-owned Master Database so Random Sim usage can be
        # displayed without overwriting Status (for example, "Used in Game").
        self.random_sim_usage_registry_path = default_usage_registry_path()
        self.random_sim_usage_registry = load_usage_registry(self.random_sim_usage_registry_path)

        self.model = LibraryTableModel(self)
        self.model.set_pending_changes(self.pending_changes)
        self.model.set_pending_deletions(self.pending_deletions)
        self.model.manualFieldEdited.connect(self._table_manual_edit_changed)
        self.proxy = LibraryFilterProxy(self)
        self.proxy.setSourceModel(self.model)

        self.setWindowTitle(f"{APP_NAME} — {VERSION}")
        icon_path = Path(__file__).resolve().parent / "assets" / "cas_curator.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        # Theme native QMessageBox windows through one global event filter so
        # warnings/choices feel like part of CAS Curator instead of plain white
        # Windows popups. The event filter also keeps the existing details-pane
        # resize handling below.
        QApplication.instance().installEventFilter(self)
        screen = QApplication.primaryScreen().availableGeometry()
        w = min(1500, max(1100, int(screen.width() * 0.90)))
        h = min(900, max(680, int(screen.height() * 0.88)))
        self.resize(w, h)
        self.setMinimumSize(980, 640)

        self._build_menus()
        self._build_ui()
        self._set_app_page("library")
        self.apply_theme(self.theme_name)
        self.statusBar().showMessage("No Master Database loaded — choose File > Open Master Database…")
        # v1.0 behaviour: reopen the last successfully loaded MDB unless the
        # user explicitly disabled File > Open last used MDB on startup.
        raw_open_last = self.settings.value("open_last_mdb_on_startup", True)
        open_last = raw_open_last if isinstance(raw_open_last, bool) else str(raw_open_last).strip().lower() not in {"0", "false", "no", "off", ""}
        if open_last:
            QTimer.singleShot(0, self._restore_last_mdb)

    def _autosize_plain_text_edit(self, edit: QPlainTextEdit, min_rows: int, max_rows: int):
        """Keep short multiline fields compact and grow only as content needs."""
        if edit is None:
            return
        visual_lines = 0
        block = edit.document().firstBlock()
        while block.isValid():
            layout = block.layout()
            visual_lines += max(1, layout.lineCount() if layout is not None else 1)
            block = block.next()
        visual_lines = max(1, visual_lines)
        rows = max(min_rows, min(max_rows, visual_lines))
        line_height = max(18, edit.fontMetrics().lineSpacing())
        chrome = max(14, edit.frameWidth() * 2 + 12)
        target = rows * line_height + chrome
        minimum = min_rows * line_height + chrome
        maximum = max_rows * line_height + chrome
        edit.setFixedHeight(max(minimum, min(maximum, target)))
        edit.setVerticalScrollBarPolicy(
            Qt.ScrollBarAsNeeded if visual_lines > max_rows else Qt.ScrollBarAlwaysOff
        )

    def _update_dynamic_editor_heights(self, *_):
        if hasattr(self, "cc_magic_profiles_edit"):
            self._autosize_plain_text_edit(self.cc_magic_profiles_edit, 1, 5)
        if hasattr(self, "merge_profiles_edit"):
            self._autosize_plain_text_edit(self.merge_profiles_edit, 1, 5)
        if hasattr(self, "notes_edit"):
            self._autosize_plain_text_edit(self.notes_edit, 1, 6)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Wrapped line counts can change with the details-pane width. Recalculate
        # after Qt finishes the resize layout so Notes/Merge Profiles stay compact.
        if hasattr(self, "cc_magic_profiles_edit") or hasattr(self, "notes_edit") or hasattr(self, "merge_profiles_edit"):
            QTimer.singleShot(0, self._update_dynamic_editor_heights)

    def _build_menus(self):
        file_menu = self.menuBar().addMenu("File")
        open_action = QAction("Open Master Database…", self)
        open_action.triggered.connect(self.choose_file)
        file_menu.addAction(open_action)
        reload_action = QAction("Reload", self)
        reload_action.triggered.connect(self.reload)
        file_menu.addAction(reload_action)
        self.open_last_mdb_action = QAction("Open last used MDB on startup", self, checkable=True)
        raw_open_last = self.settings.value("open_last_mdb_on_startup", True)
        open_last = raw_open_last if isinstance(raw_open_last, bool) else str(raw_open_last).strip().lower() not in {"0", "false", "no", "off", ""}
        self.open_last_mdb_action.setChecked(open_last)
        self.open_last_mdb_action.toggled.connect(self._set_open_last_mdb_on_startup)
        file_menu.addAction(self.open_last_mdb_action)
        self.legacy_status_action = QAction("Review legacy Status values…", self)
        self.legacy_status_action.setToolTip(
            "Preview old Status labels and stage the agreed conversions without saving automatically."
        )
        self.legacy_status_action.triggered.connect(self.review_legacy_status_values)
        file_menu.addAction(self.legacy_status_action)
        file_menu.addSeparator()
        self.save_action = QAction("Save Changes", self)
        self.save_action.setShortcut("Ctrl+S")
        self.save_action.setEnabled(False)
        self.save_action.triggered.connect(self.save_changes)
        file_menu.addAction(self.save_action)
        self.revert_all_action = QAction("Revert All Unsaved Changes", self)
        self.revert_all_action.setEnabled(False)
        self.revert_all_action.triggered.connect(self.revert_all_changes)
        file_menu.addAction(self.revert_all_action)
        file_menu.addSeparator()
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        self.columns_menu = self.menuBar().addMenu("Columns")

        tools_menu = self.menuBar().addMenu("Tools")
        zip_tools_action = QAction("Preview / Extract ZIP…", self)
        zip_tools_action.setToolTip("Preview a ZIP archive safely, then extract it only after explicit confirmation.")
        zip_tools_action.triggered.connect(self._choose_zip_preview_extract)
        tools_menu.addAction(zip_tools_action)
        tools_menu.addSeparator()
        set_creator_selected_action = QAction("Set Creator for Selected…", self)
        set_creator_selected_action.setToolTip("Stage one Creator value for all selected Library packages. Changes remain unsaved until Save Changes.")
        set_creator_selected_action.triggered.connect(self.set_creator_for_selected)
        tools_menu.addAction(set_creator_selected_action)
        move_selected_action = QAction("Move Selected to Location…", self)
        move_selected_action.setToolTip("Preview and move the selected original package files to a saved File Location, then update File Location in the MDB after hash verification.")
        move_selected_action.triggered.connect(self.move_selected_to_location)
        tools_menu.addAction(move_selected_action)
        copy_dropin_action = QAction("Copy Selected to CC Magic Import…", self)
        copy_dropin_action.setToolTip("Copy selected original package files to the configured CC Magic Import folder. Originals remain in place and File Location is not changed.")
        copy_dropin_action.triggered.connect(self.copy_selected_to_cc_magic_dropin)
        tools_menu.addAction(copy_dropin_action)
        tools_menu.addSeparator()
        set_dropin_action = QAction("Set CC Magic Import Folder…", self)
        set_dropin_action.triggered.connect(self.set_cc_magic_dropin_folder)
        tools_menu.addAction(set_dropin_action)
        open_dropin_action = QAction("Open CC Magic Import Folder", self)
        open_dropin_action.triggered.connect(self.open_cc_magic_dropin_folder)
        tools_menu.addAction(open_dropin_action)

        options_menu = self.menuBar().addMenu("Options")

        def setting_bool(key: str, default: bool = True) -> bool:
            raw = self.settings.value(key, default)
            if isinstance(raw, bool):
                return raw
            return str(raw).strip().lower() not in {"0", "false", "no", "off", ""}

        # Save confirmation was intentionally removed for PUBLIC: Save Changes
        # always performs the existing automatic MDB backup without an extra prompt.
        self.settings.setValue("confirm_before_save", False)

        self.confirm_fill_action = QAction("Confirm before fill-down", self, checkable=True)
        self.confirm_fill_action.setChecked(setting_bool("confirm_before_fill_down", True))
        self.confirm_fill_action.setToolTip("Turn this off to apply drag-fill immediately to the displayed rows crossed by the drag.")
        self.confirm_fill_action.toggled.connect(
            lambda checked: (self.settings.setValue("confirm_before_fill_down", checked), self.settings.sync())
        )
        options_menu.addAction(self.confirm_fill_action)

        if IS_PUBLIC_BUILD:
            options_menu.addSeparator()
            self.advanced_features_action = QAction("Enable Advanced features", self, checkable=True)
            self.advanced_features_action.setChecked(bool(self.advanced_features))
            self.advanced_features_action.setToolTip(
                "Show extra technical information such as CASP IDs, Biggest Texture, HQ/Poly warnings, Risk, Health Issues, Last Scanned, Package Hash and All fields."
            )
            self.advanced_features_action.toggled.connect(self._set_advanced_features)
            options_menu.addAction(self.advanced_features_action)

            self.experimental_features_action = QAction("Enable Experimental features", self, checkable=True)
            self.experimental_features_action.setChecked(bool(self.experimental_features))
            self.experimental_features_action.setToolTip(
                "Show provisional scanner-derived classifications such as Subtype, Skin Type, Outfit Category and EA Resource Status. Experimental results never control Routing automatically."
            )
            self.experimental_features_action.toggled.connect(self._set_experimental_features)
            options_menu.addAction(self.experimental_features_action)

        self.theme_menu = self.menuBar().addMenu("Theme")
        self._populate_theme_menu()

    def _set_advanced_features(self, checked: bool):
        global _ADVANCED_FEATURES_ENABLED
        self.advanced_features = bool(checked)
        _ADVANCED_FEATURES_ENABLED = bool(checked)
        self.settings.setValue("advanced_features", bool(checked))
        self.settings.sync()
        self._refresh_feature_visibility()
        self.statusBar().showMessage(
            "Advanced features enabled — extra technical fields are available in Columns and details."
            if checked else
            "Advanced features disabled — extra technical fields are hidden, not deleted.",
            9000,
        )

    def _set_experimental_features(self, checked: bool):
        global _EXPERIMENTAL_FEATURES_ENABLED
        self.experimental_features = bool(checked)
        _EXPERIMENTAL_FEATURES_ENABLED = bool(checked)
        self.settings.setValue("experimental_features", bool(checked))
        if checked:
            raw_seen = self.settings.value("experimental_warning_seen", False)
            seen = raw_seen if isinstance(raw_seen, bool) else str(raw_seen).strip().lower() not in {"0", "false", "no", "off", ""}
            if not seen:
                QMessageBox.information(
                    self,
                    "Experimental features",
                    "Experimental features show scanner-derived information that is still being validated. "
                    "Results may be incomplete or incorrect. They are for review and testing only and do not change Routing or move package files automatically.",
                )
                self.settings.setValue("experimental_warning_seen", True)
        self.settings.sync()
        self._refresh_feature_visibility()
        self.statusBar().showMessage(
            "Experimental features enabled — provisional fields are available for review."
            if checked else
            "Experimental features disabled — provisional fields are hidden, not deleted.",
            9000,
        )

    def _refresh_feature_visibility(self):
        """Apply Advanced/Experimental display layers without touching workbook data."""
        if hasattr(self, "all_fields_tab_index"):
            self.tabs.setTabVisible(self.all_fields_tab_index, bool(IS_DEV_BUILD or self.advanced_features))

        if not self.mdb:
            return

        allowed = {"#", "Thumbnail"}
        allowed.update(field for field in self.mdb.headers if _public_field_visible(field))
        filtered = [field for field in self.visible_columns if field in allowed]
        if not filtered:
            filtered = [field for field in DEFAULT_COLUMNS if field in allowed]
        self.visible_columns = filtered
        self.model.set_columns(self.visible_columns)
        self._persist_visible_columns()
        self._build_columns_menu()
        self._resize_columns()
        self._update_column_fold_buttons()
        self._populate_filters()
        self._apply_filters()
        if self.current_source_row is not None:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)

    def _load_custom_themes(self):
        """Only load private files: never shadow a built-in, even by case."""
        self.custom_theme_dir.mkdir(parents=True, exist_ok=True)
        for name in list(getattr(self, "custom_theme_files", {})):
            self.themes.pop(name, None)
        self.custom_theme_files = {}
        self.custom_theme_info = {}
        protected = {name.casefold() for name in THEMES}
        used = set(protected)

        for path in sorted(self.custom_theme_dir.glob("*.s3theme")):
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
                name = str(payload.get("name", "")).strip()
                theme = payload.get("theme")
                required = {"accent", "accent_soft", "selection", "panel_alt", "thumb_border", "text", "muted"}
                if (not name or name.casefold() in used or not isinstance(theme, dict)
                        or not required.issubset(theme)):
                    continue
                theme = dict(theme)
                theme["_custom"] = True
                self.themes[name] = theme
                self.custom_theme_files[name] = path
                self.custom_theme_info[name] = payload
                used.add(name.casefold())
            except (OSError, ValueError, TypeError):
                # Invalid user files must not prevent the rest of the app launching.
                continue

    def _populate_theme_menu(self):
        self.theme_menu.clear()
        group = QActionGroup(self.theme_menu)
        group.setExclusive(True)

        if IS_PUBLIC_BUILD:
            # PUBLIC uses one flat curated list: no category submenus and no DEV themes.
            for name in PUBLIC_THEME_ORDER:
                if name not in self.themes:
                    continue
                act = QAction(name, self, checkable=True)
                act.setChecked(name == self.theme_name)
                act.triggered.connect(lambda checked=False, n=name: self.apply_theme(n))
                group.addAction(act)
                self.theme_menu.addAction(act)
        else:
            for group_name, names in THEME_GROUPS.items():
                available = [name for name in names if name in self.themes]
                if not available:
                    continue
                submenu = self.theme_menu.addMenu(group_name)
                for name in available:
                    act = QAction(name, self, checkable=True)
                    act.setChecked(name == self.theme_name)
                    act.triggered.connect(lambda checked=False, n=name: self.apply_theme(n))
                    group.addAction(act)
                    submenu.addAction(act)

        if IS_DEV_BUILD:
            custom_names = sorted(self.custom_theme_files, key=str.casefold)
            if custom_names:
                custom_menu = self.theme_menu.addMenu("Custom (Private)")
                for name in custom_names:
                    act = QAction(name, self, checkable=True)
                    act.setChecked(name == self.theme_name)
                    act.triggered.connect(lambda checked=False, n=name: self.apply_theme(n))
                    group.addAction(act)
                    custom_menu.addAction(act)

            self.theme_menu.addSeparator()
            dev_menu = self.theme_menu.addMenu("Private DEV Palette Tools")

            import_file = QAction("Import Palette File…", self)
            import_file.triggered.connect(self.import_palette_file)
            dev_menu.addAction(import_file)

            paste_palette = QAction("Paste HEX / Website URL…", self)
            paste_palette.triggered.connect(self.paste_palette)
            dev_menu.addAction(paste_palette)

            manager_action = QAction("Manage Custom Themes…", self)
            manager_action.triggered.connect(self.manage_custom_themes)
            dev_menu.addAction(manager_action)

            studio_action = QAction("Visual Theme Studio…", self)
            studio_action.triggered.connect(self.visual_theme_studio)
            dev_menu.addAction(studio_action)

            dev_menu.addSeparator()
            open_folder = QAction("Open Custom Themes Folder", self)
            open_folder.triggered.connect(self.open_custom_theme_folder)
            dev_menu.addAction(open_folder)

            reload_themes = QAction("Reload Custom Themes", self)
            reload_themes.triggered.connect(self.reload_custom_themes)
            dev_menu.addAction(reload_themes)

    def _next_custom_theme_path(self, name: str, exclude: Optional[Path] = None) -> Path:
        # Display names are independent of safe disk filenames. Never replace a
        # different imported theme because sanitising produced a collision.
        safe = re.sub(r"[^A-Za-z0-9._ -]+", "_", name).strip(" ._")[:80] or "Imported_Theme"
        candidate = self.custom_theme_dir / (safe + ".s3theme")
        n = 2
        while candidate.exists() and candidate != exclude:
            candidate = self.custom_theme_dir / f"{safe}_{n}.s3theme"
            n += 1
        return candidate

    @staticmethod
    def _save_theme_payload(path: Path, payload: dict):
        # Atomic save avoids corrupting the only saved copy if interrupted.
        temp = path.with_suffix(".s3theme.tmp")
        try:
            temp.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
            temp.replace(path)
        finally:
            if temp.exists():
                temp.unlink()

    def _install_palette(self, colours: List[str], meta: dict, theme_override: Optional[dict] = None):
        if len(colours) < 2 and theme_override is None:
            QMessageBox.warning(self, "Palette Import", "I could not find at least two valid colours. For Lospec, download its GPL/HEX file first.")
            return

        suggested = meta.get("suggested_name") or "Custom Colour Palette"
        original_name = str(meta.get("original_name", "") or "").strip()
        # Editable suggestion chooser: creative *options*, never a claim that
        # an unnamed Coolors URL has a real title. Keep a site's original
        # name in the first slot whenever one was actually available.
        name_dialog = QDialog(self)
        name_dialog.setWindowTitle("Install Private Theme — Choose a Name")
        name_dialog.setMinimumWidth(470)
        name_layout = QVBoxLayout(name_dialog)
        message = QLabel(("Original palette name found. Keep it, choose a creative alternative, "
                          "or type a different name:" if original_name else
                          "The source did not supply an original name. Choose a nature/food/"
                          "atmosphere-inspired suggestion, or type your own:"), name_dialog)
        message.setWordWrap(True)
        name_layout.addWidget(message)
        name_options = QComboBox(name_dialog)
        name_options.setEditable(True)
        proposed = [suggested] + list(meta.get("suggested_names") or [])
        seen = set()
        for proposal in proposed:
            if isinstance(proposal, str) and proposal.strip() and proposal.casefold() not in seen:
                seen.add(proposal.casefold())
                name_options.addItem(proposal.strip())
        if not name_options.count():
            name_options.addItem("Secret Garden")
        name_options.setCurrentText(suggested)
        name_options.setToolTip("Select a suggestion or enter any custom theme name.")
        name_layout.addWidget(name_options)
        palette_scroll = QScrollArea(name_dialog)
        palette_scroll.setWidgetResizable(True)
        palette_scroll.setMaximumHeight(139)
        palette_scroll.setMinimumHeight(51)
        palette_scroll.setToolTip("All imported palette colours; scroll if necessary.")
        palette_widget = QWidget(name_dialog)
        palette_grid = QGridLayout(palette_widget)
        palette_grid.setContentsMargins(4, 4, 4, 4)
        for index, colour in enumerate(colours[:48]):
            hex_value = _normalize_hex(colour)
            if hex_value:
                swatch = QLabel(name_dialog)
                swatch.setFixedSize(35, 29)
                swatch.setStyleSheet(f"background-color: {hex_value}; border: 1px solid #707070;")
                swatch.setToolTip(hex_value)
                palette_grid.addWidget(swatch, index // 10, index % 10)
        palette_scroll.setWidget(palette_widget)
        name_layout.addWidget(palette_scroll)
        name_layout.addWidget(QLabel(f"{len(colours[:48])} unique imported colours", name_dialog))
        note = QLabel("Suggestions are inspired by the palette, not official names from the website.", name_dialog)
        note.setWordWrap(True)
        name_layout.addWidget(note)
        name_buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
                                        parent=name_dialog)
        name_buttons.accepted.connect(name_dialog.accept)
        name_buttons.rejected.connect(name_dialog.reject)
        name_layout.addWidget(name_buttons)
        if name_dialog.exec() != QDialog.Accepted:
            return
        name = name_options.currentText().strip()
        if not name:
            QMessageBox.warning(self, "Palette Name", "Enter a theme name before installing.")
            return
        if any(name.casefold() == built.casefold() for built in THEMES):
            QMessageBox.warning(self, "Protected Theme", "That name belongs to a built-in theme. Choose another name.")
            return
        existing = next((k for k in self.custom_theme_files if k.casefold() == name.casefold()), None)
        if existing:
            answer = QMessageBox.question(self, "Replace Imported Theme?",
                f"An imported theme called '{existing}' already exists. Replace it?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if answer != QMessageBox.Yes:
                return

        try:
            if theme_override is None:
                theme = build_theme_from_palette(name, colours)
            else:
                required = {"accent", "accent_soft", "selection", "panel_alt", "thumb_border", "text", "muted"}
                if not isinstance(theme_override, dict) or not required.issubset(theme_override):
                    raise ValueError("This .s3theme file has incomplete theme settings.")
                theme = dict(theme_override)  # do not recreate/discard downloaded theme settings
            theme["_custom"] = True
            self.custom_theme_dir.mkdir(parents=True, exist_ok=True)
            old_path = self.custom_theme_files.get(existing) if existing else None
            path = old_path or self._next_custom_theme_path(name)
            # Reimporting the same name is an explicit replacement. Its original
            # metadata can still be carried forward when the new source omits it.
            old_info = self.custom_theme_info.get(existing, {}) if existing else {}
            # Never attach the old palette's original name/source to an entirely
            # different palette just because the new import shares a display name.
            same_source = bool(
                (meta.get("source_url") and meta["source_url"] == old_info.get("source_url"))
                or (meta.get("source_filename") and meta["source_filename"] == old_info.get("source_filename"))
            )
            saved_original = original_name or (old_info.get("original_name", "") if same_source else "")
            payload = {
                "schema": 2,
                "name": name,
                "original_name": saved_original,
                "name_source": ("original" if saved_original and name.casefold() == str(saved_original).casefold()
                                else "suggested" if name.casefold() in {
                                    str(v).casefold() for v in meta.get("suggested_names", [])
                                } else "custom"),
                "source_site": meta.get("source_site", "") or (old_info.get("source_site", "") if same_source else ""),
                "source_url": meta.get("source_url", "") or (old_info.get("source_url", "") if same_source else ""),
                "source_filename": meta.get("source_filename", "") or (old_info.get("source_filename", "") if same_source else ""),
                "private_dev_only": True,
                "palette": colours,
                "theme": theme,
            }
            self._save_theme_payload(path, payload)
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "Palette Import", str(exc))
            return

        self._load_custom_themes()
        self.apply_theme(name)
        self._populate_theme_menu()
        label_detail = (f"Original palette name: {payload['original_name']}"
                        if payload['original_name'] else "No original palette name supplied; suggested a name.")
        QMessageBox.information(self, "Theme Installed",
            f"Installed '{name}' for this private DEV copy.\n{label_detail}\n\nSaved to:\n{path}")

    def import_palette_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Palette", str(Path.home()),
            "Palette files (*.hex *.txt *.gpl *.pal *.css *.json *.s3theme);;All files (*)")
        if not path:
            return
        try:
            raw = Path(path).read_text(encoding="utf-8-sig", errors="replace")
            existing_data = json.loads(raw) if Path(path).suffix.lower() in {".json", ".s3theme"} else {}
            if not isinstance(existing_data, dict):
                existing_data = {}
            palette = existing_data.get("palette", [])
            colours = [c for c in palette[:48] if isinstance(c, str)] if isinstance(palette, list) else []
            if len(colours) < 2:
                colours = parse_palette_text(raw)
            meta = palette_metadata(raw, colours, filename=path, online_lookup=False)
            # A native theme preserves all stored settings and original provenance.
            theme_override = existing_data.get("theme") if path.lower().endswith(".s3theme") else None
            if theme_override is not None:
                # In a saved v2 theme, 'name' can be a user-edited display name.
                # An absent original_name is intentional and must stay absent.
                if "original_name" in existing_data:
                    meta["original_name"] = existing_data.get("original_name", "")
                if existing_data.get("name"):
                    meta["suggested_name"] = str(existing_data["name"])
                for key in ("name_source", "source_site", "source_url", "source_filename"):
                    if existing_data.get(key):
                        meta[key] = existing_data[key]
        except (OSError, ValueError) as exc:
            QMessageBox.critical(self, "Palette Import", f"Could not read that palette file.\n\n{exc}")
            return
        self._install_palette(colours, meta, theme_override=theme_override)

    def paste_palette(self):
        # A real QPlainTextEdit guarantees the normal right-click > Paste menu.
        dialog = QDialog(self)
        dialog.setWindowTitle("Paste Palette or Website URL")
        dialog.resize(610, 350)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(
            "Paste HEX colours, GIMP palette text, Coolors palette URL or supported website data.\n"
            "Right-click in this box to Paste, or use Ctrl+V.\n"
            "If a website has no available palette name, one will be suggested."))
        editor = QPlainTextEdit(dialog)
        editor.setPlaceholderText("https://coolors.co/00ced1-20b2aa-4682b4-b0e0e6")
        layout.addWidget(editor)
        row = QHBoxLayout()
        paste_button = QPushButton("Paste from Clipboard", dialog)
        paste_button.clicked.connect(lambda: editor.insertPlainText(QApplication.clipboard().text()))
        row.addWidget(paste_button)
        row.addStretch(1)
        layout.addLayout(row)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=dialog)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        raw = editor.toPlainText().strip()
        if not raw:
            return
        colours = parse_palette_text(raw)
        meta = palette_metadata(raw, colours, online_lookup=True)
        self._install_palette(colours, meta)

    def _rename_custom_theme(self, name: str) -> Optional[str]:
        if name not in self.custom_theme_files:
            QMessageBox.warning(self, "Protected Theme", "Only imported themes can be renamed.")
            return None
        new_name, ok = QInputDialog.getText(self, "Rename Imported Theme", "New display name:",
                                             QLineEdit.Normal, name)
        if not ok or not new_name.strip():
            return None
        new_name = new_name.strip()
        if any(new_name.casefold() == n.casefold() for n in THEMES):
            QMessageBox.warning(self, "Protected Theme", "A built-in theme already has this name.")
            return None
        if any(new_name.casefold() == n.casefold() and n != name for n in self.custom_theme_files):
            QMessageBox.warning(self, "Theme Name", "Another imported theme already has this name.")
            return None
        if new_name == name:
            return name
        old_path = self.custom_theme_files[name]
        path = old_path if new_name.casefold() == name.casefold() else self._next_custom_theme_path(new_name, old_path)
        payload = dict(self.custom_theme_info[name])
        payload["name"] = new_name
        # Rename changes only the display name; keep original palette provenance
        # and every theme colour/setting exactly as imported.
        original = str(payload.get("original_name", "") or "")
        payload["name_source"] = "original" if original and new_name.casefold() == original.casefold() else "custom"
        try:
            self._save_theme_payload(path, payload)
            if path != old_path:
                old_path.unlink()
        except OSError as exc:
            if path != old_path and path.exists():
                path.unlink(missing_ok=True)
            QMessageBox.warning(self, "Rename Failed", f"Theme files were not renamed.\n\n{exc}")
            return None
        was_active = self.theme_name == name
        self._load_custom_themes()
        if was_active:
            self.apply_theme(new_name)  # also updates QSettings for next launch
        self._populate_theme_menu()
        return new_name

    def _delete_custom_theme(self, name: str) -> bool:
        if name not in self.custom_theme_files:
            QMessageBox.warning(self, "Protected Theme", "Built-in themes cannot be deleted.")
            return False
        answer = QMessageBox.question(self, "Delete Imported Theme?",
            f"Permanently delete the imported theme '{name}'?\n\n"
            "Your built-in themes and all CC records will remain unchanged.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            return False
        try:
            self.custom_theme_files[name].unlink()
        except OSError as exc:
            QMessageBox.warning(self, "Delete Failed", str(exc))
            return False
        was_active = self.theme_name == name
        self._load_custom_themes()
        if was_active:
            self.apply_theme("Shades of Grey")
        self._populate_theme_menu()
        return True

    def manage_custom_themes(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Custom Theme Manager — Private DEV")
        dialog.resize(780, 480)
        layout = QVBoxLayout(dialog)
        instruction = QLabel("Only imported themes can be renamed or deleted. Built-in themes are protected.")
        instruction.setWordWrap(True)
        layout.addWidget(instruction)
        table = QTableWidget(0, 3, dialog)
        table.setHorizontalHeaderLabels(["Imported theme", "Original palette name", "Source"])
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        layout.addWidget(table)
        description = QLabel("Select a custom theme to view its source and palette.")
        description.setWordWrap(True)
        layout.addWidget(description)
        palette_scroll = QScrollArea(dialog)
        palette_scroll.setWidgetResizable(True)
        palette_scroll.setFixedHeight(109)
        palette_widget = QWidget(dialog)
        palette_grid = QGridLayout(palette_widget)
        palette_grid.setContentsMargins(4, 4, 4, 4)
        palette_grid.setSpacing(4)
        palette_scroll.setWidget(palette_widget)
        layout.addWidget(palette_scroll)
        buttons_row = QHBoxLayout()
        apply_btn = QPushButton("Apply", dialog)
        rename_btn = QPushButton("Rename…", dialog)
        delete_btn = QPushButton("Delete…", dialog)
        studio_btn = QPushButton("Visual Theme Studio…", dialog)
        close_btn = QPushButton("Close", dialog)
        for btn in (apply_btn, rename_btn, delete_btn, studio_btn):
            buttons_row.addWidget(btn)
        buttons_row.addStretch(1)
        buttons_row.addWidget(close_btn)
        layout.addLayout(buttons_row)
        close_btn.clicked.connect(dialog.accept)

        def selected_name():
            row = table.currentRow()
            item = table.item(row, 0) if row >= 0 else None
            return item.text() if item else ""

        def refresh(preferred: str = ""):
            table.setRowCount(0)
            for name in sorted(self.custom_theme_files, key=str.casefold):
                info = self.custom_theme_info.get(name, {})
                row = table.rowCount()
                table.insertRow(row)
                for col, value in enumerate((name, info.get("original_name") or "(not recorded)",
                                             info.get("source_site") or "Unknown")):
                    table.setItem(row, col, QTableWidgetItem(str(value)))
                if name == preferred:
                    table.selectRow(row)
            if table.rowCount() and table.currentRow() < 0:
                table.selectRow(0)
            update_selection()

        def update_selection():
            name = selected_name()
            for btn in (apply_btn, rename_btn, delete_btn, studio_btn):
                btn.setEnabled(bool(name))
            while palette_grid.count():
                item = palette_grid.takeAt(0)
                widget = item.widget()
                if widget:
                    widget.deleteLater()
            if not name:
                description.setText("No imported themes yet. Use Theme > Private DEV Palette Tools to import one.")
                return
            info = self.custom_theme_info.get(name, {})
            original = info.get("original_name") or "Not recorded (older theme or suggested name)"
            source = info.get("source_url") or info.get("source_filename") or info.get("source_site") or "Not recorded"
            description.setText(f"Original: {original}\nSource: {source}")
            palette = info.get("palette") or self.themes[name].get("_palette", [])
            for index, hex_value in enumerate(palette[:48]):
                if not isinstance(hex_value, str) or not _normalize_hex(hex_value):
                    continue
                swatch = QLabel(dialog)
                swatch.setFixedSize(28, 25)
                swatch.setToolTip(hex_value)
                swatch.setStyleSheet(f"background-color: {_normalize_hex(hex_value)}; border: 1px solid #777;")
                palette_grid.addWidget(swatch, index // 12, index % 12)

        def apply_selected():
            name = selected_name()
            if name:
                self.apply_theme(name)
                self._populate_theme_menu()

        def rename_selected():
            name = selected_name()
            if name:
                new_name = self._rename_custom_theme(name)
                if new_name:
                    refresh(new_name)

        def delete_selected():
            name = selected_name()
            if name and self._delete_custom_theme(name):
                refresh()

        def studio_selected():
            name = selected_name()
            if name:
                self.visual_theme_studio(name)
                refresh(self.theme_name if self.theme_name in self.custom_theme_files else name)

        table.itemSelectionChanged.connect(update_selection)
        apply_btn.clicked.connect(apply_selected)
        rename_btn.clicked.connect(rename_selected)
        delete_btn.clicked.connect(delete_selected)
        studio_btn.clicked.connect(studio_selected)
        refresh(self.theme_name)
        dialog.exec()

    def visual_theme_studio(self, selected_name: str = ""):
        """Private DEV-only isolated preview editor; never edits a built-in."""
        if not IS_DEV_BUILD:
            return
        if not self.custom_theme_files:
            QMessageBox.information(self, "Visual Theme Studio",
                "Import a custom palette first: Theme > Private DEV Palette Tools > Paste HEX / Website URL.")
            return
        from theme_studio import ThemeStudio
        initial = selected_name if selected_name in self.custom_theme_files else self.theme_name
        studio = ThemeStudio(self, initial, self)
        studio.exec()

    def open_custom_theme_folder(self):
        self.custom_theme_dir.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.custom_theme_dir)))

    def reload_custom_themes(self):
        self._load_custom_themes()
        if self.theme_name not in self.themes:
            self.theme_name = "Shades of Grey"
        self._populate_theme_menu()
        self.apply_theme(self.theme_name)
        self.statusBar().showMessage(f"Reloaded {len(self.custom_theme_files)} private custom theme(s).", 5000)

    def _build_ui(self):
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(10, 8, 10, 10)
        root_layout.setSpacing(6)
        self.setCentralWidget(root)

        # Compact shared Library / Scan navigation lives inside the themed top
        # panel instead of consuming a separate white toolbar row. The Scan page
        # remains a non-executing shell in v0.4.1.
        self.filter_card = QFrame()
        self.filter_card.setObjectName("FilterCard")
        card_layout = QVBoxLayout(self.filter_card)
        card_layout.setContentsMargins(10, 5, 10, 6)
        card_layout.setSpacing(6)

        stats = QHBoxLayout()
        stats.setSpacing(6)

        self.library_page_btn = QPushButton("Library")
        self.scan_page_btn = QPushButton("Scan")
        self.guide_page_btn = QPushButton("Guide")
        self.random_sim_page_btn = QPushButton("Random Sim") if IS_DEV_BUILD else None
        nav_buttons = [self.library_page_btn, self.scan_page_btn, self.guide_page_btn]
        if self.random_sim_page_btn is not None:
            nav_buttons.append(self.random_sim_page_btn)
        for button in nav_buttons:
            button.setObjectName("AppPageButton")
            button.setCheckable(True)
            button.setMinimumWidth(66)
            button.setMaximumHeight(32)
        self.library_page_btn.clicked.connect(lambda checked=False: self._set_app_page("library"))
        self.scan_page_btn.clicked.connect(lambda checked=False: self._set_app_page("scan"))
        self.guide_page_btn.clicked.connect(lambda checked=False: self._set_app_page("guide"))
        if self.random_sim_page_btn is not None:
            self.random_sim_page_btn.clicked.connect(lambda checked=False: self._set_app_page("random_sim"))
        self.library_page_btn.setChecked(True)
        stats.addWidget(self.library_page_btn)
        stats.addWidget(self.scan_page_btn)
        stats.addWidget(self.guide_page_btn)
        if self.random_sim_page_btn is not None:
            stats.addWidget(self.random_sim_page_btn)
        stats.addSpacing(8)

        self.packages_label = QLabel("Packages: —")
        self.packages_label.setObjectName("StatsStrong")
        self.shown_label = QLabel("Shown: —")
        self.thumbs_label = QLabel("Thumbnails: —")
        self.unsaved_label = QLabel("Unsaved: 0")
        self.unsaved_label.setObjectName("UnsavedLabel")
        stats.addWidget(self.packages_label)
        stats.addSpacing(10)
        stats.addWidget(self.shown_label)
        stats.addSpacing(10)
        stats.addWidget(self.thumbs_label)
        stats.addSpacing(10)
        stats.addWidget(self.unsaved_label)
        stats.addStretch(1)

        self.view_label = QLabel("View:")
        self.view_label.setObjectName("ViewLabel")
        self.table_view_btn = QPushButton("Table")
        self.gallery_view_btn = QPushButton("Gallery")
        # RC58: retain the compact RC42 control sizing, but group actions by
        # function instead of forcing unlike controls into vertical columns.
        # View toggles therefore sit together as one compact pair.
        for button in (self.table_view_btn, self.gallery_view_btn):
            button.setObjectName("ViewToggle")
            button.setCheckable(True)
            button.setMinimumWidth(68)
            button.setMaximumHeight(28)

        self.table_view_btn.clicked.connect(lambda checked=False: self._set_library_view("table"))
        self.gallery_view_btn.clicked.connect(lambda checked=False: self._set_library_view("gallery"))
        view_actions = QWidget()
        view_actions.setObjectName("LibraryActionGroup")
        view_actions_layout = QHBoxLayout(view_actions)
        view_actions_layout.setContentsMargins(0, 0, 0, 0)
        view_actions_layout.setSpacing(6)
        view_actions_layout.addWidget(self.table_view_btn)
        view_actions_layout.addWidget(self.gallery_view_btn)
        stats.addWidget(self.view_label)
        stats.addWidget(view_actions)

        self.scan_shell_badge = QLabel("CAS Scanner" if IS_PUBLIC_BUILD else "Scanner: DEV")
        self.scan_shell_badge.setObjectName("ScanShellBadge")
        self.scan_shell_badge.setVisible(False)
        stats.addWidget(self.scan_shell_badge)

        self.library_header_widgets = [
            self.packages_label, self.shown_label, self.thumbs_label, self.unsaved_label,
            self.view_label, self.table_view_btn, self.gallery_view_btn,
        ]
        card_layout.addLayout(stats)

        self.library_mdb_row = QFrame()
        self.library_mdb_row.setObjectName("LibraryTopRow")
        mdb_row = QHBoxLayout(self.library_mdb_row)
        mdb_row.setContentsMargins(0, 0, 0, 0)
        mdb_row.setSpacing(8)
        mdb_caption = QLabel("Opened MDB:")
        mdb_caption.setObjectName("MDBPathCaption")
        self.mdb_path_label = QLabel("None — choose a Master Database")
        self.mdb_path_label.setObjectName("MDBPathLabel")
        self.mdb_path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.mdb_path_label.setToolTip("No Master Database loaded")
        create_mdb_btn = QPushButton("Create new MDB…")
        create_mdb_btn.setObjectName("OpenMDBButton")
        create_mdb_btn.clicked.connect(self.create_new_mdb)
        create_mdb_btn.setMinimumWidth(125)
        open_mdb_btn = QPushButton("Open MDB…")
        open_mdb_btn.setObjectName("OpenMDBButton")
        open_mdb_btn.clicked.connect(self.choose_file)
        open_mdb_btn.setMinimumWidth(100)
        mdb_actions = QWidget()
        mdb_actions.setObjectName("LibraryActionGroup")
        mdb_actions_layout = QHBoxLayout(mdb_actions)
        mdb_actions_layout.setContentsMargins(0, 0, 0, 0)
        mdb_actions_layout.setSpacing(6)
        mdb_actions_layout.addWidget(create_mdb_btn)
        mdb_actions_layout.addWidget(open_mdb_btn)
        mdb_row.addWidget(mdb_caption)
        mdb_row.addWidget(self.mdb_path_label, 1)
        mdb_row.addWidget(mdb_actions)
        card_layout.addWidget(self.library_mdb_row)

        self.library_filters_row = QFrame()
        self.library_filters_row.setObjectName("LibraryFiltersRow")
        filters = QGridLayout(self.library_filters_row)
        filters.setContentsMargins(0, 0, 0, 0)
        filters.setHorizontalSpacing(8)
        filters.setVerticalSpacing(5)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Filename, CASP ID, type, notes…")
        self.search_edit.setMinimumWidth(230)
        # Native trailing X clears only the search text; Reset filters clears search plus every filter.
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self._apply_filters)

        self.filter_boxes: Dict[str, QComboBox] = {}
        for field, _label in FILTER_FIELDS:
            box = ArrowComboBox()
            box.setMinimumWidth(160 if field == "CC Magic Profiles" else (190 if field == "File Location" else (135 if field == "CC Magic Import" else 120)))
            box.addItem("All")
            box.currentTextChanged.connect(self._apply_filters)
            self.filter_boxes[field] = box

        # Compact two-row filter layout. Labels are right-aligned inside their
        # grid cells so short labels such as Status do not float far away from
        # their controls merely because a longer label shares the same column.
        def filter_label(text: str) -> QLabel:
            label = QLabel(text)
            label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            return label

        clear_btn = QPushButton("Reset filters")
        clear_btn.clicked.connect(self.clear_filters)
        clear_btn.setFixedWidth(108)
        clear_btn.setToolTip("Clear Search, permanent filters and all filters added with + Add Filter…")

        self.new_profile_btn = QPushButton("New Profile…")
        self.new_profile_btn.setToolTip("Create an empty general Profile before assigning any packages. Profiles can represent playthroughs, themes, saves, testing groups, or other organisational groups.")
        self.new_profile_btn.clicked.connect(self._create_profile)
        self.new_profile_btn.setMinimumWidth(100)
        self.new_profile_btn.setMaximumWidth(125)
        self.manage_profiles_btn = QPushButton("Manage Profiles…")
        self.manage_profiles_btn.setToolTip("Rename or remove general Profiles and review how many packages use each one. This does not restore the old bulk Profile Tools workflow.")
        self.manage_profiles_btn.clicked.connect(self._manage_cc_magic_profiles)
        self.manage_profiles_btn.setMinimumWidth(120)
        self.manage_profiles_btn.setMaximumWidth(150)

        if IS_PUBLIC_BUILD:
            # Keep only broad, everyday filters permanently visible. More
            # specific metadata (Health, File Location, CC Magic Import, CAS
            # classifications, Advanced/Experimental fields, etc.) lives under
            # + Add Filter… instead of consuming permanent toolbar space.
            # RC61 keeps the preferred RC58 grouping, but shortens the left
            # Search/CC Type field area enough to give the three permanent
            # filter groups proper columns. Status now sits directly above
            # Profile, while Creator keeps its own middle column.
            self.add_filter_btn = QPushButton("+ Add Filter…")
            self.add_filter_btn.setToolTip("Filter by any data field currently available in this Library mode.")
            self.add_filter_btn.setFixedWidth(112)
            self.add_filter_btn.clicked.connect(self._add_column_filter)

            filter_actions = QWidget()
            filter_actions.setObjectName("LibraryActionGroup")
            filter_actions_layout = QHBoxLayout(filter_actions)
            filter_actions_layout.setContentsMargins(0, 0, 0, 0)
            filter_actions_layout.setSpacing(6)
            filter_actions_layout.addWidget(self.add_filter_btn)
            filter_actions_layout.addWidget(clear_btn)

            profile_actions = QWidget()
            profile_actions.setObjectName("LibraryActionGroup")
            profile_actions_layout = QHBoxLayout(profile_actions)
            profile_actions_layout.setContentsMargins(0, 0, 0, 0)
            profile_actions_layout.setSpacing(6)
            profile_actions_layout.addWidget(self.new_profile_btn)
            profile_actions_layout.addWidget(self.manage_profiles_btn)

            # Keep the RC61 Status/Profile alignment, but let Search use the
            # otherwise-empty Creator columns on the first row. This makes the
            # Search field wider without moving the aligned Status/Profile pair.
            filters.addWidget(filter_label("Search:"), 0, 0)
            filters.addWidget(self.search_edit, 0, 1, 1, 3)
            filters.addWidget(filter_label("Status:"), 0, 4)
            filters.addWidget(self.filter_boxes["Status"], 0, 5)
            filters.addWidget(filter_actions, 0, 6)

            filters.addWidget(filter_label("CC Type:"), 1, 0)
            filters.addWidget(self.filter_boxes["CC Type"], 1, 1)
            filters.addWidget(filter_label("Creator:"), 1, 2)
            filters.addWidget(self.filter_boxes["Creator"], 1, 3)
            filters.addWidget(filter_label("Profile:"), 1, 4)
            filters.addWidget(self.filter_boxes["CC Magic Profiles"], 1, 5)
            filters.addWidget(profile_actions, 1, 6)

            filters.setColumnStretch(0, 0)
            filters.setColumnStretch(1, 4)
            filters.setColumnStretch(2, 0)
            filters.setColumnStretch(3, 1)
            filters.setColumnStretch(4, 0)
            filters.setColumnStretch(5, 1)
            filters.setColumnStretch(6, 0)
        else:
            # DEV retains its existing compact technical quick-filter row.
            filters.addWidget(filter_label("Search:"), 0, 0)
            filters.addWidget(self.search_edit, 0, 1, 1, 3)
            filters.addWidget(filter_label("Status:"), 0, 4)
            filters.addWidget(self.filter_boxes["Status"], 0, 5)
            filters.addWidget(filter_label("Health:"), 0, 6)
            filters.addWidget(self.filter_boxes["Health Status"], 0, 7)
            filters.addWidget(clear_btn, 0, 8)
            filters.addWidget(filter_label("CC Type:"), 1, 0)
            filters.addWidget(self.filter_boxes["CC Type"], 1, 1)
            filters.addWidget(filter_label("Risk:"), 1, 2)
            filters.addWidget(self.filter_boxes["Risk"], 1, 3)
            filters.addWidget(filter_label("File Location:"), 1, 4)
            filters.addWidget(self.filter_boxes["File Location"], 1, 5)
            filters.addWidget(filter_label("Profile:"), 1, 6)
            filters.addWidget(self.filter_boxes["CC Magic Profiles"], 1, 7)
            profile_actions = QWidget()
            profile_actions_layout = QHBoxLayout(profile_actions)
            profile_actions_layout.setContentsMargins(0, 0, 0, 0)
            profile_actions_layout.setSpacing(6)
            profile_actions_layout.addWidget(self.new_profile_btn)
            profile_actions_layout.addWidget(self.manage_profiles_btn)
            filters.addWidget(profile_actions, 1, 8)
            filters.setColumnStretch(1, 1)
            filters.setColumnStretch(3, 0)
            filters.setColumnStretch(5, 0)
            filters.setColumnStretch(7, 0)
        card_layout.addWidget(self.library_filters_row)

        # Universal filters are displayed as removable chips on a compact,
        # horizontally scrollable strip. It remains hidden when no extra
        # filters are active so the normal Library stays uncluttered.
        self.active_filter_bar = QFrame()
        self.active_filter_bar.setObjectName("ActiveFilterBar")
        active_outer = QHBoxLayout(self.active_filter_bar)
        active_outer.setContentsMargins(0, 1, 0, 0)
        active_outer.setSpacing(6)
        active_label = QLabel("Active filters:")
        active_label.setObjectName("ActiveFilterLabel")
        active_outer.addWidget(active_label)
        self.active_filter_scroll = QScrollArea()
        self.active_filter_scroll.setObjectName("ActiveFilterScroll")
        self.active_filter_scroll.setWidgetResizable(True)
        self.active_filter_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.active_filter_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.active_filter_scroll.setFrameShape(QFrame.NoFrame)
        self.active_filter_scroll.setFixedHeight(32)
        self.active_filter_chip_host = QWidget()
        self.active_filter_chip_host.setObjectName("ActiveFilterChipHost")
        self.active_filter_chip_layout = QHBoxLayout(self.active_filter_chip_host)
        self.active_filter_chip_layout.setContentsMargins(0, 0, 0, 0)
        self.active_filter_chip_layout.setSpacing(5)
        self.active_filter_chip_layout.addStretch(1)
        self.active_filter_scroll.setWidget(self.active_filter_chip_host)
        active_outer.addWidget(self.active_filter_scroll, 1)
        self.active_filter_bar.setVisible(False)
        card_layout.addWidget(self.active_filter_bar)
        root_layout.addWidget(self.filter_card)

        # Persistent package pins. Hidden until at least one package is pinned.
        self.pinned_bar = QFrame()
        self.pinned_bar.setObjectName("PinnedBar")
        pinned_outer = QHBoxLayout(self.pinned_bar)
        pinned_outer.setContentsMargins(4, 0, 4, 0)
        pinned_outer.setSpacing(6)
        pinned_label = QLabel("Pinned:")
        pinned_label.setObjectName("PinnedLabel")
        pinned_outer.addWidget(pinned_label)
        self.pinned_scroll = QScrollArea()
        self.pinned_scroll.setObjectName("PinnedScroll")
        self.pinned_scroll.setWidgetResizable(True)
        self.pinned_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.pinned_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.pinned_scroll.setFixedHeight(34)
        self.pinned_scroll_widget = QWidget()
        self.pinned_tabs_layout = QHBoxLayout(self.pinned_scroll_widget)
        self.pinned_tabs_layout.setContentsMargins(0, 0, 0, 0)
        self.pinned_tabs_layout.setSpacing(5)
        self.pinned_tabs_layout.addStretch(1)
        self.pinned_scroll.setWidget(self.pinned_scroll_widget)
        pinned_outer.addWidget(self.pinned_scroll, 1)
        self.pinned_bar.setFixedHeight(40)
        self.pinned_bar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.pinned_bar.setVisible(False)
        # The pinned strip belongs to the Library side only. Keep its height fixed so it
        # can never absorb spare vertical space above the package table.

        self.splitter = QSplitter(Qt.Horizontal)
        root_layout.addWidget(self.splitter, 1)

        self.scan_page = self._build_scan_shell_page()
        self.scan_page.setVisible(False)
        root_layout.addWidget(self.scan_page, 1)

        self.guide_page = self._build_guide_page()
        self.guide_page.setVisible(False)
        root_layout.addWidget(self.guide_page, 1)

        self.random_sim_page = self._build_random_sim_page() if IS_DEV_BUILD else None
        if self.random_sim_page is not None:
            self.random_sim_page.setVisible(False)
            root_layout.addWidget(self.random_sim_page, 1)

        # Main library table.
        left = QFrame()
        left.setObjectName("MainPanel")
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)
        self.table = OutlineTableView(self.theme)
        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(False)
        self.table.setShowGrid(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(27)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.table.horizontalHeader().setSectionsMovable(True)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.setWordWrap(False)
        self.table.setEditTriggers(QTableView.DoubleClicked | QTableView.EditKeyPressed)
        self.table.setToolTip("Double-click an editable cell to edit it. To copy a value down, select an editable cell and drag the small fill handle at its lower-right corner across the displayed rows. Fill-down respects sorting/filtering and remains unsaved until Save Changes. Blank source values are blocked to prevent accidental clearing. Ctrl+click selects individual packages; Shift+click selects a range. Right-click a row to open its assigned File Location, search the web for the package or its filled Creator, copy its file name or stage Delete Record. Deletions require Save Changes and create a backup; package files are never deleted. Unsaved rows get an amber marker at the left.")
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_table_context_menu)
        self.table.fillRequested.connect(self._table_fill_requested)
        self.table.clicked.connect(self._show_selected)
        self.table.selectionModel().selectionChanged.connect(lambda *_: self._show_selected())

        self.gallery = QListView()
        self.gallery.setObjectName("GalleryView")
        self.gallery.setModel(self.proxy)
        self.gallery.setModelColumn(0)
        self.gallery.setViewMode(QListView.IconMode)
        self.gallery.setResizeMode(QListView.Adjust)
        self.gallery.setMovement(QListView.Static)
        self.gallery.setWrapping(True)
        self.gallery.setSpacing(6)
        self.gallery.setGridSize(QSize(196, 228))
        self.gallery.setUniformItemSizes(True)
        self.gallery.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.gallery.setSelectionModel(self.table.selectionModel())
        self.gallery_delegate = GalleryDelegate(self.proxy, self.model, self.gallery)
        self.gallery.setItemDelegate(self.gallery_delegate)
        self.gallery.clicked.connect(self._gallery_clicked)
        self.gallery.doubleClicked.connect(self._gallery_clicked)

        # Compact table-only column folds. The existing Columns menu remains
        # the fine-grained chooser; these buttons are merely quick show/hide
        # shortcuts for related information.
        self.column_fold_bar = QFrame()
        self.column_fold_bar.setObjectName("ColumnFoldBar")
        column_fold_layout = QHBoxLayout(self.column_fold_bar)
        column_fold_layout.setContentsMargins(6, 3, 6, 3)
        column_fold_layout.setSpacing(5)
        column_fold_label = QLabel("Columns:")
        column_fold_label.setObjectName("ColumnFoldLabel")
        column_fold_layout.addWidget(column_fold_label)
        self.column_fold_buttons = {}
        for group_name in COLUMN_FOLD_GROUPS:
            button = QPushButton(f"{group_name} +")
            button.setObjectName("ColumnFoldButton")
            button.setMaximumHeight(25)
            button.setMinimumWidth(78 if group_name != "Organise" else 92)
            button.clicked.connect(lambda checked=False, g=group_name: self._toggle_column_fold_group(g))
            self.column_fold_buttons[group_name] = button
            column_fold_layout.addWidget(button)
        column_fold_layout.addStretch(1)
        self.column_fold_bar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        # FIX46 Gallery quick-filter visual layout. Keep the release single-row
        # control order, labels, margins, spacing and control widths exactly.
        self.gallery_quick_bar = QFrame()
        self.gallery_quick_bar.setObjectName("GalleryQuickBar")
        gallery_quick_layout = QHBoxLayout(self.gallery_quick_bar)
        gallery_quick_layout.setContentsMargins(7, 4, 7, 4)
        gallery_quick_layout.setSpacing(5)
        gallery_quick_label = QLabel("Gallery filters:")
        gallery_quick_label.setObjectName("GalleryQuickLabel")
        gallery_quick_layout.addWidget(gallery_quick_label)

        self.gallery_status_buttons = {}
        for status_name in PUBLIC_STATUS_OPTIONS:
            button = QPushButton(status_name)
            button.setObjectName("GalleryQuickButton")
            button.setCheckable(True)
            button.clicked.connect(lambda checked=False, s=status_name: self._gallery_status_filter_clicked(s, checked))
            self.gallery_status_buttons[status_name] = button
            gallery_quick_layout.addWidget(button)

        gallery_type_label = QLabel("CC Type:")
        gallery_type_label.setObjectName("GalleryQuickLabel")
        gallery_quick_layout.addWidget(gallery_type_label)

        self.gallery_type_filter = ArrowComboBox()
        self.gallery_type_filter.setObjectName("GalleryQuickCombo")
        self.gallery_type_filter.setMinimumWidth(130)
        self.gallery_type_filter.setMaximumWidth(170)
        self.gallery_type_filter.addItem("All")
        self.gallery_type_filter.currentTextChanged.connect(self._apply_gallery_quick_filters)
        gallery_quick_layout.addWidget(self.gallery_type_filter)

        self.gallery_thumb_filter_btn = QPushButton("Has thumbnail")
        self.gallery_thumb_filter_btn.setObjectName("GalleryQuickButton")
        self.gallery_thumb_filter_btn.setCheckable(True)
        self.gallery_thumb_filter_btn.clicked.connect(self._apply_gallery_quick_filters)
        gallery_quick_layout.addWidget(self.gallery_thumb_filter_btn)

        self.gallery_pinned_filter_btn = QPushButton("Pinned")
        self.gallery_pinned_filter_btn.setObjectName("GalleryQuickButton")
        self.gallery_pinned_filter_btn.setCheckable(True)
        self.gallery_pinned_filter_btn.clicked.connect(self._apply_gallery_quick_filters)
        gallery_quick_layout.addWidget(self.gallery_pinned_filter_btn)

        self.gallery_unsaved_filter_btn = QPushButton("Unsaved")
        self.gallery_unsaved_filter_btn.setObjectName("GalleryQuickButton")
        self.gallery_unsaved_filter_btn.setCheckable(True)
        self.gallery_unsaved_filter_btn.clicked.connect(self._apply_gallery_quick_filters)
        gallery_quick_layout.addWidget(self.gallery_unsaved_filter_btn)

        gallery_quick_layout.addStretch(1)
        self.gallery_quick_clear_btn = QPushButton("Clear quick")
        self.gallery_quick_clear_btn.setObjectName("GalleryQuickClear")
        self.gallery_quick_clear_btn.clicked.connect(self._clear_gallery_quick_filters)
        gallery_quick_layout.addWidget(self.gallery_quick_clear_btn)

        # Preserve the DEV splitter fix without changing the FIX46 visual row.
        # Keep overflow inside an invisible viewport so the right-side details
        # panel remains freely draggable/widenable.  The horizontal scrollbar is
        # deliberately hidden: showing it steals vertical space from the FIX46
        # row and makes the Gallery filter band taller than the release layout.
        self.gallery_quick_bar.setMinimumWidth(self.gallery_quick_bar.sizeHint().width())
        self.gallery_quick_scroll = QScrollArea()
        self.gallery_quick_scroll.setObjectName("GalleryQuickScroll")
        self.gallery_quick_scroll.setFrameShape(QFrame.NoFrame)
        self.gallery_quick_scroll.setWidgetResizable(True)
        self.gallery_quick_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.gallery_quick_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.gallery_quick_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.gallery_quick_scroll.setMinimumWidth(0)
        self.gallery_quick_scroll.setFixedHeight(self.gallery_quick_bar.sizeHint().height())
        self.gallery_quick_scroll.viewport().setAutoFillBackground(False)
        self.gallery_quick_scroll.setWidget(self.gallery_quick_bar)
        self.gallery_quick_scroll.setVisible(False)
        left_layout.addWidget(self.column_fold_bar, 0)
        left_layout.addWidget(self.gallery_quick_scroll, 0)
        left_layout.addWidget(self.pinned_bar, 0)

        self.library_stack = QStackedWidget()
        self.library_stack.setObjectName("LibraryStack")
        self.library_stack.addWidget(self.table)
        self.library_stack.addWidget(self.gallery)
        left_layout.addWidget(self.library_stack, 1)
        self.splitter.addWidget(left)

        # Right details card.
        right = QFrame()
        right.setObjectName("DetailPanel")
        # Slightly wider than the old compact pane, but still leaves most space to the Library.
        right.setMinimumWidth(300)
        # Keep the compact default, but let the user drag the splitter left to widen Edit/details when desired.
        right.setMaximumWidth(720)
        right_layout = QVBoxLayout(right)
        # Reclaim vertical space *above* Overview so the image frame itself
        # starts higher. Shrinking only the top padding within the photo mat
        # (as in v0.4.50) cannot accomplish that when the image fills the mat.
        right_layout.setContentsMargins(8, 4, 8, 7)
        right_layout.setSpacing(3)

        title_row = QHBoxLayout()
        title_row.setSpacing(5)
        self.detail_title = QLabel("Select a package")
        self.detail_title.setObjectName("DetailTitle")
        self.detail_title.setWordWrap(True)
        self.detail_title.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.detail_title.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.detail_title.setToolTip("Full package name")
        # The old 48px fixed title clipped filenames requiring more than two lines.
        # A compact scroll area expands for normal filenames and scrolls only if
        # the name is exceptionally long; the adjacent Copy / Pin buttons stay put.
        self.detail_title_scroll = QScrollArea()
        self.detail_title_scroll.setObjectName("DetailTitleScroll")
        self.detail_title_scroll.setFrameShape(QFrame.NoFrame)
        self.detail_title_scroll.setWidgetResizable(True)
        self.detail_title_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.detail_title_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.detail_title_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.detail_title_scroll.setFixedHeight(44)
        self.detail_title_scroll.setToolTip("Scroll here for unusually long filenames")
        title_content = QWidget()
        title_content.setObjectName("DetailTitleContent")
        title_content_layout = QVBoxLayout(title_content)
        title_content_layout.setContentsMargins(0, 0, 3, 0)
        title_content_layout.setSpacing(0)
        title_content_layout.addWidget(self.detail_title)
        self.detail_title_scroll.setWidget(title_content)
        self._detail_title_viewport_width = 0
        self.detail_title_scroll.viewport().installEventFilter(self)
        title_row.addWidget(self.detail_title_scroll, 1)
        self.copy_filename_btn = QPushButton("Copy name")
        self.copy_filename_btn.setObjectName("PinButton")
        self.copy_filename_btn.setEnabled(False)
        self.copy_filename_btn.setToolTip("Copy this CC file name to the clipboard")
        self.copy_filename_btn.clicked.connect(self._copy_current_filename)
        title_row.addWidget(self.copy_filename_btn)
        self.pin_button = QPushButton("Pin")
        self.pin_button.setObjectName("PinButton")
        self.pin_button.setEnabled(False)
        self.pin_button.setToolTip("Keep this package in the persistent Pinned strip")
        self.pin_button.clicked.connect(self._toggle_current_pin)
        title_row.addWidget(self.pin_button)
        right_layout.addLayout(title_row)

        # Keep the compact at-a-glance row from the distributed v1.0 layout:
        # Status, CC Type and Highest Polys. Health and Risk Score belong in the
        # Overview details below, avoiding duplicate information and freeing
        # vertical space for the thumbnail.
        badges = QGridLayout()
        badges.setHorizontalSpacing(4)
        badges.setVerticalSpacing(3)
        badges.setContentsMargins(3, 0, 0, 0)
        self.status_badge = QLabel("Status: —")
        self.cc_type_badge = QLabel("CC Type: —")
        self.polys_badge = QLabel("Highest Polys: —")
        for badge in (self.status_badge, self.cc_type_badge, self.polys_badge):
            badge.setObjectName("Badge")
        badges.addWidget(self.status_badge, 0, 0)
        badges.addWidget(self.cc_type_badge, 0, 1)
        badges.addWidget(self.polys_badge, 0, 2)
        badges.setColumnStretch(3, 1)
        self.random_sim_usage_badge = None
        if IS_DEV_BUILD:
            self.random_sim_usage_badge = QLabel("Sim use: —")
            self.random_sim_usage_badge.setObjectName("Badge")
            self.random_sim_usage_badge.setToolTip("Random Sim Studio collections that use this original CC package")
            self.random_sim_usage_badge.setMinimumWidth(0)
            self.random_sim_usage_badge.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
            self.random_sim_usage_badge.setVisible(False)
            badges.addWidget(self.random_sim_usage_badge, 2, 0, 1, 3)
        right_layout.addLayout(badges)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("DetailTabs")
        self.tabs.setMinimumWidth(0)
        right_layout.addWidget(self.tabs, 1)

        overview = QWidget()
        overview.setObjectName("OverviewPage")
        overview_layout = QVBoxLayout(overview)
        overview_layout.setContentsMargins(5, 1, 5, 5)
        overview_layout.setSpacing(7)

        self.thumbnail_card = QFrame()
        self.thumbnail_card.setObjectName("ThumbnailCard")
        thumb_layout = QVBoxLayout(self.thumbnail_card)
        thumb_layout.setContentsMargins(4, 0, 4, 4)
        self.thumbnail_label = ResponsiveThumbnailLabel("No thumbnail")
        self.thumbnail_label.setObjectName("ThumbnailImage")
        self.thumbnail_label.setAlignment(Qt.AlignCenter)
        # Make the preview the flexible, high-priority area. Do not enlarge its
        # minimum every time the title gains another line: the image will be
        # drawn to fit whatever *visible* height remains.
        self.thumbnail_label.setMinimumHeight(180)
        self.thumbnail_label.setMaximumHeight(1000)
        self.thumbnail_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
        self.thumbnail_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        thumb_layout.addWidget(self.thumbnail_label, 1)
        overview_layout.addWidget(self.thumbnail_card, 1)

        self.info_card = QFrame()
        self.info_card.setObjectName("InfoCard")
        info_layout = QVBoxLayout(self.info_card)
        info_layout.setContentsMargins(10, 9, 10, 9)
        info_layout.setSpacing(0)
        self.info_browser = QTextBrowser()
        self.info_browser.setObjectName("InfoBrowser")
        self.info_browser.setOpenExternalLinks(False)
        self.info_browser.anchorClicked.connect(self._open_random_sim_usage_link)
        info_layout.addWidget(self.info_browser)
        # Keep the info card compact because its QTextBrowser already scrolls.
        # Otherwise an expanding info area can take half the available height,
        # leaving the portrait preview clipped after a long title appears.
        # The browser inside already scrolls. Return some of its former height
        # to the thumbnail so the displayed photo can be bigger, never smaller.
        self.info_card.setMinimumHeight(95)
        self.info_card.setMaximumHeight(120)
        self.info_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        overview_layout.addWidget(self.info_card, 0)
        overview_layout.setStretch(0, 1)
        overview_layout.setStretch(1, 0)

        # If someone makes the window very short, scroll the Overview page as
        # a whole instead of showing a physically cut-off image/info panel.
        self.overview_scroll = QScrollArea()
        self.overview_scroll.setObjectName("OverviewScroll")
        self.overview_scroll.setWidgetResizable(True)
        self.overview_scroll.setFrameShape(QFrame.NoFrame)
        self.overview_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.overview_scroll.setWidget(overview)
        self.tabs.addTab(self.overview_scroll, "Overview")

        # Restored from the distributed v1.0 PUBLIC release. This tab presents
        # direct CASP-linked information separately from manual organisation.
        casp_page = QWidget()
        casp_page.setObjectName("CASPObservationsPage")
        casp_layout = QVBoxLayout(casp_page)
        casp_layout.setContentsMargins(8, 8, 8, 8)
        casp_layout.setSpacing(7)
        casp_intro = QLabel(
            "This tab shows direct CASP observations from the latest validated scan when an exact package identity link is available. "
            "CAS Curator does not guess from filenames, and older scanner classifications are not treated as new findings."
        )
        casp_intro.setWordWrap(True)
        casp_intro.setObjectName("EditHelp")
        casp_layout.addWidget(casp_intro)
        self.casp_observations_browser = QTextBrowser()
        self.casp_observations_browser.setObjectName("CASPObservationsBrowser")
        self.casp_observations_browser.setOpenExternalLinks(False)
        casp_layout.addWidget(self.casp_observations_browser, 1)
        self.tabs.addTab(casp_page, "CASP observations")

        edit_page = QWidget()
        edit_page.setObjectName("EditPage")
        edit_page_layout = QVBoxLayout(edit_page)
        edit_page_layout.setContentsMargins(0, 0, 0, 0)
        edit_page_layout.setSpacing(0)

        self.edit_scroll = QScrollArea()
        self.edit_scroll.setObjectName("EditScroll")
        self.edit_scroll.setWidgetResizable(True)
        self.edit_scroll.setFrameShape(QFrame.NoFrame)
        self.edit_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        # The scroll area must be allowed to shrink with the right-side pane in a
        # normal (non-maximized) window; the edit contents can scroll horizontally.
        self.edit_scroll.setMinimumWidth(0)
        self.edit_scroll.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)

        edit_content = QWidget()
        edit_content.setObjectName("EditScrollContent")
        # Do not impose a 500 px minimum on the whole side pane. The QScrollArea
        # handles wider edit controls with its horizontal scrollbar instead.
        edit_content.setMinimumWidth(0)
        edit_layout = QVBoxLayout(edit_content)
        edit_layout.setContentsMargins(9, 10, 9, 9)
        edit_layout.setSpacing(9)

        edit_intro = QLabel("Manual fields")
        edit_intro.setObjectName("EditTitle")
        edit_layout.addWidget(edit_intro)
        edit_note = QLabel(
            "Optional information you can add yourself. CAS Curator does not guess these values "
            "or move package files automatically."
        )
        edit_note.setWordWrap(True)
        edit_note.setObjectName("EditHelp")
        edit_layout.addWidget(edit_note)

        edit_guide_btn = QPushButton("Open Guide → Manual fields")
        edit_guide_btn.setObjectName("GuideJumpButton")
        edit_guide_btn.setToolTip("Open the Guide at the section explaining manual fields and saving changes.")
        edit_guide_btn.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        edit_guide_btn.clicked.connect(
            lambda checked=False: (
                self._set_app_page("guide"),
                self.guide_browser.scrollToAnchor("edit"),
            )
        )
        edit_layout.addWidget(edit_guide_btn, 0, Qt.AlignLeft)

        self.edit_card = QFrame()
        self.edit_card.setObjectName("EditCard")
        form = QFormLayout(self.edit_card)
        self.edit_form = form
        form.setContentsMargins(10, 10, 10, 10)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)

        self.status_edit = QComboBox()
        self.creator_edit = QLineEdit()
        self.creator_edit.setPlaceholderText("Creator / site name…")
        self.creator_edit.setToolTip("Optional manual creator name. CAS Curator does not guess this value automatically.")
        self.file_location_edit = QComboBox()
        self.file_location_edit.setToolTip(
            "Choose the folder containing the original, unmerged package. This records a location only; it does not move the file."
        )
        self.file_location_manage_btn = QPushButton("Manage Locations…")
        self.file_location_manage_btn.setToolTip("Add named folders and manage their saved paths in the Master Database.")
        self.file_location_manage_btn.setEnabled(False)
        self.file_location_open_btn = QPushButton("Open Folder")
        self.file_location_open_btn.setToolTip("Open the assigned folder in Windows File Explorer. Missing folders are reported; no substitute is chosen.")
        self.file_location_open_btn.setEnabled(False)
        file_location_tools = QWidget()
        # Keep the selected friendly name readable even when the details pane is narrow.
        # Management/open buttons sit below it instead of squeezing the combo to a few pixels.
        file_location_tools_layout = QVBoxLayout(file_location_tools)
        file_location_tools_layout.setContentsMargins(0, 0, 0, 0)
        file_location_tools_layout.setSpacing(5)
        file_location_tools_layout.addWidget(self.file_location_edit)
        file_location_button_row = QHBoxLayout()
        file_location_button_row.setContentsMargins(0, 0, 0, 0)
        file_location_button_row.setSpacing(5)
        file_location_button_row.addWidget(self.file_location_manage_btn)
        file_location_button_row.addWidget(self.file_location_open_btn)
        file_location_button_row.addStretch(1)
        file_location_tools_layout.addLayout(file_location_button_row)
        self.cc_magic_profiles_edit = QPlainTextEdit()
        self.cc_magic_profiles_edit.setObjectName("CCMagicProfilesEdit")
        self.cc_magic_profiles_edit.setReadOnly(True)
        self.cc_magic_profiles_edit.setPlaceholderText("No profiles assigned")
        self.cc_magic_profiles_edit.setToolTip(
            "Read-only list of Profiles currently assigned to this package. "
            "Use Add to / Choose Profile above to add membership; use the table Profile dropdown or fill-down for row changes."
        )
        self.cc_magic_profiles_edit.setMinimumHeight(44)
        self.cc_magic_profiles_edit.setMaximumHeight(150)
        self.cc_magic_profiles_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.cc_magic_profiles_edit.textChanged.connect(self._update_dynamic_editor_heights)
        self.notes_edit = QPlainTextEdit()
        self.notes_edit.setObjectName("NotesEdit")
        self.notes_edit.setPlaceholderText("Add a note…")
        # Notes stays compact for short text, grows with its content, and only
        # uses an internal scrollbar after the configured maximum number of lines.
        self.notes_edit.setMinimumHeight(44)
        self.notes_edit.setMaximumHeight(170)
        self.notes_edit.textChanged.connect(self._update_dynamic_editor_heights)
        form.addRow("Status:", self.status_edit)
        form.addRow("Creator:", self.creator_edit)
        form.addRow("File Location:", file_location_tools)

        # Merge Profiles are a separate managed grouping. A package can belong
        # to several merged collections at once; that assignment itself indicates
        # merge membership, so no separate "In Merge" Status is required.
        self.merge_profiles_edit = QPlainTextEdit()
        self.merge_profiles_edit.setObjectName("MergeProfilesEdit")
        self.merge_profiles_edit.setReadOnly(True)
        self.merge_profiles_edit.setPlaceholderText("No merge profiles assigned")
        self.merge_profiles_edit.setToolTip(
            "Merge Profiles record which merged collections include this package. "
            "They are separate from general Profiles and can have multiple assignments."
        )
        self.merge_profiles_edit.setMinimumHeight(44)
        self.merge_profiles_edit.setMaximumHeight(150)
        self.merge_profiles_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.merge_profiles_edit.textChanged.connect(self._update_dynamic_editor_heights)
        self.merge_profiles_set_btn = QPushButton("Set Merge Profiles…")
        self.merge_profiles_set_btn.setToolTip("Choose one or more Merge Profiles for this package.")
        self.merge_profiles_manage_btn = QPushButton("Manage Merge Profiles…")
        self.merge_profiles_manage_btn.setToolTip("Create, rename or delete names in the managed Merge Profiles list.")
        self.merge_profiles_tools = QWidget()
        merge_profiles_layout = QVBoxLayout(self.merge_profiles_tools)
        merge_profiles_layout.setContentsMargins(0, 0, 0, 0)
        merge_profiles_layout.setSpacing(5)
        merge_profiles_layout.addWidget(self.merge_profiles_edit)
        merge_profiles_buttons = QHBoxLayout()
        merge_profiles_buttons.setContentsMargins(0, 0, 0, 0)
        merge_profiles_buttons.setSpacing(5)
        merge_profiles_buttons.addWidget(self.merge_profiles_set_btn)
        merge_profiles_buttons.addWidget(self.merge_profiles_manage_btn)
        merge_profiles_buttons.addStretch(1)
        merge_profiles_layout.addLayout(merge_profiles_buttons)
        self.merge_profiles_label = QLabel("Merge Profiles:")
        form.addRow(self.merge_profiles_label, self.merge_profiles_tools)
        try:
            form.setRowVisible(self.merge_profiles_tools, False)
        except AttributeError:
            self.merge_profiles_label.setVisible(False)
            self.merge_profiles_tools.setVisible(False)

        self.preview_image_add_btn = QPushButton("Add / Replace image…")
        self.preview_image_clear_btn = QPushButton("Clear image")
        self.preview_image_add_btn.setToolTip("Choose a local image to use as this package's preview. A manual image takes priority over an extracted package thumbnail.")
        self.preview_image_clear_btn.setToolTip("Clear the current manual/linked preview. Run Scan again to restore an extractable package thumbnail.")
        preview_image_tools = QWidget()
        preview_image_tools_layout = QHBoxLayout(preview_image_tools)
        preview_image_tools_layout.setContentsMargins(0, 0, 0, 0)
        preview_image_tools_layout.setSpacing(5)
        preview_image_tools_layout.addWidget(self.preview_image_add_btn)
        preview_image_tools_layout.addWidget(self.preview_image_clear_btn)
        form.addRow("Preview image:", preview_image_tools)

        # PUBLIC profile assignment is intentionally simple: choose an existing
        # Profile to add to the current package. Bulk propagation is handled by
        # the table fill-down feature instead of separate selected/shown buttons.
        self.cc_magic_profile_quick = QComboBox()
        self.cc_magic_profile_quick.setEditable(False)
        self.cc_magic_profile_quick.setToolTip("Choose an existing Profile to add to this package.")
        form.addRow("Add to:", self.cc_magic_profile_quick)
        # Assignment area always uses plural because the control supports multiple Profiles.
        self.profiles_assigned_label = QLabel("Profiles assigned to this package:")
        form.addRow(self.profiles_assigned_label, self.cc_magic_profiles_edit)
        form.addRow("Notes:", self.notes_edit)
        edit_layout.addWidget(self.edit_card)

        self.edit_dirty_label = QLabel("No unsaved changes for this package")
        self.edit_dirty_label.setObjectName("DirtyLabel")
        self.edit_dirty_label.setWordWrap(True)
        edit_layout.addWidget(self.edit_dirty_label)

        edit_layout.addStretch(1)
        self.edit_scroll.setWidget(edit_content)
        edit_page_layout.addWidget(self.edit_scroll, 1)

        # Keep the main edit actions outside the horizontally scrolling content.
        # Revert stays left; Save is on the opposite/right side with modest padding,
        # so it remains visible at the narrowest supported details-pane width.
        edit_button_bar = QFrame()
        edit_button_bar.setObjectName("EditButtonBar")
        edit_buttons = QHBoxLayout(edit_button_bar)
        edit_buttons.setContentsMargins(9, 6, 10, 8)
        edit_buttons.setSpacing(8)
        self.revert_selected_btn = QPushButton("Revert this")
        self.revert_selected_btn.clicked.connect(self.revert_selected_changes)
        self.save_btn = QPushButton("Save Changes")
        self.save_btn.setObjectName("SaveButton")
        self.save_btn.clicked.connect(self.save_changes)
        edit_buttons.addWidget(self.revert_selected_btn)
        edit_buttons.addStretch(1)
        edit_buttons.addWidget(self.save_btn)
        edit_page_layout.addWidget(edit_button_bar)
        self.tabs.addTab(edit_page, "Edit")
        QTimer.singleShot(0, self._update_dynamic_editor_heights)

        self.status_edit.currentTextChanged.connect(self._status_editor_changed)
        self.creator_edit.textChanged.connect(self._editor_changed)
        self.file_location_edit.currentTextChanged.connect(self._editor_changed)
        self.file_location_manage_btn.clicked.connect(self.manage_file_locations)
        self.file_location_open_btn.clicked.connect(self.open_selected_file_location)
        self.merge_profiles_set_btn.clicked.connect(self._choose_merge_profiles_for_current)
        self.merge_profiles_manage_btn.clicked.connect(self._manage_merge_profiles)
        self.notes_edit.textChanged.connect(self._editor_changed)
        self.preview_image_add_btn.clicked.connect(self._choose_manual_preview_image)
        self.preview_image_clear_btn.clicked.connect(self._clear_manual_preview_image)
        self.cc_magic_profile_quick.activated.connect(self._cc_profile_quick_selected)
        self._set_editor_enabled(False)

        self.all_fields_page = QWidget()
        all_fields_page = self.all_fields_page
        af_layout = QVBoxLayout(all_fields_page)
        af_layout.setContentsMargins(7, 8, 7, 7)
        self.fields_table = QTableWidget(0, 2)
        self.fields_table.setHorizontalHeaderLabels(["Field", "Value"])
        self.fields_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.fields_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.fields_table.verticalHeader().setVisible(False)
        self.fields_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.fields_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.fields_table.setAlternatingRowColors(False)
        af_layout.addWidget(self.fields_table)
        self.all_fields_tab_index = self.tabs.addTab(all_fields_page, "All fields")
        self.tabs.setTabVisible(self.all_fields_tab_index, bool(IS_DEV_BUILD or self.advanced_features))
        self.splitter.addWidget(right)
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 0)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(7)
        self.splitter.setSizes([1192, 353])
        self.splitter.splitterMoved.connect(lambda *_: self._fit_detail_title())
        self._set_library_view(self.view_mode, persist=False)

        self.statusBar().showMessage("Open Sims3_CC_Master_Database.xlsx — editable manual fields: Status / File Location / Profiles / Notes")

    def _dialog_icon_pixmap(self, kind) -> QPixmap:
        """Small theme-aware message icon used instead of Windows' large stock symbols."""
        palette = getattr(self, "_dialog_palette", {}) or {}
        accent = QColor(str(palette.get("accent", self.theme.get("accent", "#7A668D"))))
        text_color = QColor("#FFFFFF" if _luma_hex(accent.name()) < 155 else "#202020")
        symbol = "i"
        if kind == QMessageBox.Warning:
            symbol = "!"
        elif kind == QMessageBox.Critical:
            symbol = "×"
        elif kind == QMessageBox.Question:
            symbol = "?"
        pix = QPixmap(42, 42)
        pix.fill(Qt.transparent)
        painter = QPainter(pix)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(Qt.NoPen)
        painter.setBrush(accent)
        painter.drawRoundedRect(3, 3, 36, 36, 11, 11)
        font = painter.font()
        font.setBold(True)
        font.setPointSize(18 if symbol != "i" else 17)
        painter.setFont(font)
        painter.setPen(text_color)
        painter.drawText(pix.rect().adjusted(0, -1, 0, 0), Qt.AlignCenter, symbol)
        painter.end()
        return pix

    def _dialog_stylesheet(self) -> str:
        palette = getattr(self, "_dialog_palette", {}) or {}
        bg = str(palette.get("background", "#F4F0F4"))
        panel = str(palette.get("panel", bg))
        inp = str(palette.get("input", "#FFFFFF"))
        text = str(palette.get("text", "#202020"))
        input_text = str(palette.get("input_text", text))
        button = str(palette.get("button", "#DDD2E2"))
        button_text = str(palette.get("button_text", text))
        border = str(palette.get("border", "#9E8BA7"))
        control_border = str(palette.get("control_border", border))
        accent = str(palette.get("accent", border))
        log = str(palette.get("log", inp))
        log_text = str(palette.get("log_text", input_text))
        # Message text deliberately sits on a soft neutral reading card. The
        # outer dialog, icon and buttons remain theme-aware, while the card is
        # kept consistently readable without the stark pure-white block used
        # in RC42. Use the same neutral treatment in every theme so dialogs
        # have one predictable visual language.
        message_card = "#F0EFF0"
        message_text = "#242224"
        # Keep a hint of the active theme in the outline, but soften it heavily
        # into the neutral card so it never becomes a harsh grey frame.
        message_border = _mix_hex(control_border, message_card, 0.58)
        return f"""
            QMessageBox, QDialog#ThemedDialog, QMainWindow#ThemedUtilityWindow, QWidget#ThemedUtilityPanel {{
                background: {bg}; color: {text}; font-family: 'Segoe UI'; font-size: 10pt;
            }}
            QMessageBox QLabel, QDialog#ThemedDialog QLabel, QMainWindow#ThemedUtilityWindow QLabel {{
                background: transparent; color: {text};
            }}
            QMessageBox QLabel#qt_msgbox_label, QMessageBox QLabel#qt_msgbox_informativelabel {{
                background: {message_card}; color: {message_text};
                border: 1px solid {message_border}; border-radius: 7px;
                padding: 8px 10px;
            }}
            QMessageBox QFrame#MessageHeaderCard {{
                background: {message_card}; border: 1px solid {message_border};
                border-radius: 7px;
            }}
            QMessageBox QFrame#MessageHeaderCard QLabel#qt_msgbox_label {{
                background: transparent; color: {message_text}; border: none;
                padding: 0px;
            }}
            QMessageBox QLabel#qt_msgboxex_icon_label {{
                background: transparent; border: none; padding: 0px;
            }}
            QMessageBox QPushButton, QDialog#ThemedDialog QPushButton, QMainWindow#ThemedUtilityWindow QPushButton {{
                background: {button}; color: {button_text}; border: 1px solid {control_border};
                border-radius: 6px; min-height: 28px; min-width: 82px; padding: 4px 11px;
            }}
            QMessageBox QPushButton:hover, QDialog#ThemedDialog QPushButton:hover, QMainWindow#ThemedUtilityWindow QPushButton:hover {{
                border-color: {accent};
            }}
            QMessageBox QPushButton:pressed, QDialog#ThemedDialog QPushButton:pressed, QMainWindow#ThemedUtilityWindow QPushButton:pressed {{
                background: {panel};
            }}
            QDialog#ThemedDialog QComboBox {{
                background: {inp}; color: {input_text}; border: 1px solid {control_border};
                border-radius: 6px; min-height: 28px; padding: 3px 38px 3px 8px;
            }}
            QDialog#ThemedDialog QComboBox::drop-down {{
                background: {button}; border: none; border-left: 1px solid {control_border};
                min-width: 34px;
            }}
            QDialog#ThemedDialog QComboBox::down-arrow {{ image: none; width: 0px; height: 0px; }}
            QDialog#ThemedDialog QAbstractItemView {{
                background: {inp}; color: {input_text}; selection-background-color: {button};
                selection-color: {button_text}; border: 1px solid {control_border};
            }}
            QPlainTextEdit#ExpandedScannerLog {{
                /* RC68: keep the expanded log as a neutral reading surface in every theme. */
                background: #E8EAED; color: #000000; border: 1px solid {control_border};
                border-radius: 7px; padding: 8px; font-family: Consolas; font-size: 10pt;
                selection-background-color: #BFC5CC; selection-color: #000000;
            }}
        """

    def _wrap_dialog_text_to_width(self, text: str, font, max_width: int) -> str:
        """Wrap plain QMessageBox text to a real pixel width.

        QLabel wordWrap alone is unreliable for QMessageBox labels once Qt has
        already calculated the native message-box geometry, and path-like
        tokens have no normal break points.  Insert real newlines so every
        themed dialog (information, warning, critical and question) uses the
        same compact reading width.  Existing paragraph/newline structure is
        preserved, and long paths/tokens are split without changing visible
        characters.
        """
        raw = str(text or "")
        if not raw or max_width <= 40:
            return raw
        # Leave rich-text messages alone; their own markup controls wrapping.
        if re.search(r"<\s*/?\w+[^>]*>", raw):
            return raw

        metrics = QFontMetrics(font)
        limit = max(60, int(max_width))

        def width(value: str) -> int:
            return metrics.horizontalAdvance(value)

        def split_long_token(token: str, first_limit: int) -> List[str]:
            pieces: List[str] = []
            remaining = token
            current_limit = max(35, first_limit)
            while remaining:
                if width(remaining) <= current_limit:
                    pieces.append(remaining)
                    break
                # Find the longest prefix that fits.
                lo, hi = 1, len(remaining)
                while lo < hi:
                    mid = (lo + hi + 1) // 2
                    if width(remaining[:mid]) <= current_limit:
                        lo = mid
                    else:
                        hi = mid - 1
                cut = max(1, lo)
                # Prefer a natural path/identifier boundary near the end of
                # the fitted prefix, but never make a tiny fragment merely to
                # reach one.
                preferred = -1
                floor = max(1, cut // 2)
                for idx in range(cut - 1, floor - 1, -1):
                    if remaining[idx] in "\\/_-.→":
                        preferred = idx + 1
                        break
                if preferred > 0:
                    cut = preferred
                pieces.append(remaining[:cut])
                remaining = remaining[cut:]
                current_limit = limit
            return pieces

        wrapped_paragraphs: List[str] = []
        for source_line in raw.split("\n"):
            if source_line == "":
                wrapped_paragraphs.append("")
                continue
            # Preserve leading indentation/bullets while wrapping the content.
            leading_match = re.match(r"^\s*", source_line)
            leading = leading_match.group(0) if leading_match else ""
            content = source_line[len(leading):]
            words = re.findall(r"\S+", content)
            if not words:
                wrapped_paragraphs.append(source_line)
                continue

            out_lines: List[str] = []
            current = leading
            for word in words:
                separator = "" if not current.strip() else " "
                candidate = current + separator + word
                if width(candidate) <= limit:
                    current = candidate
                    continue

                if current.strip():
                    out_lines.append(current.rstrip())
                    current = leading

                available = max(35, limit - width(leading))
                if width(leading + word) <= limit:
                    current = leading + word
                    continue

                chunks = split_long_token(word, available)
                for chunk in chunks[:-1]:
                    out_lines.append((leading + chunk).rstrip())
                current = leading + chunks[-1]

            if current or not out_lines:
                out_lines.append(current.rstrip())
            wrapped_paragraphs.extend(out_lines)
        return "\n".join(wrapped_paragraphs)

    def _prepare_dialog_label(self, label: Optional[QLabel], width: int):
        """Force a QMessageBox label to wrap within *width* logical px."""
        if label is None:
            return
        width = max(80, int(width))
        label.setWordWrap(True)
        label.setMinimumWidth(0)
        label.setMaximumWidth(width)
        # Ignore QLabel's one-line minimum-size hint. This is the part Qt's
        # native QMessageBox layout otherwise honours and clips against.
        policy = QSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        policy.setHeightForWidth(True)
        label.setSizePolicy(policy)
        wrapped = self._wrap_dialog_text_to_width(label.text(), label.font(), width)
        if wrapped != label.text():
            label.setText(wrapped)
        label.setFixedWidth(width)
        label.updateGeometry()

    def _finalize_message_box_geometry(self, box: QMessageBox):
        """Recalculate a styled message box after its text was reflowed."""
        try:
            if not box.isVisible():
                return
            layout = box.layout()
            if layout is not None:
                layout.invalidate()
                layout.activate()
            box.adjustSize()
            box.updateGeometry()
        except RuntimeError:
            # The user can dismiss a very short-lived popup before the queued
            # geometry pass runs.
            pass

    def _style_message_box(self, box: QMessageBox):
        if box.property("casCuratorThemed"):
            return
        box.setProperty("casCuratorThemed", True)
        box.setWindowIcon(self.windowIcon())
        kind = box.icon()
        if kind != QMessageBox.NoIcon:
            box.setIconPixmap(self._dialog_icon_pixmap(kind))
        box.setStyleSheet(self._dialog_stylesheet())

        # Keep the proven RC44 Library/header layout completely untouched.
        # Only the QMessageBox composition changes here: the compact CAS
        # Curator icon sits inside the heading card, and the heading, reading
        # card and buttons form one centred stack.
        icon_label = box.findChild(QLabel, "qt_msgboxex_icon_label")
        title_label = box.findChild(QLabel, "qt_msgbox_label")
        info_label = box.findChild(QLabel, "qt_msgbox_informativelabel")
        layout = box.layout()

        if (isinstance(layout, QGridLayout) and icon_label is not None
                and title_label is not None and not box.property("casCuratorHeaderWrapped")):
            box.setProperty("casCuratorHeaderWrapped", True)
            layout.removeWidget(icon_label)
            layout.removeWidget(title_label)

            detail_text = (info_label.text() if info_label is not None
                           else title_label.text())
            # Keep centred dialogs compact.  RC48 proved the centred stack, but
            # 600 logical pixels made long prompts feel like a banner on some
            # Windows DPI settings.  RC63 trims the long/medium cards a little
            # further while retaining wrapping and enough room for the buttons.
            card_width = 450 if len(detail_text) > 150 else (430 if len(detail_text) > 70 else 400)

            header = QFrame(box)
            header.setObjectName("MessageHeaderCard")
            header.setFixedWidth(card_width)
            header_layout = QHBoxLayout(header)
            header_layout.setContentsMargins(10, 7, 10, 7)
            header_layout.setSpacing(9)

            icon_label.setParent(header)
            icon_label.setFixedSize(42, 42)
            icon_label.setAlignment(Qt.AlignCenter)
            icon_label.setContentsMargins(0, 0, 0, 0)
            title_label.setParent(header)
            title_width = max(180, card_width - 71)
            self._prepare_dialog_label(title_label, title_width)
            # Keep the icon visually attached to the opening line on long
            # confirmation prompts. Short information boxes still look best
            # vertically centred, but centring a 42 px icon beside several
            # wrapped lines makes it appear to float halfway down the message.
            icon_alignment = Qt.AlignTop if len(detail_text) > 150 else Qt.AlignVCenter
            header_layout.addWidget(icon_label, 0, icon_alignment)
            header_layout.addWidget(title_label, 1, Qt.AlignVCenter)
            # QMessageBox uses an internal multi-column grid.  Spanning only
            # the first two columns centres the cards inside the left portion
            # of the dialog, not the full window.  Span every current grid
            # column so both cards share the actual dialog centre line.
            full_span = max(1, layout.columnCount())
            layout.addWidget(header, 0, 0, 1, full_span, Qt.AlignHCenter)

            if info_label is not None:
                layout.removeWidget(info_label)
                info_label.setParent(box)
                # The label itself owns the reading-card border/padding, so
                # leave a little internal room when pre-wrapping its text.
                self._prepare_dialog_label(info_label, card_width)
                info_label.setText(
                    self._wrap_dialog_text_to_width(
                        info_label.text(), info_label.font(), max(180, card_width - 28)
                    )
                )
                info_label.updateGeometry()
                layout.addWidget(info_label, 1, 0, 1, full_span, Qt.AlignHCenter)

        if isinstance(layout, QGridLayout):
            layout.setContentsMargins(18, 16, 18, 16)
            layout.setHorizontalSpacing(0)
            layout.setVerticalSpacing(10)

        button_box = box.findChild(QDialogButtonBox)
        if button_box is not None:
            button_box.setCenterButtons(True)

        # QMessageBox calculates its native size before the Show event. Reflow
        # above therefore needs one explicit geometry pass now and one queued
        # pass after Qt has finished showing the window. This prevents the
        # pre-RC70 behaviour where text had wrapping enabled but the already-
        # sized label/card still clipped the right edge.
        self._finalize_message_box_geometry(box)
        QTimer.singleShot(0, lambda b=box: self._finalize_message_box_geometry(b))

    def _style_custom_dialog(self, dialog: QDialog):
        dialog.setObjectName("ThemedDialog")
        dialog.setWindowIcon(self.windowIcon())
        dialog.setStyleSheet(self._dialog_stylesheet())

    def apply_theme(self, name: str):
        self.theme_name = name
        self.theme = self.themes[name]
        self.settings.setValue("theme", name)
        t = self.theme

        # General light-theme roles.
        top_fill = _mix_hex(t["accent_soft"], t["accent"], 0.10)
        side_fill = top_fill
        header_fill = _mix_hex(top_fill, "#F5F5F5", 0.42)
        info_fill = _mix_hex(top_fill, "#FFFFFF", 0.55)
        filter_button = info_fill
        scrollbar_color = t["accent"]
        scrollbar_hover = t["selection"]
        thumb_fill = _mix_hex(t["thumb"], "#686868", 0.30)
        thumb_image_border = "#D5D5D5"

        # Guide surfaces use a light tint from the active theme instead of a
        # plain white document.  Keep them deliberately light so guide text
        # stays readable across both light and dark/cool themes.
        guide_page_fill = _mix_hex(t["accent_soft"], "#FFFFFF", 0.68)
        guide_card_fill = _mix_hex(t["accent_soft"], "#FFFFFF", 0.48)
        guide_browser_fill = _mix_hex(t["accent_soft"], "#FFFFFF", 0.78)
        guide_border = _mix_hex(t["accent"], "#B8B8B8", 0.48)

        # Slightly darker neutral table body across all themes.
        main_fill = "#FAFAFA"
        row_alt = "#EEEEEE"
        row_line = "#D6D6D6"

        # Neutral themes: balanced cool structure with softer contrast surfaces.
        if self.theme_name == "Shades of Grey":
            # Deliberately darker than the old silver/near-white version.
            top_fill = "#858A90"
            side_fill = "#989DA3"
            header_fill = "#A5AAB0"
            info_fill = "#B9BDC1"
            filter_button = "#ABB0B5"
            scrollbar_color = "#626970"
            scrollbar_hover = "#535A61"
            thumb_fill = "#858A90"
            # Keep the darker neutral chrome, but preserve the lighter
            # alternating Library rows from the earlier version for readability.
            main_fill = "#FAFAFA"
            row_alt = "#EEEEEE"
            row_line = "#D6D6D6"
            guide_page_fill = "#A5AAAF"
            guide_card_fill = "#969BA0"
            guide_browser_fill = "#B8BCC0"
        elif self.theme_name == "Carbon Blue":
            top_fill = "#8FA4B4"
            side_fill = top_fill
            header_fill = "#BAC8D2"
            info_fill = "#DCE4EA"
            filter_button = info_fill
            scrollbar_color = "#506B82"
            scrollbar_hover = "#405A70"

        # Rose Grey: cool grey structure with restrained dusty-rose accents.
        elif self.theme_name == "Rose Grey":
            top_fill = "#CFC8CF"
            side_fill = "#D8D1D8"
            header_fill = "#DED6DE"
            info_fill = "#E7C8D3"
            filter_button = "#DDA9BC"
            scrollbar_color = "#C97FA0"
            scrollbar_hover = "#B96B90"

        elif self.theme_name == "Cherry Blossom Pink":
            top_fill = "#F2AEBC"
            side_fill = top_fill
            header_fill = "#F7CBD4"
            info_fill = "#FBE1E7"
            filter_button = "#D88AA0"
            scrollbar_color = "#C97591"
            scrollbar_hover = "#B76580"
        elif self.theme_name == "Strawberry Milk":
            top_fill = "#F4B7C9"
            side_fill = top_fill
            header_fill = "#F8D8E1"
            info_fill = "#FBE9EE"
            filter_button = info_fill
            scrollbar_color = "#D96F91"
            scrollbar_hover = "#C45A7D"
        elif self.theme_name == "Jellyfish Glow":
            top_fill = "#7997E6"
            side_fill = top_fill
            header_fill = "#AAB8EE"
            info_fill = "#F5D3DF"
            filter_button = "#EAAFC6"
            scrollbar_color = "#D989AE"
            scrollbar_hover = "#C86F98"
        elif self.theme_name == "Celestial Lilac":
            top_fill = "#B39BC9"
            side_fill = top_fill
            header_fill = "#CEC0DB"
            info_fill = "#E7DDED"
            filter_button = "#D6C6E2"
            scrollbar_color = "#4E7AB1"
            scrollbar_hover = "#3E679A"
        elif self.theme_name == "Plum Ice":
            # FIX46: Library/Main page had become too grey-brown and muted.
            # Keep the cool Plum Ice identity, but clean the chrome toward a
            # soft pink-lilac so it matches the prettier scanner controls.
            top_fill = "#D9BDCC"
            side_fill = "#E2CCD7"
            header_fill = "#CA9DB3"
            info_fill = "#E8D2DC"
            filter_button = "#DDB8C9"
            scrollbar_color = "#7A4E64"
            scrollbar_hover = "#8E6278"
        elif self.theme_name == "Lotus Pond":
            top_fill = "#DDE2BE"
            side_fill = top_fill
            header_fill = "#EAEBCF"
            info_fill = "#F5D7DF"
            filter_button = "#EAB8C5"
            scrollbar_color = "#839958"
            scrollbar_hover = "#6F8448"
        elif self.theme_name == "Berry Cloud":
            top_fill = "#E6D6E6"
            side_fill = "#EEDFE8"
            header_fill = "#E8D3E0"
            info_fill = "#D7B0C6"
            filter_button = "#C78EAF"
            scrollbar_color = "#B66C98"
            scrollbar_hover = "#A85A89"
        elif self.theme_name == "Sea Glass":
            top_fill = "#B0E0E6"
            side_fill = "#CAF0F8"
            header_fill = "#D8F3F6"
            info_fill = "#D9F0F2"
            filter_button = "#82BCC3"
            scrollbar_color = "#4682B4"
            scrollbar_hover = "#376F9C"
        elif self.theme_name == "Blue Coral":
            top_fill = "#94C4C1"
            side_fill = "#DDEDEC"
            header_fill = "#DDEDEC"
            info_fill = "#F9C7BE"
            filter_button = "#F1AFA5"
            scrollbar_color = "#42899B"
            scrollbar_hover = "#337487"
        elif self.theme_name == "Aqua Depth":
            top_fill = "#48BBC4"
            side_fill = "#D8F2F4"
            header_fill = "#C8EBEE"
            info_fill = "#BDE7EA"
            filter_button = "#4FA3AE"
            scrollbar_color = "#006C83"
            scrollbar_hover = "#00596C"
        elif self.theme_name == "Periwinkle Mist":
            top_fill = "#C8BFE7"
            side_fill = "#E8E2F2"
            header_fill = "#E1D8EE"
            info_fill = "#D9D8F0"
            filter_button = "#AEB9DF"
            scrollbar_color = "#6C7EB8"
            scrollbar_hover = "#5C6DA7"
        elif self.theme_name == "Atlantic":
            top_fill = "#8FB6DD"
            side_fill = "#E6EEF7"
            header_fill = "#D5E2F2"
            info_fill = "#C7D8F0"
            filter_button = "#6F9DCC"
            scrollbar_color = "#174E8A"
            scrollbar_hover = "#113F70"
        elif self.theme_name == "Ice Lagoon":
            top_fill = "#9AD3DE"
            side_fill = "#F2FBFC"
            header_fill = "#E4F5F7"
            info_fill = "#D8F0F3"
            filter_button = "#A8D9DF"
            scrollbar_color = "#3A8795"
            scrollbar_hover = "#2D7481"
        elif self.theme_name == "Jelly Blue Pink":
            top_fill = "#7997E6"
            side_fill = "#CAA9F3"
            header_fill = "#D8D8F6"
            info_fill = "#F2C2DA"
            filter_button = "#DFA9C6"
            scrollbar_color = "#206ABC"
            scrollbar_hover = "#17589E"
        elif self.theme_name == "Seafoam Pearl":
            top_fill = "#8FD2CA"
            side_fill = "#DFF5F3"
            header_fill = "#D2ECEA"
            info_fill = "#BFE8E3"
            filter_button = "#92CAC5"
            scrollbar_color = "#3E7C7B"
            scrollbar_hover = "#326968"
        elif self.theme_name == "Blue Grey Rose":
            top_fill = "#A9B9C8"
            side_fill = "#D8E0E7"
            header_fill = "#CCD8E2"
            # Cleaner rose accents: less grey/muddy than the earlier dusty pink.
            info_fill = "#F2CEDA"
            filter_button = "#E8B4C5"
            scrollbar_color = "#6E89A3"
            scrollbar_hover = "#5C748C"
        elif self.theme_name == "Blue Porcelain":
            top_fill = "#DCEAF6"
            side_fill = "#F7FBFE"
            header_fill = "#E9F2F9"
            info_fill = "#EDF5FB"
            filter_button = "#B9D2E8"
            scrollbar_color = "#5E8FC3"
            scrollbar_hover = "#4D7DB0"

        # True dark themes: dark/intense top + side surfaces, lighter contrast card/button.
        elif self.theme_name == "Navy Bloom":
            top_fill = "#1F3C5A"
            side_fill = "#2C4F73"
            header_fill = "#4F7096"
            info_fill = "#D7E2F0"
            filter_button = "#9BB8DB"
            scrollbar_color = "#5D84B8"
            scrollbar_hover = "#7297C6"
            thumb_fill = "#344D68"
        elif self.theme_name == "Storm Purple":
            top_fill = "#50506D"
            side_fill = top_fill
            header_fill = "#696B88"
            info_fill = "#7E809B"
            filter_button = "#A9AAC0"
            scrollbar_color = "#8E90AA"
            scrollbar_hover = "#A4A6BC"
            thumb_fill = "#454553"
        elif self.theme_name == "Retro Terminal":
            top_fill = "#213B34"
            side_fill = top_fill
            header_fill = "#31574B"
            info_fill = "#416B5E"
            filter_button = "#86B0A1"
            scrollbar_color = "#659485"
            scrollbar_hover = "#7EAA9B"
            thumb_fill = "#303936"
        elif self.theme_name == "Midnight Steel":
            top_fill = "#172033"
            side_fill = top_fill
            header_fill = "#29394F"
            info_fill = "#34465E"
            filter_button = "#8195AC"
            scrollbar_color = "#5F748D"
            scrollbar_hover = "#7388A0"
            thumb_fill = "#303742"
        elif self.theme_name == "Deep Ocean":
            top_fill = "#103B47"
            side_fill = top_fill
            header_fill = "#215966"
            info_fill = "#2F6F7D"
            filter_button = "#82ADB5"
            scrollbar_color = "#4F8792"
            scrollbar_hover = "#6599A2"
            thumb_fill = "#2F3D40"
        elif self.theme_name == "Burgundy Ink":
            top_fill = "#492632"
            side_fill = top_fill
            header_fill = "#603643"
            info_fill = "#754553"
            filter_button = "#B48A94"
            scrollbar_color = "#95606C"
            scrollbar_hover = "#A87580"
            thumb_fill = "#3B3134"
        elif self.theme_name == "Obsidian Blue":
            top_fill = "#1E2734"
            side_fill = top_fill
            header_fill = "#2E3E52"
            info_fill = "#3D5068"
            filter_button = "#879AB1"
            scrollbar_color = "#637C99"
            scrollbar_hover = "#7890AA"
            thumb_fill = "#323A45"
        elif self.theme_name == "Black Cherry":
            top_fill = "#351E28"
            side_fill = top_fill
            header_fill = "#4C2B39"
            info_fill = "#63394A"
            filter_button = "#B47E91"
            scrollbar_color = "#8A5368"
            scrollbar_hover = "#A0647B"
            thumb_fill = "#3A3034"
        elif self.theme_name == "Dark Teal":
            top_fill = "#173A3A"
            side_fill = top_fill
            header_fill = "#24514F"
            info_fill = "#2F6663"
            # Cooler/lighter aqua controls make Dark Teal feel cleaner and more
            # modern without changing its dark teal panel structure.
            filter_button = "#9BC5C1"
            scrollbar_color = "#4F8E91"
            scrollbar_hover = "#68A6A8"
            thumb_fill = "#2F3C3B"
        elif self.theme_name == "Espresso Copper":
            top_fill = "#E2D0C2"
            side_fill = top_fill
            header_fill = "#EEE1D7"
            info_fill = "#F5EEE8"
            filter_button = "#C99163"
            scrollbar_color = "#8A5A3B"
            scrollbar_hover = "#A46A46"
            thumb_fill = "#6F6763"
        elif self.theme_name == "Charcoal Copper":
            top_fill = "#302B29"
            side_fill = top_fill
            header_fill = "#463A35"
            info_fill = "#5A4941"
            filter_button = "#C28E6C"
            scrollbar_color = "#A06F52"
            scrollbar_hover = "#B58060"
            thumb_fill = "#383432"
        elif self.theme_name == "Night Plum":
            top_fill = "#2C2437"
            side_fill = top_fill
            header_fill = "#443650"
            info_fill = "#574863"
            filter_button = "#AA91B7"
            scrollbar_color = "#826E93"
            scrollbar_hover = "#967FA6"
            thumb_fill = "#37323D"
        else:
            # Remaining themes stay light but use visible contrast between surfaces.
            top_fill = _mix_hex(t["accent_soft"], t["accent"], 0.12)
            side_fill = top_fill
            header_fill = _mix_hex(top_fill, t["accent"], 0.16)
            info_fill = _mix_hex(top_fill, "#FFFFFF", 0.52)
            filter_button = _mix_hex(info_fill, t["accent"], 0.16)

        filter_panel_fill = top_fill
        table_header_fill = header_fill
        # Downloaded private palettes can provide direct surface roles. This keeps
        # their original colour relationships instead of forcing them through the
        # built-in-theme formula above.
        if t.get("_custom"):
            top_fill = str(t.get("ui_top_fill", top_fill))
            side_fill = str(t.get("ui_side_fill", side_fill))
            header_fill = str(t.get("ui_header_fill", header_fill))
            info_fill = str(t.get("ui_info_fill", info_fill))
            filter_button = str(t.get("ui_filter_button", filter_button))
            scrollbar_color = str(t.get("ui_scrollbar_color", scrollbar_color))
            scrollbar_hover = str(t.get("ui_scrollbar_hover", scrollbar_hover))
            thumb_fill = str(t.get("ui_thumb_fill", thumb_fill))
            filter_panel_fill = str(t.get("ui_filter_panel_fill", top_fill))
            table_header_fill = str(t.get("ui_table_header_fill", header_fill))
            main_fill = str(t.get("ui_main_fill", main_fill))
            row_alt = str(t.get("ui_row_alt", row_alt))
        table_header_text = _button_text_for(table_header_fill, dark=t["text"], light="#FFFFFF")
        surface_text = _button_text_for(top_fill, dark=t["text"], light="#FFFFFF")
        header_text = _button_text_for(header_fill, dark=t["text"], light="#FFFFFF")
        info_text = _button_text_for(info_fill, dark=t["text"], light="#FFFFFF")
        # Detail-card roles can differ from the general info surface. Shades of
        # Grey benefits from a lighter Overview card while keeping the rest of
        # its neutral chrome intact.
        detail_info_fill = info_fill
        detail_info_text = info_text
        detail_title_text = surface_text
        if self.theme_name == "Shades of Grey":
            detail_info_fill = "#D8DBDE"
            detail_info_text = "#171717"
            detail_title_text = "#171717"
        if t.get("_custom"):
            detail_info_fill = str(t.get("ui_info_card_fill", detail_info_fill))
            detail_info_text = _button_text_for(detail_info_fill, dark=t["text"], light="#FFFFFF")
        button_text = _button_text_for(filter_button, dark=t["text"], light="#FFFFFF")
        arrow_color = _button_text_for(filter_button, dark="#3F3F3F", light="#FFFFFF")
        badge_fill = "#FFFFFF"
        badge_text = t["text"]
        tab_fill = header_fill
        tab_text = header_text
        tab_selected_fill = "#FFFFFF"
        tab_selected_text = t["text"]
        info_card_border = "#D2D2D2"
        thumb_card_border = "#D5D5D5"
        preview_mat_fill = "#FCFCFC"
        preview_mat_border = "#E7E7E7"

        # The preview is intentionally lighter and flatter than the old grey frame.
        # The outer thumbnail wrapper blends into the side panel; only the image mat
        # gets a very thin outline. Dark themes use a brighter outline for clarity.
        if _luma_hex(side_fill) < 145:
            preview_mat_fill = "#F7F7F7"
            preview_mat_border = "#F0F0F0"
        if t.get("_custom") and t.get("ui_thumb_fill"):
            preview_mat_fill = str(t["ui_thumb_fill"])
            preview_mat_border = str(t.get("ui_info_card_border", preview_mat_border))

        if self.theme_name in {"Rose Grey", "Strawberry Milk", "Jellyfish Glow", "Lotus Pond"}:
            badge_fill = filter_button
            badge_text = _button_text_for(badge_fill, dark=t["text"], light="#FFFFFF")

        if self.theme_name == "Lotus Pond":
            tab_fill = filter_button
            tab_text = _button_text_for(tab_fill, dark=t["text"], light="#FFFFFF")
            tab_selected_fill = info_fill
            tab_selected_text = _button_text_for(tab_selected_fill, dark=t["text"], light="#FFFFFF")

        if self.theme_name in {"Rose Grey", "Strawberry Milk", "Jellyfish Glow", "Lotus Pond", "Cherry Blossom Pink", "Berry Cloud", "Sakura Tea", "Cafe Rose", "Soft Blush Slate", "Lavender Cream", "Sunset Palace", "Celestial Lilac", "Wisteria Aqua", "Sea Glass", "Blue Coral", "Aqua Depth", "Periwinkle Mist", "Atlantic", "Ice Lagoon", "Jelly Blue Pink", "Seafoam Pearl", "Blue Grey Rose", "Blue Porcelain"}:
            info_card_border = "#FFFFFF"

        if self.theme_name == "Rose Grey":
            info_card_border = "#FFFFFF"

        if t.get("_custom"):
            tab_fill = str(t.get("ui_tab_fill", tab_fill))
            tab_selected_fill = str(t.get("ui_tab_selected_fill", tab_selected_fill))
            tab_text = _button_text_for(tab_fill, dark=t["text"], light="#FFFFFF")
            tab_selected_text = _button_text_for(tab_selected_fill, dark=t["text"], light="#FFFFFF")
            info_card_border = str(t.get("ui_info_card_border", info_card_border))

        # Border policy shared by Guide + Scan. These families looked best
        # with their original light/white outlines, so restore them explicitly.
        blue_sea_themes = {"Sea Glass", "Aqua Depth", "Atlantic", "Ice Lagoon", "Seafoam Pearl", "Blue Porcelain"}
        blue_purple_themes = {"Lavender Cream", "Jellyfish Glow", "Sunset Palace", "Celestial Lilac", "Wisteria Aqua", "Periwinkle Mist"}
        nature_themes = {"Lotus Pond"}
        dark_cool_themes = {"Storm Purple", "Retro Terminal", "Midnight Steel", "Deep Ocean", "Burgundy Ink",
                            "Obsidian Blue", "Black Cherry", "Dark Teal", "Night Plum", "Plum Ice", "Navy Bloom"}
        pastel_white_border_themes = {"Sakura Tea", "Cafe Rose", "Soft Blush Slate"}
        white_border_themes = blue_sea_themes | blue_purple_themes | nature_themes | dark_cool_themes | pastel_white_border_themes
        scan_card_border = "#CFCFCF"
        scan_row_border = info_card_border
        scan_border_width = 1
        guide_border_width = 1

        if self.theme_name in white_border_themes:
            structural_border = "#FFFFFF"
            info_card_border = structural_border
            guide_border = structural_border
            scan_card_border = structural_border
            scan_row_border = structural_border
        else:
            if self.theme_name == "Espresso Copper":
                # Coffee needs a clearer, slightly heavier outline.
                structural_border = "#8A654D"
                scan_border_width = 2
                guide_border_width = 2
            elif self.theme_name == "Shades of Grey":
                structural_border = "#5F656B"
            else:
                structural_border = _mix_hex(t["accent"], "#777777", 0.58)
            info_card_border = structural_border
            guide_border = structural_border
            scan_card_border = structural_border
            scan_row_border = structural_border

        # Custom studio borders must win over the built-in theme-family border
        # policy. v0.4.43 accidentally overwritten this imported role.
        if t.get("_custom") and t.get("ui_info_card_border"):
            info_card_border = str(t["ui_info_card_border"])
            scan_card_border = info_card_border
            scan_row_border = info_card_border

        # Scanner surfaces use the same theme-tinted treatment as PUBLIC so DEV
        # does not fall back to large, glaring white controls.
        scan_page_fill = _mix_hex(side_fill, info_fill, 0.70)
        if _luma_hex(side_fill) < 145:
            scan_row_fill = _mix_hex(side_fill, header_fill, 0.42)
            scan_input_fill = _mix_hex(side_fill, header_fill, 0.72)
            scan_log_fill = _mix_hex(side_fill, header_fill, 0.34)
            scan_progress_fill = _mix_hex(side_fill, header_fill, 0.52)
            scan_button_fill = _mix_hex(header_fill, filter_button, 0.24)
        else:
            scan_row_fill = _mix_hex(side_fill, header_fill, 0.48)
            scan_input_fill = _mix_hex(info_fill, scrollbar_color, 0.20)
            scan_log_fill = _mix_hex(info_fill, scrollbar_color, 0.12)
            scan_progress_fill = _mix_hex(info_fill, scrollbar_color, 0.14)
            scan_button_fill = _mix_hex(info_fill, scrollbar_color, 0.10)
        if self.theme_name == "Shades of Grey":
            # Keep Shades of Grey neutral but make the scanner cards noticeably
            # lighter than v0.4.39; the previous medium-grey blocks felt heavy.
            scan_page_fill = _mix_hex(scan_page_fill, "#FFFFFF", 0.30)
            scan_row_fill = _mix_hex(scan_row_fill, "#FFFFFF", 0.24)
            scan_input_fill = _mix_hex(scan_input_fill, "#FFFFFF", 0.12)
            scan_log_fill = _mix_hex(scan_log_fill, "#FFFFFF", 0.24)
            scan_progress_fill = _mix_hex(scan_progress_fill, "#FFFFFF", 0.24)
            scan_button_fill = _mix_hex(scan_button_fill, "#FFFFFF", 0.12)
        elif self.theme_name == "Blue Grey Rose":
            # Keep Scan controls warm but muted against the blue-grey shell.
            # FIX14 deliberately avoids the brighter candy-pink look while also
            # removing the earlier blue/mauve cast from the path/value fields.
            scan_input_fill = "#D8B6C0"
            scan_button_fill = "#D4A3B1"

        scan_card_fill = side_fill

        # FIX42: scanner-only depth polish.  Keep Library/Main page roles
        # untouched and give each problem Scan theme its own small tonal ladder
        # (page -> section card -> setup row -> control) instead of relying on
        # pale outlines between nearly identical fills. Strawberry Milk remains
        # unchanged and acts as a light-theme reference.
        scan_depth_palettes = {
            # Light / pastel scanner pages
            "Plum Ice": {
                # Scanner-only: keep the cool plum shell, but give the working
                # controls a cleaner dusty-pink/lilac lift so the page feels
                # fresher and matches the Library chrome without becoming bright.
                "page": "#E9DEE4", "card": "#C9ACBA", "row": "#DCC4CF",
                "input": "#EDDBE4", "button": "#DEB7C9", "log": "#DFCBD4",
                "progress": "#D7C0CB", "card_border": "#7B4E5E", "row_border": "#9F7485",
            },
            "Aqua Depth": {
                "page": "#CDEBED", "card": "#A9D8DC", "row": "#B7E0E3",
                "input": "#D9EFF0", "button": "#75B6BE", "log": "#B9DEE1",
                "progress": "#AED7DB", "card_border": "#55949C", "row_border": "#73AAB0",
            },
            "Blue Grey Rose": {
                "page": "#D7E0E8", "card": "#C8D5DF", "row": "#D4DEE6",
                "input": "#F0D3DC", "button": "#E4B6C5", "log": "#D9D7DF",
                "progress": "#CDD8E1", "card_border": "#738A9C", "row_border": "#8FA0AE",
            },
            "Carbon Blue": {
                "page": "#D5DEE5", "card": "#A9BBC9", "row": "#B9C8D3",
                "input": "#D2DCE4", "button": "#B2C4D1", "log": "#BDC9D2",
                "progress": "#B2C1CC", "card_border": "#71899C", "row_border": "#8298A8",
            },
            "Berry Cloud": {
                "page": "#E8DDE8", "card": "#DCC2D3", "row": "#E6CFDC",
                "input": "#F0E0E8", "button": "#D2A6BD", "log": "#E4CFDB",
                "progress": "#DBC2D0", "card_border": "#9F6F8C", "row_border": "#B48AA3",
            },
            "Cherry Blossom Pink": {
                "page": "#FBE0E6", "card": "#EDA0B2", "row": "#F3BAC8",
                "input": "#F8D8E0", "button": "#F2C2CD", "log": "#F4C4CF",
                "progress": "#F0B8C6", "card_border": "#B76F84", "row_border": "#C58A9A",
            },

            # Dark / cool scanner pages.  Changes are deliberately restrained:
            # retain each theme's identity, but separate the nested surfaces.
            "Midnight Steel": {
                "page": "#172033", "card": "#203047", "row": "#2B3E56",
                "input": "#3A506A", "button": "#6F849D", "log": "#26384F",
                "progress": "#2F435B", "card_border": "#657A92", "row_border": "#788BA0",
            },
            "Deep Ocean": {
                "page": "#103B47", "card": "#174B58", "row": "#245F6B",
                "input": "#34727D", "button": "#5D858D", "log": "#1F5865",
                "progress": "#286570", "card_border": "#67929A", "row_border": "#739AA1",
            },
            "Storm Purple": {
                "page": "#50506D", "card": "#595A78", "row": "#686985",
                "input": "#777991", "button": "#9C9DB4", "log": "#62637F",
                "progress": "#6D6E89", "card_border": "#A1A2B8", "row_border": "#B0B1C3",
            },
            "Burgundy Ink": {
                "page": "#341A24", "card": "#43232F", "row": "#563542",
                "input": "#684652", "button": "#91707A", "log": "#4A2B37",
                "progress": "#573744", "card_border": "#896675", "row_border": "#9A7885",
            },
            "Obsidian Blue": {
                "page": "#18212D", "card": "#223044", "row": "#30445B",
                "input": "#405873", "button": "#71869E", "log": "#293A50",
                "progress": "#344960", "card_border": "#667B93", "row_border": "#798DA2",
            },
            "Black Cherry": {
                "page": "#351E28", "card": "#3F2530", "row": "#51313E",
                "input": "#624150", "button": "#A56F83", "log": "#482B37",
                "progress": "#563542", "card_border": "#8E6373", "row_border": "#9F7483",
            },
            "Night Plum": {
                "page": "#2C2437", "card": "#352B43", "row": "#493C57",
                "input": "#5B4D69", "button": "#806D8D", "log": "#40344D",
                "progress": "#4D405A", "card_border": "#867792", "row_border": "#90819B",
            },
            "Dark Teal": {
                # Scanner-only: tone the buttons down so they stay readable
                # without jumping forward as pale/minty controls.
                "page": "#173A3A", "card": "#1E4746", "row": "#2A5A57",
                "input": "#376A67", "button": "#709B98", "log": "#244F4D",
                "progress": "#2E5F5C", "card_border": "#6B9794", "row_border": "#7FA7A4",
            },
            "Navy Bloom": {
                "page": "#243F5E", "card": "#315779", "row": "#406789",
                "input": "#557A9A", "button": "#7695B3", "log": "#365E80",
                "progress": "#456E90", "card_border": "#7D9AB5", "row_border": "#8FA9C1",
            },
            "Retro Terminal": {
                "page": "#213B34", "card": "#284940", "row": "#365A4E",
                "input": "#456C5F", "button": "#77A494", "log": "#304F45",
                "progress": "#3A5F53", "card_border": "#73988B", "row_border": "#86A79B",
            },
        }
        depth = scan_depth_palettes.get(self.theme_name)
        if depth:
            scan_page_fill = depth["page"]
            scan_card_fill = depth["card"]
            scan_row_fill = depth["row"]
            scan_input_fill = depth["input"]
            scan_button_fill = depth["button"]
            scan_log_fill = depth["log"]
            scan_progress_fill = depth["progress"]
            scan_card_border = depth["card_border"]
            scan_row_border = depth["row_border"]

        # Only the Dark & Cool Guide gets a little more section colour.  The
        # Library/Main page remains unchanged, and other Guide themes keep their
        # existing lighter treatment.
        if self.theme_name in dark_cool_themes:
            guide_page_fill = _mix_hex(side_fill, "#FFFFFF", 0.86)
            guide_card_fill = _mix_hex(header_fill, "#FFFFFF", 0.72)
            guide_browser_fill = _mix_hex(header_fill, "#FFFFFF", 0.84)
            guide_border = _mix_hex(scrollbar_color, "#FFFFFF", 0.45)
        elif self.theme_name == "Berry Cloud":
            # Keep the Guide noticeably calmer than the pink Scan/Library accents:
            # more cloud/lilac, less overall berry wash.
            guide_page_fill = "#F2EBF1"
            guide_card_fill = "#E9DFE8"
            guide_browser_fill = "#F8F4F7"
            guide_border = "#B990A8"

        if t.get("_custom"):
            guide_page_fill = str(t.get("ui_guide_page_fill", guide_page_fill))
            guide_card_fill = str(t.get("ui_guide_card_fill", guide_card_fill))
            guide_browser_fill = str(t.get("ui_guide_browser_fill", guide_browser_fill))
            guide_border = str(t.get("ui_guide_border", guide_border))
            scan_page_fill = str(t.get("ui_scan_page_fill", scan_page_fill))
            scan_card_fill = str(t.get("ui_scan_card_fill", scan_card_fill))
            scan_row_fill = str(t.get("ui_scan_row_fill", scan_row_fill))
            scan_input_fill = str(t.get("ui_scan_input_fill", scan_input_fill))
            scan_button_fill = str(t.get("ui_scan_button_fill", scan_button_fill))
            scan_log_fill = str(t.get("ui_scan_log_fill", scan_log_fill))
            scan_progress_fill = str(t.get("ui_scan_progress_fill", scan_progress_fill))
        scan_card_text = _button_text_for(scan_card_fill, dark=t["text"], light="#FFFFFF")
        scan_input_text = _button_text_for(scan_input_fill, dark="#171717", light="#F7F7F7")
        # Scanner logs are diagnostic text, so use a neutral high-contrast
        # foreground rather than a decorative theme foreground. This protects
        # readability across dark/strong themes and future imported themes too.
        scan_log_text = "#F2F2F2" if _luma_hex(scan_log_fill) < 160 else "#202020"
        scan_progress_text = _button_text_for(scan_progress_fill, dark="#171717", light="#F7F7F7")
        scan_button_text = _button_text_for(scan_button_fill, dark="#171717", light="#F7F7F7")
        scan_placeholder_text = _button_text_for(scan_input_fill, dark="#404040", light="#E1E1E1")
        if self.theme_name == "Blue Grey Rose":
            # Slightly stronger near-black text on the muted rose Scan controls.
            scan_input_text = "#121212"
            scan_button_text = "#121212"
            scan_placeholder_text = "#303030"
        # Button outlines are independent from structural/card borders. Cherry
        # Blossom Pink uses a very light neutral-grey outline: visible against
        # the pink, but softer than FIX10's near-white edge.
        button_border = "#D8D2D4" if self.theme_name == "Cherry Blossom Pink" else "#C7C7C7"
        strong_button_border = "#D2CCCE" if self.theme_name == "Cherry Blossom Pink" else "#AAAAAA"
        guide_button_border = "#D8D2D4" if self.theme_name == "Cherry Blossom Pink" else "#BDBDBD"
        # Guide jump buttons may use a quieter fill than general action buttons.
        # In Blue Grey Rose the top navigation pink already works well, but the
        # large Guide-page button grid looked too bright against the pale page.
        guide_jump_fill = filter_button
        guide_jump_text = button_text
        if self.theme_name == "Blue Grey Rose":
            guide_jump_fill = "#D7B6C0"
            guide_jump_text = "#20242A"

        # RC40: navigation/Guide may retain the restrained RC39 depth treatment,
        # but ordinary Library buttons stay flat. Their outline now follows the
        # same theme-aware border role used by Scan controls instead of the old
        # generic grey edge.
        nav_button_top = _mix_hex(info_fill, "#FFFFFF", 0.12 if _luma_hex(info_fill) >= 145 else 0.07)
        nav_button_bottom = _mix_hex(info_fill, "#000000", 0.045 if _luma_hex(info_fill) >= 145 else 0.11)
        nav_edge_light = _mix_hex(info_card_border, "#FFFFFF", 0.16 if _luma_hex(info_fill) >= 145 else 0.08)
        nav_edge_dark = _mix_hex(info_card_border, "#000000", 0.08 if _luma_hex(info_fill) >= 145 else 0.16)
        guide_jump_top = _mix_hex(guide_jump_fill, "#FFFFFF", 0.12 if _luma_hex(guide_jump_fill) >= 145 else 0.07)
        guide_jump_bottom = _mix_hex(guide_jump_fill, "#000000", 0.045 if _luma_hex(guide_jump_fill) >= 145 else 0.11)
        guide_edge_button_light = _mix_hex(guide_button_border, "#FFFFFF", 0.18 if _luma_hex(guide_jump_fill) >= 145 else 0.09)
        guide_edge_button_dark = _mix_hex(guide_button_border, "#000000", 0.09 if _luma_hex(guide_jump_fill) >= 145 else 0.17)

        combo_button_divider = "#D8D2D4" if self.theme_name == "Cherry Blossom Pink" else "#AFAFAF"
        scan_shell_badge_fill = info_fill
        scan_shell_badge_text = info_text
        scan_value_badge_fill = filter_button
        scan_value_badge_text = button_text
        if IS_DEV_BUILD and self.theme_name in {
            "Rose Grey", "Cherry Blossom Pink", "Strawberry Milk", "Berry Cloud", "Sakura Tea", "Cafe Rose", "Soft Blush Slate",
            "Blue Coral", "Jelly Blue Pink", "Blue Grey Rose"
        }:
            # The DEV/Edition badge was too pale in the pink/peach and blue+pink
            # families. Use the stronger theme accent only for these DEV badges.
            scan_dev_badge_fill = _mix_hex(t["accent"], "#000000", 0.06)
            scan_shell_badge_fill = scan_dev_badge_fill
            scan_value_badge_fill = scan_dev_badge_fill
            scan_shell_badge_text = _button_text_for(scan_dev_badge_fill, dark="#171717", light="#FFFFFF")
            scan_value_badge_text = scan_shell_badge_text

        # Unified setup-row treatment: the row is the main visual container.
        # Inner controls use softer edges and tiny tonal shifts for restrained
        # depth, avoiding the previous nested/double-border look.
        scan_row_top = _mix_hex(scan_row_fill, "#FFFFFF", 0.18 if _luma_hex(scan_row_fill) >= 145 else 0.09)
        scan_row_bottom = _mix_hex(scan_row_fill, "#000000", 0.07 if _luma_hex(scan_row_fill) >= 145 else 0.16)
        scan_row_edge_light = _mix_hex(scan_row_border, "#FFFFFF", 0.22 if _luma_hex(scan_row_fill) >= 145 else 0.12)
        scan_row_edge_dark = _mix_hex(scan_row_border, "#000000", 0.12 if _luma_hex(scan_row_fill) >= 145 else 0.22)
        scan_control_top = _mix_hex(scan_input_fill, "#FFFFFF", 0.10 if _luma_hex(scan_input_fill) >= 145 else 0.06)
        scan_control_bottom = _mix_hex(scan_input_fill, "#000000", 0.035 if _luma_hex(scan_input_fill) >= 145 else 0.10)
        scan_button_top = _mix_hex(scan_button_fill, "#FFFFFF", 0.13 if _luma_hex(scan_button_fill) >= 145 else 0.07)
        scan_button_bottom = _mix_hex(scan_button_fill, "#000000", 0.05 if _luma_hex(scan_button_fill) >= 145 else 0.12)
        scan_control_border = _mix_hex(scan_row_border, scan_input_fill, 0.72)
        scan_button_border = scan_control_border
        scan_inner_divider = _mix_hex(scan_row_border, scan_row_fill, 0.48)
        scan_focus_border = scrollbar_color
        if self.theme_name in dark_cool_themes:
            # Scanner-only tinted bevels: visible, but no longer bright white.
            scan_row_edge_light = _mix_hex(scan_row_border, "#FFFFFF", 0.12)
            scan_row_edge_dark = _mix_hex(scan_row_border, "#000000", 0.18)
            scan_control_border = _mix_hex(scan_row_border, scan_input_fill, 0.38)
            scan_button_border = scan_control_border
            scan_inner_divider = _mix_hex(scan_row_border, scan_row_fill, 0.30)
        elif self.theme_name == "Aqua Depth":
            scan_row_edge_light = _mix_hex(scan_row_border, "#FFFFFF", 0.18)
            scan_row_edge_dark = _mix_hex(scan_row_border, scrollbar_color, 0.16)
            scan_control_border = _mix_hex(scan_row_border, scan_input_fill, 0.46)
            scan_button_border = scan_control_border
            scan_inner_divider = _mix_hex(scan_row_border, scan_row_fill, 0.36)
        elif self.theme_name == "Carbon Blue":
            # Carbon Blue needed shape rather than extra brightness: strengthen
            # the button bevel and its edge while leaving the input fields calm.
            scan_button_top = _mix_hex(scan_button_fill, "#FFFFFF", 0.20)
            scan_button_bottom = _mix_hex(scan_button_fill, "#000000", 0.11)
            scan_button_border = "#70899B"
        elif self.theme_name in white_border_themes:
            # Keep the established crisp light outlines in the other theme
            # families that still benefit from them.
            scan_row_edge_light = "#FFFFFF"
            scan_row_edge_dark = "#FFFFFF"
            scan_control_border = "#FFFFFF"
            scan_button_border = scan_control_border
            scan_inner_divider = "#FFFFFF"
        elif self.theme_name == "Cherry Blossom Pink":
            # Use a rose-tinted hierarchy rather than one repeated pale outline.
            scan_row_edge_light = _mix_hex(scan_row_border, "#FFFFFF", 0.20)
            scan_row_edge_dark = _mix_hex(scan_row_border, "#8F5365", 0.18)
            scan_control_border = "#D7A8B5"
            scan_button_border = scan_control_border
            scan_inner_divider = "#D2A0AE"

        # Accessibility/readability roles shared with PUBLIC. Normal explanatory
        # text uses near-black on light surfaces and near-white on dark surfaces;
        # grey is kept for disabled controls rather than active information.
        active_dark_text = "#171717"
        active_light_text = "#F7F7F7"
        scan_page_text = _button_text_for(scan_page_fill, dark=active_dark_text, light=active_light_text)
        scan_hint_text = _button_text_for(side_fill, dark=active_dark_text, light=active_light_text)
        guide_active_text = _button_text_for(guide_browser_fill, dark=active_dark_text, light=active_light_text)
        scan_status_text = surface_text
        scan_row_text = scan_card_text
        guide_page_text = guide_active_text
        guide_card_text = guide_active_text
        thumbnail_text = active_dark_text
        if t.get("_custom"):
            scan_hint_text = scan_card_text
            scan_status_text = scan_card_text
            scan_row_text = _button_text_for(scan_row_fill, dark=active_dark_text, light=active_light_text)
            guide_page_text = _button_text_for(guide_page_fill, dark=active_dark_text, light=active_light_text)
            guide_card_text = _button_text_for(guide_card_fill, dark=active_dark_text, light=active_light_text)
            thumbnail_text = _button_text_for(preview_mat_fill, dark=active_dark_text, light=active_light_text)
        guide_edge_light = _mix_hex(guide_border, "#FFFFFF", 0.18 if _luma_hex(guide_card_fill) >= 145 else 0.10)
        guide_edge_dark = _mix_hex(guide_border, "#000000", 0.09 if _luma_hex(guide_card_fill) >= 145 else 0.18)
        if self.theme_name in white_border_themes and self.theme_name not in dark_cool_themes:
            guide_edge_light = "#FFFFFF"
            guide_edge_dark = "#FFFFFF"
        # Active Edit-page text follows the actual side-panel surface rather
        # than the top-bar surface. This matters for themes such as Shades of
        # Grey where those two surfaces have intentionally different brightness.
        edit_page_text = _button_text_for(side_fill, dark=active_dark_text, light=active_light_text)
        edit_help_text = edit_page_text
        # RC65: these built-in dark themes use a light/white reading surface in
        # the Edit tab.  Their theme-derived Edit help text was near-white,
        # which made Manual fields/help copy almost disappear.  Keep the rest
        # of each theme unchanged and force only the Edit reading copy to black.
        black_edit_reading_themes = {
            "Navy Bloom", "Midnight Steel", "Deep Ocean", "Storm Purple",
            "Retro Terminal", "Black Cherry", "Dark Teal", "Night Plum",
        }
        edit_reading_text = "#000000" if self.theme_name in black_edit_reading_themes else edit_help_text
        # RC67: the lower Edit save/status line and CASP observations use the
        # same light reading surfaces, so keep their foreground readable too.
        edit_status_text = "#000000" if self.theme_name in black_edit_reading_themes else edit_page_text
        casp_reading_text = "#000000" if self.theme_name in black_edit_reading_themes else edit_help_text
        light_surface_text = active_dark_text

        # Shared palette for top-level dialogs/message boxes and the expanded
        # Scanner Log. This is deliberately based on the Scan surfaces because
        # those already have the clearest theme-aware depth and borders.
        self._dialog_palette = {
            "background": scan_card_fill,
            "panel": scan_row_fill,
            "input": scan_input_fill,
            "text": scan_card_text,
            "input_text": scan_input_text,
            "button": scan_button_fill,
            "button_text": scan_button_text,
            "border": scan_card_border,
            "control_border": scan_button_border,
            "accent": scrollbar_color,
            "log": scan_log_fill,
            "log_text": scan_log_text,
        }

        self.setStyleSheet(f"""
            QMainWindow, QWidget {{
                background: #FFFFFF;
                color: {t['text']};
                font-family: 'Segoe UI';
                font-size: 10pt;
            }}
            QMenuBar {{ background: #FFFFFF; border-bottom: 1px solid #E4E4E4; }}
            QMenuBar::item {{ padding: 5px 8px; background: transparent; }}
            QMenuBar::item:selected {{ background: {t['panel_alt']}; }}
            QMenu {{ background: #FFFFFF; border: 1px solid {t['thumb_border']}; }}
            QMenu::item:selected {{ background: {t['panel_alt']}; }}

            QPushButton#AppPageButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {nav_button_top}, stop:1 {nav_button_bottom});
                color: {info_text};
                border: 1px solid {strong_button_border if self.theme_name == "Cherry Blossom Pink" else info_card_border};
                border-top-color: {nav_edge_light};
                border-left-color: {nav_edge_light};
                border-bottom-color: {nav_edge_dark};
                border-right-color: {nav_edge_dark};
                border-radius: 7px;
                padding: 5px 11px;
                min-height: 22px;
                font-weight: 650;
            }}
            QPushButton#AppPageButton:hover {{ border-color: {scrollbar_color}; }}
            QPushButton#AppPageButton:pressed {{
                background: {nav_button_bottom};
                border-top-color: {nav_edge_dark}; border-left-color: {nav_edge_dark};
                border-bottom-color: {nav_edge_light}; border-right-color: {nav_edge_light};
            }}
            QPushButton#AppPageButton:checked {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {_mix_hex(filter_button, '#FFFFFF', 0.06)}, stop:1 {_mix_hex(filter_button, '#000000', 0.07)});
                color: {button_text};
                border: 1px solid {strong_button_border if self.theme_name == "Cherry Blossom Pink" else scrollbar_color};
            }}
            QFrame#LibraryTopRow {{
                background: {top_fill};
                border: none;
                border-radius: 5px;
            }}
            QFrame#LibraryTopRow QLabel {{ background: transparent; color: {surface_text}; }}
            QFrame#LibraryFiltersRow {{ background: transparent; border: none; }}
            QWidget#LibraryActionSlot {{ background: transparent; border: none; }}
            QWidget#LibraryActionGroup {{ background: transparent; border: none; }}
            QLabel#ScanShellBadge {{
                background: {scan_shell_badge_fill};
                color: {scan_shell_badge_text};
                border: 1px solid {info_card_border};
                border-radius: 6px;
                padding: 3px 8px;
                font-weight: 600;
            }}
            QScrollArea#ScanPageScroll {{ background: {scan_page_fill}; border: none; }}
            QScrollArea#ScanPageScroll > QWidget > QWidget {{ background: {scan_page_fill}; }}
            QFrame#ScanPage {{ background: {scan_page_fill}; border: none; }}
            QLabel#ScanTitle {{ font-size: 16pt; font-weight: 750; color: {scan_page_text}; background: transparent; }}
            QLabel#ScanSubtitle {{ color: {scan_page_text}; background: transparent; }}
            QFrame#ScanCard {{
                background: {scan_card_fill};
                border: {scan_border_width}px solid {scan_card_border};
                border-radius: 8px;
            }}
            QFrame#ScanCard QLabel {{ background: transparent; color: {scan_card_text}; }}
            QSplitter#ScanVerticalSplitter::handle:vertical {{
                background: {scan_row_border};
                margin: 2px 90px;
                border-radius: 3px;
            }}
            QSplitter#ScanVerticalSplitter::handle:vertical:hover {{ background: {scrollbar_color}; }}
            QFrame#ScanCard QWidget#ScanSetupRow QLabel {{ color: {scan_row_text}; }}
            QWidget#ScanSetupRow {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {scan_row_top}, stop:1 {scan_row_bottom});
                border: {scan_border_width}px solid {scan_row_border};
                border-top-color: {scan_row_edge_light}; border-left-color: {scan_row_edge_light};
                border-bottom-color: {scan_row_edge_dark}; border-right-color: {scan_row_edge_dark};
                border-radius: 6px;
            }}
            QLabel#ScanSectionTitle {{ font-size: 11pt; font-weight: 700; color: {surface_text}; }}
            QLabel#ScanStatus, QLabel#ScanProgressLabel, QLabel#ScanCurrentFile {{ color: {scan_status_text}; }}
            QLabel#ScanHint {{ color: {scan_hint_text}; font-size: 9pt; }}
            QFrame#ScanCard QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {scan_button_top}, stop:1 {scan_button_bottom});
                color: {scan_button_text};
                border: 1px solid {scan_button_border}; border-radius: 5px;
                min-height: 24px; padding: 4px 10px;
            }}
            QFrame#ScanCard QPushButton:hover {{ border-color: {scan_focus_border}; }}
            QFrame#ScanCard QPushButton:pressed {{
                background: {scan_button_bottom};
                border-top-color: {scan_row_edge_dark}; border-left-color: {scan_row_edge_dark};
                border-bottom-color: {scan_row_edge_light}; border-right-color: {scan_row_edge_light};
            }}
            QLabel#ScanValueBadge {{
                background: {scan_value_badge_fill}; color: {scan_value_badge_text};
                border: 1px solid {scrollbar_color}; border-radius: 5px; padding: 4px 8px; font-weight: 700;
            }}
            QLineEdit#ScanPathField {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {scan_control_top}, stop:1 {scan_control_bottom});
                color: {scan_input_text}; placeholder-text-color: {scan_placeholder_text};
                border: 1px solid {scan_control_border}; border-radius: 5px;
                min-height: 28px; padding: 3px 8px;
            }}
            QLineEdit#ScanPathField:focus {{ border-color: {scan_focus_border}; }}
            QComboBox#ScanModeCombo {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {scan_control_top}, stop:1 {scan_control_bottom});
                color: {scan_input_text};
                border: 1px solid {scan_control_border}; border-radius: 5px;
                padding: 3px 8px;
            }}
            QComboBox#ScanModeCombo:focus {{ border-color: {scan_focus_border}; }}
            QComboBox#ScanModeCombo::drop-down {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {scan_button_top}, stop:1 {scan_button_bottom});
                border: none; border-left: 1px solid {scan_inner_divider}; min-width: 32px;
            }}
            QComboBox#ScanModeCombo QAbstractItemView {{
                background: {scan_input_fill}; color: {scan_input_text};
                selection-background-color: {header_fill}; selection-color: {header_text};
            }}
            QPlainTextEdit#ScanLog {{
                background: {scan_log_fill}; color: {scan_log_text}; border: 1px solid {scan_control_border};
                border-radius: 6px; padding: 6px; font-family: Consolas; font-size: 10pt;
            }}
            QTabWidget#ScanOutputTabs::pane {{ background: {scan_log_fill}; border: {scan_border_width}px solid {scan_row_border}; }}
            QTabWidget#ScanOutputTabs QTabBar::tab {{
                background: {scan_button_fill}; color: {scan_button_text};
                border: {scan_border_width}px solid {scan_row_border}; border-bottom: none; padding: 5px 9px;
            }}
            QTabWidget#ScanOutputTabs QTabBar::tab:selected {{ background: {scan_input_fill}; color: {scan_input_text}; }}
            QWidget#ScanOutputPage {{ background: {scan_log_fill}; }}
            QTableWidget#RoutingPreviewTable {{
                background: {scan_log_fill}; color: {scan_log_text}; alternate-background-color: {scan_row_fill};
                border: {scan_border_width}px solid {scan_row_border}; gridline-color: {scan_row_border};
            }}
            QProgressBar {{
                background: {scan_progress_fill}; color: {scan_progress_text}; border: 1px solid {scan_control_border};
                border-radius: 7px; text-align: center; min-height: 18px;
            }}
            QProgressBar::chunk {{ background: {scrollbar_color}; border-radius: 6px; }}
            QPushButton#ScanStartButton {{
                background: {filter_button}; color: {button_text}; border: 1px solid {scrollbar_color};
                border-radius: 7px; padding: 7px 18px; font-weight: 700;
            }}

            QMainWindow#ThemedUtilityWindow, QWidget#ThemedUtilityPanel {{
                background: {scan_card_fill}; color: {scan_card_text};
            }}
            QMainWindow#ThemedUtilityWindow QLabel {{ background: transparent; color: {scan_card_text}; }}
            QMainWindow#ThemedUtilityWindow QPushButton {{
                background: {scan_button_fill}; color: {scan_button_text};
                border: 1px solid {scan_button_border}; border-radius: 6px;
                min-height: 26px; padding: 4px 11px;
            }}
            QMainWindow#ThemedUtilityWindow QPushButton:hover {{ border-color: {scrollbar_color}; }}
            QPlainTextEdit#ExpandedScannerLog {{
                /* RC68: expanded logs stay neutral/readable regardless of active theme. */
                background: #E8EAED; color: #000000;
                border: 1px solid {scan_control_border}; border-radius: 7px;
                padding: 8px; font-family: Consolas; font-size: 10pt;
                selection-background-color: #BFC5CC; selection-color: #000000;
            }}

            QWidget#GuidePage {{ background: {guide_page_fill}; }}
            QLabel#GuideTitle {{ font-size: 16pt; font-weight: 750; color: {guide_page_text}; background: transparent; }}
            QLabel#GuideSubtitle {{ color: {guide_page_text}; background: transparent; }}
            QLabel#GuideGroupTitle {{ color: {guide_page_text}; background: transparent; font-size: 11pt; font-weight: 700; margin-top: 3px; }}
            QFrame#GuideJumpCard QLabel {{ color: {guide_card_text}; background: transparent; }}
            QFrame#GuideJumpCard {{
                background: {guide_card_fill}; border: {guide_border_width}px solid {guide_border};
                border-top-color: {guide_edge_light}; border-left-color: {guide_edge_light};
                border-bottom-color: {guide_edge_dark}; border-right-color: {guide_edge_dark};
                border-radius: 8px;
            }}
            QPushButton#GuideJumpButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {guide_jump_top}, stop:1 {guide_jump_bottom});
                color: {guide_jump_text}; border: 1px solid {guide_button_border};
                border-top-color: {guide_edge_button_light}; border-left-color: {guide_edge_button_light};
                border-bottom-color: {guide_edge_button_dark}; border-right-color: {guide_edge_button_dark};
                border-radius: 6px; padding: 5px 9px; min-height: 22px;
            }}
            QPushButton#GuideJumpButton:hover {{ border-color: {scrollbar_color}; }}
            QPushButton#GuideJumpButton:pressed {{
                background: {guide_jump_bottom};
                border-top-color: {guide_edge_button_dark}; border-left-color: {guide_edge_button_dark};
                border-bottom-color: {guide_edge_button_light}; border-right-color: {guide_edge_button_light};
            }}
            QTextBrowser#GuideBrowser {{
                background: {guide_browser_fill}; color: {guide_active_text}; border: {guide_border_width}px solid {guide_border};
                border-top-color: {guide_edge_light}; border-left-color: {guide_edge_light};
                border-bottom-color: {guide_edge_dark}; border-right-color: {guide_edge_dark};
                border-radius: 7px; padding: 10px;
            }}

            QFrame#FilterCard {{
                background: {filter_panel_fill};
                border: 1px solid {scrollbar_color};
                border-radius: 9px;
            }}
            QFrame#FilterCard QLabel {{ background: transparent; color: {_button_text_for(filter_panel_fill, dark=t["text"], light="#FFFFFF")}; }}
            QFrame#ActiveFilterBar {{ background: transparent; border: none; }}
            QScrollArea#ActiveFilterScroll {{ background: transparent; border: none; }}
            QWidget#ActiveFilterChipHost {{ background: transparent; }}
            QPushButton#FilterChip {{
                background: {filter_button}; color: {button_text}; border: 1px solid {strong_button_border};
                border-radius: 9px; padding: 3px 8px; min-height: 18px;
            }}
            QPushButton#FilterChip:hover {{ border: 2px solid {strong_button_border}; padding: 2px 7px; }}
            QWidget#OverviewPage {{ background: {side_fill}; }}
            QScrollArea#OverviewScroll {{ background: {side_fill}; border: none; }}
            QScrollArea#OverviewScroll > QWidget > QWidget#OverviewPage {{ background: {side_fill}; }}
            QLabel#StatsStrong {{ font-weight: 700; font-size: 11pt; }}
            QLabel#UnsavedLabel {{ font-weight: 650; }}
            QLabel#MDBPathCaption {{ font-weight: 700; background: transparent; color: {surface_text}; }}
            QLabel#MDBPathLabel {{ background: #FFFFFF; color: {t['text']}; border: 1px solid #C4C4C4; border-radius: 6px; padding: 3px 7px; }}
            QPushButton#OpenMDBButton {{ background: {filter_button}; color: {button_text}; border: 1px solid {strong_button_border}; border-radius: 8px; padding: 6px 12px; }}
            QWidget#EditPage {{ background: {side_fill}; }}
            QScrollArea#EditScroll {{ background: {side_fill}; border: none; }}
            QScrollArea#EditScroll > QWidget > QWidget {{ background: {side_fill}; }}
            QWidget#EditScrollContent {{ background: {side_fill}; }}
            QLabel#EditTitle {{ color: {edit_page_text}; font-size: 11pt; font-weight: 700; }}
            QLabel#EditHelp {{ color: {edit_help_text}; }}
            QWidget#EditScrollContent QLabel#EditTitle {{ color: {edit_reading_text}; }}
            QWidget#EditScrollContent QLabel#EditHelp {{ color: {edit_reading_text}; }}
            QFrame#EditCard {{ background: #FFFFFF; border: 1px solid #CFCFCF; border-radius: 8px; }}
            QFrame#EditCard QLabel {{ background: transparent; color: {active_dark_text}; }}
            QLabel#DirtyLabel {{ color: {edit_status_text}; font-weight: 600; }}
            QWidget#CASPObservationsPage QLabel#EditHelp {{ color: {casp_reading_text}; }}
            QTextBrowser#CASPObservationsBrowser {{ color: {casp_reading_text}; }}
            QPushButton#SaveButton {{ background: {filter_button}; color: {button_text}; border: 1px solid {strong_button_border}; border-radius: 8px; padding: 7px 13px; }}

            QLineEdit, QComboBox {{
                background: #FFFFFF;
                color: {active_dark_text};
                placeholder-text-color: #4A4A4A;
                border: 1px solid #C4C4C4;
                border-radius: 10px;
                padding: 4px 7px;
                min-height: 18px;
            }}
            QLineEdit:focus, QComboBox:focus {{ border: 2px solid {scrollbar_color}; }}
            QPlainTextEdit#CCMagicProfilesEdit {{
                background: #F1F1F1; color: {active_dark_text}; placeholder-text-color: #5C5C5C; border: 1px solid #C4C4C4;
                border-radius: 8px; padding: 4px 6px;
            }}
            QPlainTextEdit#NotesEdit {{
                background: #FFFFFF; color: {active_dark_text}; placeholder-text-color: #4A4A4A; border: 1px solid #C4C4C4;
                border-radius: 8px; padding: 4px 6px;
            }}
            QPlainTextEdit#CCMagicProfilesEdit:focus, QPlainTextEdit#NotesEdit:focus {{ border: 2px solid {scrollbar_color}; }}
            QFrame#FilterCard QLineEdit {{ background: #FFFFFF; color: {active_dark_text}; }}
            QFrame#FilterCard QComboBox {{
                background: #FFFFFF;
                color: {active_dark_text};
                border: 1px solid #BEBEBE;
                border-radius: 10px;
                padding: 4px 42px 4px 7px;
                min-height: 18px;
            }}
            QFrame#FilterCard QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 40px;
                background: {filter_button};
                border-left: 1px solid {combo_button_divider};
                border-top-right-radius: 9px;
                border-bottom-right-radius: 9px;
            }}
            QFrame#FilterCard QComboBox::down-arrow {{ image: none; width: 0px; height: 0px; }}
            QFrame#FilterCard QComboBox QAbstractItemView {{
                background: #FFFFFF;
                color: {active_dark_text};
                selection-background-color: {header_fill};
                selection-color: {header_text};
            }}
            QPushButton {{
                background: #F6F6F6;
                color: {t['text']};
                border: 1px solid {button_border};
                border-radius: 5px;
                padding: 5px 12px;
            }}
            QFrame#FilterCard QPushButton {{
                background: {filter_button};
                color: {button_text};
                border: 1px solid {scan_button_border};
                border-radius: 8px;
                padding: 5px 12px;
            }}
            QFrame#FilterCard QPushButton:hover {{ border-color: {scan_focus_border}; }}
            QFrame#FilterCard QPushButton:pressed {{ background: {_mix_hex(filter_button, '#000000', 0.045)}; }}

            QFrame#PinnedBar {{ background: transparent; border: none; }}
            QLabel#PinnedLabel {{ font-weight: 700; color: {t['text']}; background: transparent; }}
            QScrollArea#PinnedScroll {{ background: transparent; border: none; }}
            QScrollArea#PinnedScroll > QWidget > QWidget {{ background: transparent; }}
            QFrame#PinChip {{ background: transparent; border: none; }}
            QPushButton#PinnedTab {{
                background: {header_fill}; color: {header_text}; border: 1px solid {button_border};
                border-top-left-radius: 6px; border-bottom-left-radius: 6px;
                border-top-right-radius: 0px; border-bottom-right-radius: 0px;
                padding: 4px 8px; font-weight: 600;
            }}
            QPushButton#PinnedTab:checked {{ background: {filter_button}; color: {button_text}; border-color: {strong_button_border if self.theme_name == "Cherry Blossom Pink" else scrollbar_color}; }}
            QPushButton#PinnedRemove {{
                background: #F7F7F7; color: {t['text']}; border: 1px solid {button_border}; border-left: none;
                border-top-right-radius: 6px; border-bottom-right-radius: 6px;
                border-top-left-radius: 0px; border-bottom-left-radius: 0px; padding: 4px 2px; font-weight: 700;
            }}
            QPushButton#PinnedRemove:hover {{ background: {info_fill}; }}

            QPushButton#ViewToggle {{
                background: #FFFFFF;
                color: {t['text']};
                border: 1px solid {guide_button_border};
                border-radius: 7px;
                padding: 4px 10px;
            }}
            QPushButton#ViewToggle:checked {{
                background: {filter_button};
                color: {button_text};
                border: 1px solid {strong_button_border if self.theme_name == "Cherry Blossom Pink" else scrollbar_color};
            }}
            QLabel#ViewLabel {{ background: transparent; color: {surface_text}; }}

            QFrame#MainPanel {{ background: {main_fill}; border: 1px solid #C9C9C9; border-radius: 5px; }}
            QScrollArea#GalleryQuickScroll {{ background: transparent; border: none; }}
            QFrame#GalleryQuickBar {{
                background: {header_fill};
                border: 1px solid #C9C9C9;
                border-radius: 7px;
                margin: 4px 4px 0px 4px;
            }}
            QLabel#GalleryQuickLabel {{ background: transparent; color: {header_text}; font-weight: 650; }}
            QComboBox#GalleryQuickCombo {{
                background: #FFFFFF;
                color: {active_dark_text};
                border: 1px solid #C2C2C2;
                border-radius: 7px;
                padding: 4px 34px 4px 8px;
                min-height: 20px;
            }}
            QComboBox#GalleryQuickCombo::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 31px;
                background: {filter_button};
                border-left: 1px solid #B7B7B7;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
            }}
            QComboBox#GalleryQuickCombo::down-arrow {{ image: none; width: 0px; height: 0px; }}
            QPushButton#GalleryQuickButton, QPushButton#GalleryQuickClear {{
                background: #FFFFFF;
                color: {t['text']};
                border: 1px solid {button_border};
                border-radius: 9px;
                padding: 4px 9px;
            }}
            QPushButton#GalleryQuickButton:checked {{
                background: {filter_button};
                color: {button_text};
                border: 1px solid {strong_button_border if self.theme_name == "Cherry Blossom Pink" else scrollbar_color};
            }}
            QPushButton#GalleryQuickButton:hover, QPushButton#GalleryQuickClear:hover {{ border-color: {scrollbar_color}; }}
            QStackedWidget#LibraryStack {{ background: {main_fill}; border: none; }}
            QListView#GalleryView {{
                background: {main_fill};
                border: none;
                outline: none;
                padding: 4px;
            }}
            QListView#GalleryView::item {{ background: transparent; border: none; }}
            QTableView {{
                background: {main_fill};
                color: {t['text']};
                border: none;
                selection-background-color: transparent;
                selection-color: {t['text']};
                outline: none;
            }}
            QTableView::item {{ padding: 4px 6px; border: none; }}
            QTableView::item:hover {{ background: #F7F7F7; }}
            QHeaderView::section {{
                background: {table_header_fill};
                color: {table_header_text};
                border-top: 1px solid #B7B7B7;
                border-right: 1px solid #B7B7B7;
                border-bottom: 2px solid #AAAAAA;
                padding: 7px 8px;
                font-weight: 650;
            }}
            QHeaderView::section:first {{ border-left: 1px solid #B7B7B7; }}

            QScrollBar:vertical {{ background: #F3F3F3; width: 22px; margin: 1px; border: none; border-radius: 7px; }}
            QScrollBar::handle:vertical {{ background: {scrollbar_color}; min-height: 42px; border: none; border-radius: 8px; margin: 3px 4px; }}
            QScrollBar::handle:vertical:hover {{ background: {scrollbar_hover}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; background: transparent; border: none; }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
            QScrollBar:horizontal {{ background: #F3F3F3; height: 22px; margin: 1px; border: none; border-radius: 10px; }}
            QScrollBar::handle:horizontal {{ background: {scrollbar_color}; min-width: 42px; border: none; border-radius: 8px; margin: 4px 3px; }}
            QScrollBar::handle:horizontal:hover {{ background: {scrollbar_hover}; }}
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0px; background: transparent; border: none; }}
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{ background: transparent; }}

            QFrame#DetailPanel {{ background: {side_fill}; border: 1px solid #CFCFCF; border-radius: 6px; }}
            QScrollArea#DetailTitleScroll {{ background: transparent; border: none; }}
            QScrollArea#DetailTitleScroll > QWidget > QWidget#DetailTitleContent {{ background: transparent; }}
            QScrollArea#DetailTitleScroll QScrollBar:vertical {{ width: 16px; background: transparent; border: none; }}
            QScrollArea#DetailTitleScroll QScrollBar::handle:vertical {{ background: {scrollbar_color}; min-height: 28px; border-radius: 6px; margin: 2px; }}
            QLabel#DetailTitle {{ background: transparent; color: {detail_title_text}; font-size: 11pt; font-weight: 700; padding: 1px; }}
            QPushButton#PinButton {{ background: {badge_fill}; color: {badge_text}; border: 1px solid {button_border}; border-radius: 5px; padding: 3px 7px; font-weight: 650; }}
            QLabel#Badge {{ background: {badge_fill}; color: {badge_text}; border: 1px solid #C7C7C7; border-radius: 5px; padding: 2px 5px; }}
            QTabWidget::pane {{ border: 1px solid #CFCFCF; background: {side_fill}; }}
            QTabBar::tab {{ background: {tab_fill}; color: {tab_text}; border: 1px solid #C7C7C7; border-bottom: none; padding: 6px 10px; margin-right: 2px; }}
            QTabBar::tab:selected {{ background: {tab_selected_fill}; color: {tab_selected_text}; border-top: 1px solid #C7C7C7; }}
            QTabWidget#DetailTabs::tab-bar {{ left: 4px; }}
            QTabWidget#DetailTabs QTabBar::tab {{ padding: 4px 10px; }}
            QTabWidget#DetailTabs QTabBar::tab:last {{ margin-right: 0px; }}

            QFrame#ThumbnailCard {{ background: transparent; border: none; border-radius: 10px; }}
            QLabel#ThumbnailImage {{
                background: {preview_mat_fill};
                border: 1px solid {preview_mat_border};
                border-radius: 10px;
                color: {thumbnail_text};
            }}
            QFrame#InfoCard {{ background: {detail_info_fill}; border: 1px solid {info_card_border}; border-radius: 8px; }}
            QTextBrowser#InfoBrowser {{ background: {detail_info_fill}; color: {detail_info_text}; border: none; padding: 4px; }}
            QTableWidget {{ background: #FFFFFF; color: {t['text']}; border: 1px solid #C9C9C9; gridline-color: {row_line}; }}
            QDialog {{ background: #FFFFFF; }}
            QScrollArea {{ background: #FFFFFF; border: 1px solid #D6D6D6; }}
            QCheckBox {{ spacing: 6px; }}
            QPushButton#RoutingAckButton {{
                background: #FFFFFF;
                color: {t['text']};
                border: 1px solid {strong_button_border};
                border-radius: 6px;
                padding: 5px 10px;
                text-align: left;
                font-weight: 600;
            }}
            QPushButton#RoutingAckButton:hover {{
                border: 1px solid {scrollbar_color};
            }}
            QPushButton#RoutingAckButton:checked {{
                background: {filter_button};
                color: {button_text};
                border: 2px solid {scrollbar_color};
                font-weight: 700;
            }}

            /* Disabled editors use muted grey intentionally. Active/read-only
               information above remains high-contrast and readable. */
            QLineEdit:disabled, QComboBox:disabled, QPlainTextEdit:disabled {{
                background: #E3E3E3;
                color: #666666;
                placeholder-text-color: #777777;
                border: 1px solid #B9B9B9;
            }}

            /* Disabled buttons must look unmistakably unavailable, including
               buttons with theme/ID-specific enabled-state styling above. */
            QPushButton:disabled,
            QPushButton#AppPageButton:disabled,
            QPushButton#ScanStartButton:disabled,
            QPushButton#OpenMDBButton:disabled,
            QPushButton#SaveButton:disabled,
            QPushButton#PinButton:disabled,
            QPushButton#ViewToggle:disabled,
            QPushButton#GalleryQuickButton:disabled,
            QPushButton#GalleryQuickClear:disabled,
            QPushButton#PinnedTab:disabled,
            QPushButton#PinnedRemove:disabled,
            QPushButton#RoutingAckButton:disabled,
            QFrame#FilterCard QPushButton:disabled,
            QFrame#ScanCard QPushButton:disabled {{
                background: #DADADA;
                color: #666666;
                border: 1px solid #B9B9B9;
            }}
            QStatusBar {{ background: #F4F4F4; color: {light_surface_text}; border-top: 1px solid #DDDDDD; }}
        """)

        for box in self.filter_boxes.values():
            if isinstance(box, ArrowComboBox):
                box.set_arrow_color(arrow_color)
        if hasattr(self, "gallery_type_filter") and isinstance(self.gallery_type_filter, ArrowComboBox):
            self.gallery_type_filter.set_arrow_color(arrow_color)
        if hasattr(self, "gallery_quick_scroll") and hasattr(self, "gallery_quick_bar"):
            # Re-evaluate after the theme stylesheet is active so the transparent
            # overflow viewport is exactly the natural FIX46 quick-bar height.
            self.gallery_quick_scroll.setFixedHeight(self.gallery_quick_bar.sizeHint().height())

        delegate = self.table.itemDelegate()
        theme_for_rows = dict(self.theme)
        theme_for_rows["row_main"] = main_fill if t.get("_custom") else "#FFFFFF"
        theme_for_rows["line_alt"] = row_alt
        theme_for_rows["line"] = row_line
        theme_for_rows["dirty_marker"] = "#E09A32"
        if not isinstance(delegate, RowOutlineDelegate):
            delegate = RowOutlineDelegate(self.table, theme_for_rows, self._table_editor_options, self.table)
            self.table.setItemDelegate(delegate)
        else:
            delegate.set_theme(theme_for_rows)
            delegate.set_options_provider(self._table_editor_options)
        self.table.set_theme(theme_for_rows)
        self.table.viewport().update()
        if hasattr(self, "gallery_delegate"):
            self.gallery_delegate.set_theme({
                "card": "#FFFFFF",
                "border": "#CFCFCF",
                "image": "#F7F7F7",
                "text": t["text"],
                "muted": t["muted"],
                "selection": scrollbar_color,
                "badge": info_fill,
                "badge_text": info_text,
                "dirty": "#E09A32",
            })

    def _set_library_view(self, mode: str, persist: bool = True):
        mode = "gallery" if mode == "gallery" else "table"
        self.view_mode = mode
        if hasattr(self, "library_stack"):
            self.library_stack.setCurrentIndex(1 if mode == "gallery" else 0)
        if hasattr(self, "gallery_quick_scroll"):
            self.gallery_quick_scroll.setVisible(mode == "gallery")
        elif hasattr(self, "gallery_quick_bar"):
            # Compatibility fallback for older DEV layouts.
            self.gallery_quick_bar.setVisible(mode == "gallery")
        if hasattr(self, "column_fold_bar"):
            self.column_fold_bar.setVisible(mode == "table")
        if hasattr(self, "gallery_status_buttons"):
            self._apply_gallery_quick_filters()
        if hasattr(self, "table_view_btn"):
            self.table_view_btn.blockSignals(True)
            self.gallery_view_btn.blockSignals(True)
            self.table_view_btn.setChecked(mode == "table")
            self.gallery_view_btn.setChecked(mode == "gallery")
            self.table_view_btn.blockSignals(False)
            self.gallery_view_btn.blockSignals(False)
        if persist:
            self.settings.setValue("library_view", mode)
            self.settings.sync()
        current = self.table.selectionModel().currentIndex() if hasattr(self, "table") else QModelIndex()
        if current.isValid():
            if mode == "gallery" and hasattr(self, "gallery"):
                self.gallery.scrollTo(current, QAbstractItemView.PositionAtCenter)
            elif mode == "table":
                self.table.scrollTo(current, QTableView.PositionAtCenter)

    def _gallery_status_filter_clicked(self, status_name: str, checked: bool):
        # Only one status shortcut at a time; the other Gallery filters can stack.
        if checked:
            for name, button in self.gallery_status_buttons.items():
                if name != status_name:
                    button.blockSignals(True)
                    button.setChecked(False)
                    button.blockSignals(False)
        self._apply_gallery_quick_filters()

    def _gallery_pinned_source_rows(self):
        rows = set()
        for pin in self.pinned_packages:
            row = self._source_row_for_pin(pin.get("id", ""))
            if row is not None:
                rows.add(row)
        return rows

    def _apply_gallery_quick_filters(self, *_):
        if not hasattr(self, "gallery_status_buttons"):
            return
        # Keep the shortcuts truly Gallery-only. Switching back to Table restores
        # the ordinary global filters without losing the checked Gallery chips.
        if getattr(self, "view_mode", "table") != "gallery":
            self.proxy.set_gallery_quick_filters()
            self.shown_label.setText(f"Shown: {self.proxy.rowCount():,}")
            return
        status = next((name for name, button in self.gallery_status_buttons.items() if button.isChecked()), "")
        cc_type = self.gallery_type_filter.currentText() if hasattr(self, "gallery_type_filter") else "All"
        self.proxy.set_gallery_quick_filters(
            status=status,
            cc_type="" if cc_type == "All" else cc_type,
            thumbnails_only=self.gallery_thumb_filter_btn.isChecked(),
            pinned_only=self.gallery_pinned_filter_btn.isChecked(),
            dirty_only=self.gallery_unsaved_filter_btn.isChecked(),
            pinned_rows=self._gallery_pinned_source_rows(),
        )
        self.shown_label.setText(f"Shown: {self.proxy.rowCount():,}")

    def _clear_gallery_quick_filters(self):
        for button in self.gallery_status_buttons.values():
            button.blockSignals(True)
            button.setChecked(False)
            button.blockSignals(False)
        for button in (self.gallery_thumb_filter_btn, self.gallery_pinned_filter_btn, self.gallery_unsaved_filter_btn):
            button.blockSignals(True)
            button.setChecked(False)
            button.blockSignals(False)
        if hasattr(self, "gallery_type_filter"):
            self.gallery_type_filter.blockSignals(True)
            self.gallery_type_filter.setCurrentText("All")
            self.gallery_type_filter.blockSignals(False)
        self._apply_gallery_quick_filters()

    def _gallery_clicked(self, index: QModelIndex):
        if not index.isValid():
            return
        selection = self.table.selectionModel()
        selection.setCurrentIndex(index, QItemSelectionModel.ClearAndSelect | QItemSelectionModel.Rows)
        self._show_selected()

    def _record_pin_id(self, rec: Dict[str, str]) -> str:
        package_hash = (rec.get("Package Hash", "") or "").strip().lower()
        if package_hash:
            return "hash:" + package_hash
        source_path = (rec.get("Original Source Path", "") or "").strip()
        if source_path:
            return "path:" + source_path.casefold()
        return "package:" + (rec.get("Package", "") or "").strip().casefold()

    def _pin_settings_key(self) -> Optional[str]:
        if not self.mdb:
            return None
        normalized = str(self.mdb.path.resolve()).casefold().encode("utf-8", errors="replace")
        digest = hashlib.sha256(normalized).hexdigest()
        return f"pins/{digest}"

    def _load_pins_for_current_mdb(self):
        self.pinned_packages = []
        key = self._pin_settings_key()
        if not key:
            return
        raw = self.settings.value(key, "")
        if not raw:
            return
        try:
            data = json.loads(str(raw))
        except Exception:
            return
        if not isinstance(data, list):
            return
        seen = set()
        for item in data:
            if not isinstance(item, dict):
                continue
            pin_id = str(item.get("id", "")).strip()
            label = str(item.get("label", "")).strip()
            if not pin_id or pin_id in seen:
                continue
            seen.add(pin_id)
            self.pinned_packages.append({"id": pin_id, "label": label or "Pinned package"})

    def _save_pins_for_current_mdb(self):
        key = self._pin_settings_key()
        if key:
            self.settings.setValue(key, json.dumps(self.pinned_packages, ensure_ascii=False))
            self.settings.sync()

    def _source_row_for_pin(self, pin_id: str) -> Optional[int]:
        for row, rec in enumerate(self.model.records):
            if self._record_pin_id(rec) == pin_id:
                return row
        return None

    def _current_pin_id(self) -> Optional[str]:
        if self.current_source_row is None:
            return None
        rec = self.model.record(self.current_source_row)
        return self._record_pin_id(rec) if rec else None

    def _update_pin_button(self):
        if not hasattr(self, "pin_button"):
            return
        pin_id = self._current_pin_id()
        if not pin_id:
            self.pin_button.setEnabled(False)
            self.pin_button.setText("Pin")
            return
        pinned = any(item.get("id") == pin_id for item in self.pinned_packages)
        self.pin_button.setEnabled(True)
        self.pin_button.setText("Unpin" if pinned else "Pin")
        self.pin_button.setToolTip(
            "Remove this package from the persistent Pinned strip"
            if pinned else "Keep this package in the persistent Pinned strip"
        )

    def _toggle_current_pin(self):
        if self.current_source_row is None:
            return
        rec = self.model.record(self.current_source_row)
        if not rec:
            return
        pin_id = self._record_pin_id(rec)
        existing = next((i for i, item in enumerate(self.pinned_packages) if item.get("id") == pin_id), None)
        if existing is None:
            self.pinned_packages.append({"id": pin_id, "label": rec.get("Package", "Package")})
            self.statusBar().showMessage(f"Pinned: {rec.get('Package', 'Package')}", 4000)
        else:
            self.pinned_packages.pop(existing)
            self.statusBar().showMessage(f"Unpinned: {rec.get('Package', 'Package')}", 4000)
        self._save_pins_for_current_mdb()
        self._update_pin_button()
        self._refresh_pinned_bar()
        self._apply_gallery_quick_filters()

    def _remove_pin(self, pin_id: str):
        before = len(self.pinned_packages)
        self.pinned_packages = [item for item in self.pinned_packages if item.get("id") != pin_id]
        if len(self.pinned_packages) != before:
            self._save_pins_for_current_mdb()
            self._update_pin_button()
            self._refresh_pinned_bar()
            self._apply_gallery_quick_filters()

    def _refresh_pinned_bar(self):
        if not hasattr(self, "pinned_tabs_layout"):
            return
        while self.pinned_tabs_layout.count():
            item = self.pinned_tabs_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        current_id = self._current_pin_id()
        for pin in self.pinned_packages:
            pin_id = pin.get("id", "")
            row = self._source_row_for_pin(pin_id)
            label = pin.get("label", "Pinned package")
            if row is not None:
                rec = self.model.record(row)
                if rec and rec.get("Package"):
                    label = rec.get("Package", label)
                    pin["label"] = label

            chip = QFrame()
            chip.setObjectName("PinChip")
            chip_layout = QHBoxLayout(chip)
            chip_layout.setContentsMargins(0, 0, 0, 0)
            chip_layout.setSpacing(0)

            short = label if len(label) <= 34 else label[:31] + "…"
            tab = QPushButton(short if row is not None else "Missing: " + short)
            tab.setObjectName("PinnedTab")
            tab.setCheckable(True)
            tab.setChecked(pin_id == current_id)
            tab.setEnabled(row is not None)
            tab.setToolTip(label if row is not None else f"This pinned package is no longer present in this MDB: {label}")
            if row is not None:
                tab.clicked.connect(lambda checked=False, pid=pin_id: self._select_pinned_package(pid))
            chip_layout.addWidget(tab)

            remove_btn = QPushButton("×")
            remove_btn.setObjectName("PinnedRemove")
            remove_btn.setToolTip("Unpin")
            remove_btn.setFixedWidth(24)
            remove_btn.clicked.connect(lambda checked=False, pid=pin_id: self._remove_pin(pid))
            chip_layout.addWidget(remove_btn)
            self.pinned_tabs_layout.addWidget(chip)

        self.pinned_tabs_layout.addStretch(1)
        self.pinned_bar.setVisible(bool(self.pinned_packages))

    def _select_pinned_package(self, pin_id: str):
        row = self._source_row_for_pin(pin_id)
        if row is None:
            self.statusBar().showMessage("That pinned package is no longer present in this MDB", 5000)
            return
        source_index = self.model.index(row, 0)
        proxy_index = self.proxy.mapFromSource(source_index)
        if not proxy_index.isValid():
            self.clear_filters()
            proxy_index = self.proxy.mapFromSource(source_index)
        if not proxy_index.isValid():
            return
        self.table.setCurrentIndex(proxy_index)
        self.table.selectionModel().select(
            proxy_index,
            QItemSelectionModel.ClearAndSelect | QItemSelectionModel.Rows,
        )
        self.table.scrollTo(proxy_index, QTableView.PositionAtCenter)
        if hasattr(self, "gallery"):
            self.gallery.scrollTo(proxy_index, QAbstractItemView.PositionAtCenter)
        self.current_source_row = row
        rec = self.model.record(row)
        if rec:
            self.show_record(rec)

    def _set_open_last_mdb_on_startup(self, checked: bool):
        self.settings.setValue("open_last_mdb_on_startup", bool(checked))
        self.settings.sync()

    def _restore_last_mdb(self):
        raw_open_last = self.settings.value("open_last_mdb_on_startup", True)
        open_last = raw_open_last if isinstance(raw_open_last, bool) else str(raw_open_last).strip().lower() not in {"0", "false", "no", "off", ""}
        if not open_last:
            return
        saved = str(self.settings.value("last_mdb", "") or "").strip()
        if not saved:
            return
        path = Path(saved)
        if path.is_file():
            self.load_file(path)
        else:
            self.statusBar().showMessage("Last used Master Database was not found — open it from its new location or create a new MDB")

    def create_new_mdb(self):
        """Create a fresh Master Database from the bundled blank template."""
        if not self._confirm_pending_changes():
            return
        start_dir = Path(str(self.settings.value("last_dir", str(Path.home())) or Path.home()))
        suggested = start_dir / "Sims3_CC_Master_Database.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Create New Master Database",
            str(suggested),
            "Excel Workbook (*.xlsx)",
        )
        if not path:
            return
        target = Path(path)
        if target.suffix.lower() != ".xlsx":
            target = target.with_suffix(".xlsx")
        template = _asset_path("blank_master_database.xlsx")
        if not template.is_file():
            QMessageBox.critical(self, "Create MDB", "The bundled blank Master Database template is missing.")
            return
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(template, target)
        except Exception as exc:
            QMessageBox.critical(self, "Create MDB", f"Could not create the Master Database:\n{exc}")
            return
        self.settings.setValue("last_dir", str(target.parent))
        self.scan_mdb_override = False
        self.scan_mdb_path = str(target.resolve())
        if hasattr(self, "scan_mdb_edit"):
            self.scan_mdb_edit.setText(self.scan_mdb_path)
        self.statusBar().showMessage(f"Created new Master Database: {target}")
        self.load_file(target)

    def _build_guide_page(self):
        # Public-release oriented SBS guide. DEV-only experimental tools are
        # intentionally not documented here.
        page = QWidget()
        page.setObjectName("GuidePage")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        heading = QLabel("CAS Curator — Step-by-step Guide")
        heading.setObjectName("GuideTitle")
        layout.addWidget(heading)

        subtitle_text = (
            "CAS Curator is intended for CAS custom content. Keep script/gameplay mods outside the folder you scan. "
            "Profiles are optional organisational groups and do not move files."
            if IS_PUBLIC_BUILD else
            "CAS Curator DEV keeps the full development toolset. Use the compact topic menus below to jump straight to the instructions you need. "
            "Profiles remain optional organisational groups and do not move files."
        )
        subtitle = QLabel(subtitle_text)
        subtitle.setObjectName("GuideSubtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        self.guide_jump_buttons = []
        nav = QFrame()
        nav.setObjectName("GuideJumpCard")
        nav_layout = QHBoxLayout(nav)
        nav_layout.setContentsMargins(9, 7, 9, 7)
        nav_layout.setSpacing(8)

        library_label = QLabel("Library:")
        library_label.setObjectName("GuideGroupTitle")
        self.guide_library_combo = QComboBox()
        self.guide_library_combo.setMinimumWidth(225)
        library_topics = [
            ("Choose a Library topic…", ""),
            ("Getting started / MDB", "start"), ("Table & Gallery", "views"),
            ("Search & filters", "filters"), ("Select packages", "selection"),
            ("Pins", "pins"), ("Details", "details"),
            ("Edit & save", "edit"), ("Profiles", "profiles"),
            ("What can I do?", "workflows"),
            ("Advanced & experimental", "advanced"),
        ]
        for label, anchor in library_topics:
            self.guide_library_combo.addItem(label, anchor)

        scan_label = QLabel("Scan:")
        scan_label.setObjectName("GuideGroupTitle")
        self.guide_scan_combo = QComboBox()
        self.guide_scan_combo.setMinimumWidth(225)
        scan_topics = [
            ("Choose a Scan topic…", ""), ("Scan", "scanner"),
            ("Scan results", "results"), ("Sort files", "routing"),
            ("Recovery", "recovery"), ("Backups & logs", "backups"),
            ("Troubleshooting", "troubleshooting"),
        ]
        for label, anchor in scan_topics:
            self.guide_scan_combo.addItem(label, anchor)

        def jump_from_combo(combo):
            anchor = str(combo.currentData() or "")
            if anchor:
                self.guide_browser.scrollToAnchor(anchor)
                combo.blockSignals(True)
                combo.setCurrentIndex(0)
                combo.blockSignals(False)

        self.guide_library_combo.currentIndexChanged.connect(lambda _i: jump_from_combo(self.guide_library_combo))
        self.guide_scan_combo.currentIndexChanged.connect(lambda _i: jump_from_combo(self.guide_scan_combo))
        back_top = QPushButton("Back to top")
        back_top.setObjectName("GuideJumpButton")
        back_top.clicked.connect(lambda checked=False: self.guide_browser.scrollToAnchor("start"))

        nav_layout.addWidget(library_label)
        nav_layout.addWidget(self.guide_library_combo, 1)
        nav_layout.addSpacing(8)
        nav_layout.addWidget(scan_label)
        nav_layout.addWidget(self.guide_scan_combo, 1)
        nav_layout.addWidget(back_top)
        layout.addWidget(nav)

        self.guide_browser = QTextBrowser()
        self.guide_browser.setObjectName("GuideBrowser")
        self.guide_browser.setOpenExternalLinks(False)
        # Keep literal filenames/extensions readable in every Guide theme.
        # QTextBrowser's default <code> colour can otherwise become too faint
        # when a palette/theme is applied.
        self.guide_browser.document().setDefaultStyleSheet(
            "code { color: #202020; background-color: #F1F1F1; "
            "font-family: 'Consolas'; font-weight: 600; }"
        )
        self.guide_browser.setHtml(r'''
<h1><a name="start"></a>CAS Curator — Step-by-step Guide</h1>
<p>CAS Curator helps you browse, organise, scan and safely sort Sims 3 CAS custom content. Use the two compact topic menus below to jump straight to the instructions you need.</p>
<p><b>New user?</b> You do not need to make a spreadsheet yourself. If you do not already have a Master Database, click <b>Create new MDB…</b>, choose where to save it, then run your first Scan. CAS Curator will add the scanned packages to the new database.</p>
<p><b>Important:</b> Run Scan generates direct CASP IDs, supported CC Type values and Highest Polys from readable GEOM resources, extracts a dedicated THUM preview when one is available, and records the exact package hash and scan time. Whether a package contains or overlaps with EA resources is not assessed. Skin type, script/gameplay status, general IMG texture purpose and performance-risk ratings are also not supported scan results. Filenames are never used as evidence for the supported classifications.</p>

<h1>Library &amp; organisation</h1>
<h2><a name="library"></a>Library — open, create and browse your Master Database</h2>
<p>The Master Database (MDB) is CAS Curator's searchable index of your CC. Opening it does not move package files.</p>
<h3>If you are using CAS Curator for the first time</h3>
<ol>
<li>Stay on <b>Library</b> and click <b>Create new MDB…</b>.</li>
<li>Choose where you want to save <b>Sims3_CC_Master_Database.xlsx</b>.</li>
<li>CAS Curator opens the new blank MDB automatically and also uses it on the Scan page.</li>
<li>Open <b>Scan</b>, choose your CAS CC folder and click <b>Run Scan</b>. New package rows are added automatically.</li>
</ol>
<h3>If you already have an MDB</h3>
<ol>
<li>Click <b>Open MDB…</b>, or use <b>File &gt; Open Master Database…</b>.</li>
<li>Choose the <b>.xlsx</b> Master Database you want to use.</li>
<li>Wait until the package and thumbnail counts finish loading.</li>
</ol>
<p>By default CAS Curator remembers the last MDB and opens it next time. To use another one, simply click <b>Open MDB…</b>. You can turn automatic reopening off or on with <b>File &gt; Open last used MDB on startup</b>.</p>
<p>The Scan page always uses the Master Database currently open in Library. Scan shows that MDB path read-only; create, open, or change the Master Database from <b>Library</b>.</p>

<h2><a name="views"></a>Table and Gallery views</h2>
<p>Table is best for detailed information and bulk work. Gallery is best when you recognise CC visually.</p>
<h3>Table</h3>
<ol>
<li>Click <b>Table</b> at the top right of the Library.</li>
<li>Click a column heading to sort by that column.</li>
<li>Drag column headings left or right if you prefer a different order.</li>
<li>The <b>Columns</b> menu can hide or show individual columns whenever you want; it does not change the data. In Table view, the compact <b>CAS + / −</b>, <b>Health + / −</b> and <b>Organise + / −</b> buttons quickly show or fold related column groups. Your visible-column choices are remembered. Use <b>Columns &gt; Reset to default visible columns</b> whenever you want to restore the normal set. Advanced and Experimental columns are offered only while those options are enabled.</li>
</ol>
<h3>Gallery</h3>
<ol>
<li>Click <b>Gallery</b>.</li>
<li>Use the Gallery quick buttons for <b>In Game Use</b>, <b>To Keep</b>, <b>To Review</b>, <b>Has thumbnail</b>, <b>Pinned</b> or <b>Unsaved</b>.</li>
<li>Use the Gallery <b>CC Type</b> box to narrow the visual list further.</li>
<li>Gallery quick filters stack with Search, the permanent Library filters and any <b>+ Add Filter…</b> chips that are already active.</li>
<li>Click <b>Clear quick</b> to remove only the Gallery quick filters; your normal Library filters stay in place.</li>
</ol>

<h2><a name="filters"></a>Search and filters</h2>
<p>Narrow a large Library to the packages you actually want to work with without filling the top panel with dozens of permanent controls.</p>
<ol>
<li>Type in <b>Search</b> to look through package names and indexed text such as CASP IDs, types and notes.</li>
<li>Use the permanent <b>Status</b>, <b>CC Type</b>, <b>Creator</b> and <b>Profile</b> boxes for common everyday filtering.</li>
<li>For any other available data field, click <b>+ Add Filter…</b>, choose a field, condition and value, then click <b>Add Filter</b>.</li>
<li>Added filters appear as removable chips below the filter bar. Click a chip to remove just that filter. You can stack several added filters together with Search and the permanent filters.</li>
<li>The list of fields follows the current Library mode. Advanced and Experimental fields become filterable only while those options are enabled.</li>
<li>The <b>Shown</b> counter tells you how many packages match.</li>
<li>Click the small <b>×</b> inside Search to clear only the search text, or <b>Reset filters</b> to clear Search, permanent filters and all added-filter chips.</li>
</ol>
<p><b>How the conditions work:</b> <b>is</b> means an exact match; <b>contains</b> means the typed value appears anywhere inside the field. <b>is not</b> and <b>does not contain</b> are their opposites. <b>is blank</b> and <b>is not blank</b> do not need a Value. Text matching is not case-sensitive.</p>
<p>Examples: <b>Creator is Nightcrawler</b>, <b>Package contains lipstick</b>, <b>Age contains Teen</b>, <b>Outfit Category contains Formal</b>, or <b>CC Magic Import is Not in Import</b>. Several active filters work together: a package must match <b>all</b> of them to remain shown. The value box suggests existing values when practical, but you can also type your own text.</p>
<p>If you turn Advanced or Experimental features off, added filters that depend on those hidden fields are removed automatically.</p>
<p>Older MDBs may still physically contain legacy analysis/internal columns. CAS Curator preserves that data on disk but hides unsupported/internal fields from the normal Library views.</p>

<h2><a name="selection"></a>Select packages</h2>
<p>Selection lets you work with one package or prepare a group for a repeated edit.</p>
<ol>
<li><b>Normal click:</b> select one package.</li>
<li><b>Ctrl+click:</b> add or remove individual packages.</li>
<li><b>Shift+click:</b> select a continuous range.</li>
<li>The focused package is the one displayed in the right-side details panel.</li>
<li>For repeated metadata changes, edit one source cell and use the table fill handle to copy that exact value down displayed rows.</li>
</ol>

<h2><a name="pins"></a>Pins</h2>
<p>Pin packages you want to compare or return to without searching again.</p>
<ol>
<li>Select a package and click <b>Pin</b> beside its name.</li>
<li>A Pinned strip appears above the Library.</li>
<li>Click a pinned package to jump back to it.</li>
<li>Click its small <b>×</b>, or <b>Unpin</b>, to remove it.</li>
</ol>
<p>Pins are remembered separately for each Master Database.</p>

<h2><a name="details"></a>Right-side details panel</h2>
<p>Select a package in Table or Gallery to inspect it without losing your place.</p>
<ul>
<li><b>Overview:</b> thumbnail and compact package summary, including Age/Gender and Package Format when available. Advanced/Experimental details appear here too when their options are enabled.</li>
<li><b>CASP observations:</b> direct CASP information from the latest validated scan when an exact identity link is available.</li>
<li><b>Edit:</b> manual fields you are allowed to change.</li>
<li><b>All fields:</b> available when <b>Advanced features</b> are enabled; it shows the currently enabled user-facing stored fields in read-only form.</li>
</ul>
<p>Drag the vertical divider whenever you want more room for the Library or details panel.</p>

<h2><a name="edit"></a>Edit and save manual information</h2>
<p>Add your own organisation without changing the package itself.</p>
<ol>
<li>Select a package and open <b>Edit</b>.</li>
<li>Edit <b>Status</b>, <b>Creator</b>, <b>File Location</b>, <b>Merge Profiles</b> and/or <b>Notes</b>. Use <b>Add to:</b> for general Profile membership.</li>
<li>Use <b>Add / Replace image…</b> when you want your own preview. Choose a JPG, JPEG, PNG or BMP. A manual image takes priority over an automatically extracted package preview.</li>
<li>Click <b>Clear image</b> if you want to remove the manual image and let CAS Curator fall back to an automatically extracted package preview, if one exists.</li>
<li>An unsaved marker appears and the Unsaved counter increases.</li>
<li>Use <b>Revert this</b> to discard changes for the focused package, or <b>File &gt; Revert All Unsaved Changes</b> for everything.</li>
<li>Click <b>Save Changes</b> or press <b>Ctrl+S</b>. CAS Curator makes an automatic MDB backup first.</li>
</ol>
<h3>What the organisation fields mean</h3>
<ul>
<li><b>Status</b> — what you intend to do with the CC, such as In Game Use, To Keep or To Review.</li>
<li><b>Creator</b> — the creator or site you manually attribute the package to.</li>
<li><b>Profile</b> — any general group you create, such as a playthrough, testing set or style collection.</li>
<li><b>Merge Profile</b> — which merged collection or collections contain that package.</li>
<li><b>File Location</b> — the physical folder where the original unmerged package is stored.</li>
</ul>
<p>These fields are separate on purpose: changing one does not silently change the others.</p>
<h3>Status</h3>
<ul>
<li><b>In Game Use</b> — the original, unmerged package is used directly by the game.</li>
<li><b>To Keep</b> — worth retaining, but not currently intended for use.</li>
<li><b>To Review</b> — still needs a decision.</li>
<li><b>Blank</b> — no Status is assigned yet.</li>
</ul>
<p>Status describes what you intend to do with the CC; it is not a folder location.</p>
<p>If an older MDB still contains legacy Status labels, use <b>File &gt; Review legacy Status values…</b>. CAS Curator first shows a preview and stages the safe conversions rather than writing immediately: <b>Keep → To Keep</b>, <b>Review → To Review</b>, and old Favourite/Unreviewed/In Merge labels → blank. Legacy Delete/Deleted values are deliberately left unchanged. The normal backed-up <b>Save Changes</b> is still required.</p>
<h3>Creator</h3>
<p>Creator is optional manual information. CAS Curator does not guess or infer it automatically. Type the creator or site name yourself in Edit, edit the Creator cell in the Table, or select several packages and use <b>Set Creator for Selected…</b> from Tools or the Table right-click menu. Bulk Creator edits remain pending until <b>Save Changes</b>, so they can still be reverted before saving. Creator names are included in normal Library search, and <b>Creator</b> also has its own permanent filter at the top of Library.</p>
<p>When Creator is filled, <b>Search Web for Package</b> includes it in the search query, and the row right-click menu also offers <b>Search Web for Creator</b>.</p>
<h3>File Location</h3>
<p>File Location records the physical folder containing the original, unmerged package. Choosing a File Location records the association only; it does not move, copy or delete the file.</p>
<p>Use <b>Manage Locations…</b> to create friendly location names and link each name to its real folder. Each saved location has its own <b>Open</b> button, and the dialog shows the currently selected location explicitly for rename/path/remove actions. <b>Open Folder</b> in Edit opens exactly the saved folder for the selected package. You can also right-click a package row with a File Location assigned and choose <b>Open File Location</b>. If that folder is unavailable, CAS Curator reports it as unavailable and does not guess or substitute another folder. Older Library Location categories are preserved separately and are never treated as physical paths.</p>
<p><b>Controlled managed moves:</b> select one or more packages and use <b>Move Selected to Location…</b> from Tools or the Table right-click menu. CAS Curator resolves the original package conservatively, previews every destination, refuses overwrite/name conflicts, copies and verifies SHA-256 before deleting the original, then writes the new File Location to the MDB with the normal backup. If the MDB update fails, CAS Curator attempts to roll the physical move back.</p>
<p><b>CC Magic Import is copy-only:</b> first set the destination with <b>Tools &gt; Set CC Magic Import Folder…</b>, then use <b>Copy Selected to CC Magic Import…</b>. The organised original stays in place and File Location does not change. The live <b>CC Magic Import</b> column is derived from the real Import folder and is not written into the MDB. When you want to filter by it, use <b>+ Add Filter… &gt; CC Magic Import</b>.</p>
<p><b>CC Magic Import and File Location are different:</b> CC Magic Import is the destination for a deployment copy; <b>File Location</b> tells CAS Curator where the original/master package lives. Scanning records package metadata but does not guarantee that a managed File Location has been assigned. If Copy to CC Magic says the original has no valid File Location, either <b>route the package first</b> so CAS Curator records its routed folder, or select the package and assign <b>File Location</b> in <b>Edit</b> (use <b>Manage Locations…</b> there first if you need to add the folder).</p>
<ul>
<li><b>In Import</b> — a same-name Import copy is present and matches the package identity/hash CAS Curator can verify.</li>
<li><b>Not in Import</b> — no same-name package is present in the configured Import folder.</li>
<li><b>Conflict</b> — the same filename exists in Import, but it is a different file/hash.</li>
<li><b>Unverifiable</b> — a file is present but CAS Curator cannot safely verify that it is the same package.</li>
<li><b>Unavailable</b> — the Import folder is not configured or cannot currently be accessed.</li>
</ul>
<p><b>After CC Magic processes the Import folder:</b> manage or remove that content from inside CC Magic itself. CAS Curator does not delete or move files from CC Magic-managed storage. If you use CAS Curator Status or Profiles to track your own workflow, update those manually after making changes in CC Magic.</p>
<p>To look for a package online, right-click its row and choose <b>Search Web for Package</b>. CAS Curator opens a Google search in your default browser using the package name with the <code>.package</code> ending removed and common separators cleaned up. If Creator is filled in, it is included automatically to make the search more specific.</p>
<h3>Merge Profiles</h3>
<p>Merge Profiles record which merged collections contain the package. A package can belong to several Merge Profiles at once. Merge membership is separate from Status and from general Profiles.</p>
<p>Use <b>Set Merge Profiles…</b> to choose one or several Merge Profiles for the focused package. <b>Manage Merge Profiles…</b> lets you create, rename or remove names in the managed list. Assignment and list changes remain pending until <b>Save Changes</b>.</p>
<h3>Fill-down</h3>
<p>Select an editable source cell and drag its small fill handle downward across displayed rows. CAS Curator copies the exact value to those displayed records, respects sorting/filtering, and keeps the edits pending until Save Changes. Blank source values are blocked. Use <b>Options &gt; Confirm before fill-down</b> if you want a confirmation before each fill operation; turn it off if you prefer direct fill-down.</p>
<h3>Delete Record</h3>
<p>Right-click a Table row and choose <b>Delete Record…</b> to stage removal of the CAS Curator Library/MDB record only. Saving creates the normal MDB backup. The actual <code>.package</code> file is never deleted or moved. If that physical package is scanned again later, CAS Curator can create a Library record for it again.</p>

<h3>Add or replace a preview image</h3>
<p>If a package has no automatic preview, or you prefer a different picture, you can add one manually.</p>
<ol>
<li>Select the package and open <b>Edit</b>.</li>
<li>Click <b>Add / Replace image…</b> and choose your JPG, JPEG, PNG or BMP.</li>
<li>Click <b>Save Changes</b>. The Library will show <b>Thumbnail = Yes</b> and use your chosen picture.</li>
<li>You can safely scan the package again later. CAS Curator keeps your manual image and does not overwrite it with an automatic package preview.</li>
<li>To return to the automatic preview, click <b>Clear image</b>, save, then rescan if needed. If the package has a supported dedicated preview resource, CAS Curator can extract it again.</li>
</ol>

<h2><a name="profiles"></a>Profiles</h2>
<p>Profiles are optional groups for a playthrough, theme, CC Magic setup, testing set, merge plan, character style, or anything else you choose. A package can belong to more than one Profile.</p>
<ol>
<li>Create an empty Profile with <b>New Profile…</b>.</li>
<li>For the focused package, use <b>Add to:</b> and choose an existing Profile.</li>
<li><b>Profiles assigned to this package</b> is read-only; it shows the current membership but is not a second editing box.</li>
<li><b>Manage Profiles…</b> lets you create, rename or delete Profiles and see how many packages use each one.</li>
<li>To apply the same Profile assignment to several displayed packages, prepare one source row and use fill-down.</li>
<li>Profile changes stay pending until <b>Save Changes</b>.</li>
</ol>

<h2><a name="workflows"></a>What can I do with CAS Curator?</h2>
<p>CAS Curator's tools are designed to work together. You can use the Library to find a group, Profiles and Creator to describe it, File Location to track where the physical packages live, and controlled Move/Copy actions when you deliberately want to reorganise files.</p>
<h3>Group one creator's CC into a folder</h3>
<ol>
<li>Select the matching packages and use <b>Set Creator for Selected…</b> to assign the creator once. The Table fill handle remains available when a contiguous fill-down is more convenient.</li>
<li>Use the permanent <b>Creator</b> filter so only that creator's packages are shown.</li>
<li>Select the shown packages.</li>
<li>Use <b>Move Selected to Location…</b>, review the source/destination preview, then confirm.</li>
</ol>
<p>Routing already provides automatic evidence-based sorting by supported type. Managed Locations are most useful for deliberate organisation such as creator folders, Testing, Archive or another collection structure that you choose yourself.</p>
<h3>Move a filtered group together</h3>
<p>Filter by <b>CC Type</b>, <b>Status</b>, <b>Creator</b> or <b>Profile</b> directly, or use <b>+ Add Filter…</b> for fields such as <b>File Location</b>, Age, Gender or any other currently available column. Select the rows you want, then use <b>Move Selected to Location…</b>. CAS Curator verifies package identity and SHA-256 before completing controlled moves.</p>
<h3>Prepare a group for CC Magic</h3>
<p>Click <b>+ Add Filter…</b> and choose <b>CC Magic Import is Not in Import</b>, then select the packages you want and use <b>Copy Selected to CC Magic Import…</b>. The Import folder is only a deployment drop-in location. Once CC Magic processes those files, use CC Magic itself to manage or remove them. CAS Curator does not delete files from CC Magic-managed storage.</p>
<h3>Bring in a folder of new downloads</h3>
<p>Use the recommended sequence <b>Run Scan → Preview Routing → Run Routing</b>. Scan adds or refreshes MDB knowledge; Preview Routing shows exactly where files would go; Run Routing performs the physical move. When ZIPs are present, Scan can extract them beside the original archives first so newly extracted packages join the same scan. When Sims3Packs are present, CAS Curator can analyse their manifests, avoid extracting a CAS package when an identical standalone <code>.package</code> is already present, and optionally extract missing CAS package copies while leaving the original Sims3Pack untouched.</p>
<h3>Find an item again later</h3>
<p>Use Creator, package-name search, Profiles, Status and File Location to narrow the Library.</p>
<p><b>Useful package shortcuts:</b> right-click a Table row for <b>Copy package name</b>, <b>Search Web for Package</b>, <b>Search Web for Creator</b> when Creator is filled, <b>Open File Location</b> when a physical folder is tracked, Pin/Unpin and the available managed Move/CC Magic actions.</p>
<h3>Review provisional metadata without letting it organise files</h3>
<p>Enable <b>Experimental features</b> to inspect Subtype, Skin Type, Outfit Category and EA Resource Status. These fields are review information only and never control Routing automatically.</p>

<h2><a name="advanced"></a>Advanced and experimental features</h2>
<p>Advanced and Experimental are optional and are off by default. Turning them off hides the extra information; it does not delete anything from the Master Database. Age and Gender are normal metadata and do not require either option.</p>
<ul>
<li><b>Options &gt; Enable Advanced features</b> shows extra stored technical information such as CASP Instance IDs, Biggest Texture, HQ/Poly warnings, Risk, Health Issues, Duplicate Files, Last Scanned, Package Hash and the All fields tab. Some of these are legacy/technical fields and are not regenerated by the conservative scanner.</li>
<li><b>Age</b> and <b>Gender</b> are normal read-only package metadata derived from direct CASP observations.</li>
<li><b>Options &gt; Enable Experimental features</b> shows provisional scanner-derived information such as Subtype, Skin Type, Outfit Category and EA Resource Status when available. These results are still being validated and may be incomplete or incorrect.</li>
<li><b>Subtype</b> is still Experimental. Current Build/Buy subtype values such as Wall, Rug or Lighting are filename hints and can be wrong when a filename contains several descriptive words. Scripted Object subtypes use package/script evidence. The Overview identifies which kind of evidence produced the displayed subtype.</li>
<li><b>Skin Type</b> is still Experimental. A TONE resource can directly identify a package as Skin, but the current Default Replacement / Non-Default Skin Type label is inferred from filename clues. If no direct Skin classification is present, the stored value may come from the broader possible-skin heuristic. The Overview identifies which evidence path applies.</li>
<li><b>Outfit Category</b> is Experimental display metadata read directly from each CASP ClothingCategory observation and combined for the package. It currently shows the seven normal CAS filters: Everyday, Formal, Sleepwear, Swimwear, Athletic, Career and Outerwear. Technical/special-purpose CASP flags such as ValidForRandom are intentionally not included in this summary.</li>
<li><b>EA Resource Status</b> is a plain-language view of the stored experimental EA-overlap result. It may show <b>No EA overlap detected</b>, <b>EA resource overlap — review</b>, <b>Likely EA override / default replacement</b>, or <b>Matches EA resources</b>. Older stored results may come from an earlier scanner pass, so treat them as review information rather than proof.</li>
</ul>
<h3>CASP observations</h3>
<p><b>CASP</b> means <b>Create-a-Sim Part</b>. The CASP observations tab shows information read directly from CASP resources in the package rather than guessed from its filename. This can include CASP Instance IDs, CC Type, Age, Gender and, while Experimental features are enabled, Outfit Category. A package can contain more than one CASP, so CAS Curator may combine the observed values into a concise package summary. These observations describe the CASP resources that were found; they do not claim to describe every other resource in the package.</p>
<p>Experimental information is for review and testing only. It does not change Routing or move package files automatically.</p>

<h1>Scan &amp; file sorting</h1>
<h2><a name="scanner"></a>Scan</h2>
<p>Scan a folder of CAS package files and record direct identity/CASP evidence in the selected Master Database. The same Scan works for a small folder or a large batch.</p>
<ol>
<li>Open <b>Scan</b>.</li>
<li>Choose the <b>CC source folder</b>. Package files inside its subfolders are included too. Use <b>Open folder</b> if you want to open the selected source in Windows Explorer.</li>
<li><b>Library and Scan always use the same Master Database.</b> Scan shows the active Library MDB as a read-only path. To create, open, or change the Master Database, use the <b>Library</b> page.</li>
<li>If you have unsaved manual Library edits, save or revert them first.</li>
<li>Click <b>Run Scan</b>. A normal scan starts immediately after validation; CAS Curator only opens a choice box when something genuinely needs your decision, such as ZIP/Sims3Pack handling or MDB compatibility.</li>
</ol>
<p><b>Progress reaches 100% only when the whole Scan is finished.</b> After the last package is read, CAS Curator shows a finalising/refreshing stage while the MDB, result files and Library refresh complete.</p>
<p><b>Run Scan does not move source package files.</b> If ZIP archives are found, CAS Curator first asks whether you want to extract them beside the original ZIPs; extraction happens only when you explicitly choose it, and the ZIPs themselves are kept. A full workbook backup is made before scanner writes. Supported rows can receive CASP IDs, CC Type from CASP metadata, Highest Polys from readable GEOM resources, the best validated dedicated Sims 3 preview/thumbnail resource when present, the exact package SHA-256 hash and Last Scanned. Existing/manual preview images are kept and are not overwritten.</p>
<p><b>MDB compatibility:</b> some spreadsheet programs save extra workbook information that the scanner cannot safely read. If CAS Curator detects this, it offers to create and verify a separate <b>CAS Curator Compatible</b> copy. The original workbook is always left untouched. Use the compatible copy for scanning.</p>
<p>Use CAS Curator's conservative scanner for CAS CC such as hair, clothing, shoes, accessories, makeup and CAS presets. Keep gameplay/script mods out of the scan batch because the scanner does not assess script/gameplay content.</p>

<h2><a name="results"></a>Understand scan results</h2>
<ul>
<li><b>CASP IDs:</b> direct CASP resource identifiers observed in the package.</li>
<li><b>CC Type:</b> generated from the CASP Clothing metadata when every observed CASP in that package can be mapped safely. Mixed valid types are shown as <b>Mixed CAS</b>.</li>
<li><b>Highest Polys:</b> the highest GEOM <b>Faces.Count</b> found in the package. The value is withheld if any GEOM resource could not be read.</li>
<li><b>Thumbnail:</b> extracted only from validated dedicated Sims 3 THUM / ICON / SNAP preview-resource families. CAS Curator chooses the best supported dedicated preview and does not use a random IMG texture. If there is no supported preview, you can add your own image later from Edit.</li>
<li><b>Package Hash:</b> an exact SHA-256 fingerprint used for identity and safety checks.</li>
<li><b>Last Scanned:</b> when CAS Curator last completed its safe scan write for that row.</li>
<li><b>CASP observations:</b> direct CASP metadata linked to the exact package identity after the current scan.</li>
<li><b>Age / Gender:</b> friendly read-only summaries derived from the direct CASP Age/Gender observations for that package. They are display metadata and are not written back into CC Master Database.</li>
<li><b>Held for review:</b> the package identity/evidence was not safe enough for the normal write path.</li>
</ul>
<p><b>Not part of the supported scan results:</b> Whether a package contains or overlaps with EA resources is not assessed. Script/gameplay status, general IMG texture purpose, poly-count performance risk, texture-size performance risk and new Health/Risk ratings are also not supported results. Experimental features may expose provisional classifications such as Skin Type, Outfit Category and EA Resource Status for testing, but they are not treated as supported scanner findings and never control Routing automatically. CC Type, Highest Polys and dedicated preview extraction are direct supported scanner results, not performance/safety ratings.</p>

<h2><a name="routing"></a>Sort files — Preview Routing, then Run Routing</h2>
<p>Sort a reviewed folder into evidence-based subfolders while seeing every destination before anything moves.</p>
<p><b>Recommended order: Run Scan → Preview Routing → Run Routing.</b> Scan adds or refreshes MDB records and scanner-derived metadata. Preview Routing reads that information and calculates destinations without moving files. Run Routing performs the physical move. Routing can still be continued deliberately without a fresh Scan, but CAS Curator warns when packages are not in the MDB or their current file hash no longer matches the stored scan identity.</p>
<h3>Pre-routing scan-state check</h3>
<p>Before Preview Routing or Run Routing, CAS Curator compares the current package files with the selected MDB:</p>
<ul>
<li><b>Ready</b> — the package is represented in the MDB and its stored Package Hash matches the current file.</li>
<li><b>Changed since scan</b> — a matching package name exists, but the current file no longer matches a stored Package Hash, or the older row has no hash that can verify the current file. Run Scan first to refresh scanner-derived details.</li>
<li><b>Not in MDB</b> — no matching MDB row exists. Run Scan first if you want a complete Library record and reliable post-routing tracking.</li>
</ul>
<p>When attention is needed, choose <b>Run Scan First</b>, <b>Continue Routing Anyway</b>, or <b>Cancel</b>. The warning itself is the final confirmation for that operation: choosing Continue Routing Anyway starts Preview Routing or Run Routing directly, so you are not asked to confirm the same operation a second time. Manual-only MDB edits such as Creator, Notes, Profiles or Status do not make a package need rescanning because they do not change the physical package file.</p>
<h3>Preview first</h3>
<ol>
<li>Choose the <b>CC source folder</b> and a separate <b>Sort destination</b>. Use <b>Open folder</b> beside Sort destination whenever you want to inspect that exact folder in Windows Explorer.</li>
<li>Click <b>Preview Routing</b>.</li>
<li>Review File, evidence route, Duplicate, Destination and Status in the Routing preview tab.</li>
<li>Preview moves no source files.</li>
</ol>
<h3>Then run routing</h3>
<ol>
<li>Continue only when you are happy with the fresh preview.</li>
<li>Tick <b>I understand: Run Routing will MOVE source files</b>.</li>
<li>Click <b>Run Routing</b> and let it finish without manually moving the same files.</li>
</ol>
<p>After a successful routing run, CAS Curator uses its own move manifest to record the exact routed folder for each safely matched <code>.package</code> row as a File Location. This lets later <b>Open File Location</b> and <b>Move Selected to Location…</b> actions work without rescanning just to rediscover a move CAS Curator already performed. If <b>Recover Last Routing</b> later restores those files, CAS Curator also restores safely matched File Location records to the original folders. Ambiguous matches are skipped rather than guessed.</p>
<p>Routes such as <b>CASP_OBSERVED/Hair</b> describe observed CASP evidence; they are not a claim that every resource in the whole package has been classified. Exact duplicate copies are routed separately to <b>DUPLICATES</b>.</p>
<p>Loose image files are routed under <b>IMAGES</b> while keeping the selected source-folder name and any subfolders beneath it. For example, an image from a selected folder named <b>Nightcrawler</b> is kept under <b>IMAGES/Nightcrawler/...</b> rather than being flattened into the IMAGES root.</p>
<p><b>ZIPs found during Scan:</b> when the selected Scan folder contains ZIP archives, CAS Curator asks before Scan starts. Choose <b>Extract &amp; Continue</b> to unpack each archive into its own folder beside the ZIP and include newly extracted package files in the same Scan, or <b>Keep Packed</b> to leave them untouched for normal ARCHIVES routing later. The prompt is the extraction decision; there is no second extraction confirmation or completion popup. Original ZIPs are never deleted or overwritten.</p>
<p><b>ZIP preview/extract:</b> you can also use <b>Tools &gt; Preview / Extract ZIP…</b>, or right-click a ZIP in Routing Preview. CAS Curator previews the archive first, preserves an existing single top-level folder, and uses the ZIP filename as an outer folder when files are loose or use multiple roots. Extraction is explicit, keeps the original ZIP, rejects unsafe paths, and does not overwrite existing files.</p>
<p><b>Sims3Packs found during Scan:</b> CAS Curator can read the Sims3Pack manifest without modifying the original file. For the conservative CAS workflow, only Sims3Packs declared as <b>CASpart</b> are eligible for package-copy extraction. If an identical standalone <code>.package</code> is already somewhere in the selected Scan tree, CAS Curator recognises it by SHA-256 and does not extract the same embedded package again. If a CASpart Sims3Pack has no equivalent standalone package, choose <b>Extract CAS &amp; Continue</b> to place the embedded package copy inside a folder named after the Sims3Pack, beside the untouched original. Extracted files use the original Sims3Pack name instead of the internal hexadecimal package name. A Sims3Pack can contain several CAS package parts; CAS Curator reports these as one multi-part set, duplicate-checks every part separately, and extracts all missing parts together using stable <code>__1</code>, <code>__2</code> and so on. The parts are not merged. Choose <b>Keep Packed &amp; Continue</b> to make no extraction changes. The prompt is the Sims3Pack decision; there is no extra extraction-complete popup before Scan continues.</p>
<p><b>Protected Sims3Packs:</b> Lots, worlds, Sims/households, patterns, objects and unknown/mixed types are deliberately kept packed in this first conservative implementation. CAS Curator does not try to turn them into Mods/Packages content. The existing Routing system can continue to move the original Sims3Pack intact to <b>SIMS3PACK</b>. This restriction is intentionally stricter than a general Sims3Pack converter because CAS Curator is focused on safe CAS review.</p>

<h2><a name="recovery"></a>Recover the last routing run</h2>
<ol>
<li>Click <b>Recover Last Routing</b>.</li>
<li>Read the confirmation carefully.</li>
<li>CAS Curator checks the saved move record and package hashes before moving anything back.</li>
<li>Recovery refuses to overwrite an existing original path. If something changed manually, investigate instead of forcing it.</li>
</ol>
<p>The button applies to the last unrecovered routing manifest available in the current session.</p>

<h2><a name="backups"></a>Backups, logs and result files</h2>
<ul>
<li><b>MDB backup:</b> scanner/manual writes create a backup before changing the workbook.</li>
<li><b>Scanner log:</b> the embedded log remains on the Scan page. Use <b>Open Log</b>, <b>Copy Log</b> or <b>Expand Log</b> only when you want a larger/full view. CAS Curator no longer opens a separate large completion window automatically.</li>
<li><b>Routing preview:</b> use <b>Open Routing Preview</b> to inspect the saved dry-run output.</li>
<li><b>Sort destination:</b> use <b>Open folder</b> beside Sort destination to open the configured routing folder directly.</li>
<li><b>Move manifest:</b> records original and routed paths plus hashes for a completed routing run.</li>
</ul>
<p>Keep the backup and move manifest until you are satisfied that a large operation is correct.</p>

<h2><a name="troubleshooting"></a>Troubleshooting</h2>
<p><b>No Master Database yet:</b> go to <b>Library</b>, click <b>Create new MDB…</b>, choose where to save it, then run your first Scan.</p>
<p><b>The wrong MDB opens at startup:</b> click <b>Open MDB…</b> and choose the one you want. That becomes the remembered MDB. You can also untick <b>File &gt; Open last used MDB on startup</b>.</p>
<p><b>Scan shows the wrong MDB:</b> go to <b>Library</b> and open the Master Database you want. Scan updates automatically and cannot use a separate MDB.</p>
<p><b>Run Scan is unavailable:</b> open a valid .xlsx Master Database in <b>Library</b>, choose a valid source folder, and save/revert pending manual Library edits.</p>
<p><b>Scan reports an MDB compatibility issue:</b> some spreadsheet programs save extra workbook information that the scanner cannot safely read. Choose <b>Create compatible copy…</b>. CAS Curator leaves your original workbook untouched, creates and verifies a separate compatible copy, and switches Scan to that copy. If the affected MDB was open in Library, the compatible copy is opened there too. Run Scan again after the copy is ready.</p>
<p><b>Run Routing is unavailable:</b> choose a valid source and separate destination, create a fresh Routing Preview, then tick the move acknowledgement.</p>
<p><b>Routing says Changed since scan / Not in MDB:</b> choose <b>Run Scan First</b> when you want current scanner-derived metadata and reliable tracking. <b>Continue Routing Anyway</b> is available for deliberate cases, but stale or unknown packages may not have complete MDB information.</p>
<p><b>Gallery looks empty:</b> click <b>Clear quick</b>, then <b>Reset filters</b> if needed.</p>
<p><b>Older fields say Unverified (legacy):</b> that means the value came from an older workflow and has not been regenerated by the current conservative scanner. Re-scan to generate the supported core fields.</p>
<p><b>A package was held for review:</b> CAS Curator did not have enough safe identity/evidence to make the normal scanner write. Do not rename it to force a result; filenames are not classification evidence.</p>
<p><b>A routing preview looks wrong:</b> do not run routing. Change the setup or leave the package for review.</p>
<p><b>A scan/routing operation fails:</b> keep the Scanner log, backup and any preview/move manifest.</p>

''')
        layout.addWidget(self.guide_browser, 1)
        return page

    def _build_scan_shell_page(self):
        # The scanner page is scrollable as a safety net for shorter windows /
        # display scaling. The cards themselves keep their intended height, so
        # Qt no longer crushes the setup rows to make everything fit.
        scroll = QScrollArea()
        scroll.setObjectName("ScanPageScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        page = QFrame()
        page.setObjectName("ScanPage")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        title = QLabel("CAS Scanner" if IS_PUBLIC_BUILD else "CC Health Scanner")
        title.setObjectName("ScanTitle")
        layout.addWidget(title)

        subtitle = QLabel(
            "Scan checks CAS package files and updates the selected Master Database without moving the source files. "
            "Preview Routing shows where files would be sorted; Run Routing moves them only after you confirm."
            if IS_PUBLIC_BUILD else
            "Scan updates the selected Master Database without moving source files. "
            "Routing Preview is a dry run; Run Routing physically moves files to the selected sort destination."
        )
        subtitle.setObjectName("ScanSubtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        source_card = QFrame()
        source_card.setObjectName("ScanCard")
        source_layout = QVBoxLayout(source_card)
        source_layout.setContentsMargins(14, 11, 14, 12)
        source_layout.setSpacing(9)

        source_heading = QLabel("Scan setup")
        source_heading.setObjectName("ScanSectionTitle")
        source_layout.addWidget(source_heading)

        # Use explicit fixed-height horizontal rows rather than QFormLayout.
        # This is intentionally a little more rigid: it prevents Qt from
        # vertically compressing the path controls until they overlap.
        label_width = 230
        action_height = 34
        row_height = 42
        action_width = 150
        source_card.setMinimumHeight(255 if IS_PUBLIC_BUILD else 305)

        def make_label(text):
            label = QLabel(text)
            label.setFixedWidth(label_width)
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            return label

        def make_row(label_text):
            row = QWidget()
            row.setObjectName("ScanSetupRow")
            row.setFixedHeight(row_height)
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(6, 4, 6, 4)
            row_layout.setSpacing(6)
            row_layout.addWidget(make_label(label_text))
            return row, row_layout

        # Source folder row.
        source_row, source_row_layout = make_row("CC source folder:")
        self.scan_source_edit = QLineEdit(self.scan_source_path)
        self.scan_source_edit.setObjectName("ScanPathField")
        self.scan_source_edit.setReadOnly(True)
        self.scan_source_edit.setFixedHeight(action_height)
        self.scan_source_edit.setPlaceholderText("Choose the folder containing the CC to scan…")
        source_row_layout.addWidget(self.scan_source_edit, 1)
        source_browse = QPushButton("Choose folder…")
        source_browse.setFixedSize(action_width, action_height)
        source_browse.clicked.connect(self._choose_scan_source_folder)
        source_row_layout.addWidget(source_browse)
        self.scan_source_open_btn = QPushButton("Open folder")
        self.scan_source_open_btn.setFixedSize(110, action_height)
        self.scan_source_open_btn.setToolTip("Open the currently selected CC source folder in Windows File Explorer.")
        self.scan_source_open_btn.clicked.connect(self._open_scan_source_folder)
        source_row_layout.addWidget(self.scan_source_open_btn)
        source_layout.addWidget(source_row)

        # MDB row.
        mdb_row, mdb_row_layout = make_row("Master Database:")
        self.scan_mdb_edit = QLineEdit()
        self.scan_mdb_edit.setObjectName("ScanPathField")
        self.scan_mdb_edit.setReadOnly(True)
        self.scan_mdb_edit.setFixedHeight(action_height)
        self.scan_mdb_edit.setReadOnly(True)
        self.scan_mdb_edit.setPlaceholderText("No Master Database open — open or create one in Library.")
        self.scan_mdb_edit.setToolTip("The Scan page uses the Master Database currently open in Library. Change it from Library.")
        mdb_row_layout.addWidget(self.scan_mdb_edit, 1)
        source_layout.addWidget(mdb_row)

        # Scanner mode controls are DEV-only. PUBLIC uses one conservative scan
        # workflow, matching the distributed v1.0 UI; keep an internal default
        # combo only because the shared scanner runner still expects a mode code.
        self.scan_mode_combo = ArrowComboBox()
        self.scan_mode_combo.setObjectName("ScanModeCombo")
        self.scan_mode_combo.setFixedHeight(action_height)
        if IS_PUBLIC_BUILD:
            self.scan_mode_combo.addItem("Individual / review folder")
            self.scan_mode_combo.setVisible(False)
        else:
            mode_row, mode_row_layout = make_row("Source mode:")
            self.scan_mode_combo.addItems([
                "Individual / review folder",
                "Big Review / bulk",
                "Thumbnail backfill",
            ])
            self.scan_mode_combo.currentTextChanged.connect(lambda _text: self._update_scan_start_state())
            mode_row_layout.addWidget(self.scan_mode_combo, 1)
            edition_caption = QLabel("Edition:")
            edition_caption.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            edition_caption.setFixedWidth(70)
            mode_row_layout.addWidget(edition_caption)
            edition = QLabel("DEV")
            edition.setObjectName("ScanValueBadge")
            edition.setAlignment(Qt.AlignCenter)
            edition.setFixedSize(95, action_height)
            mode_row_layout.addWidget(edition)
            source_layout.addWidget(mode_row)

        # Dry-run sorting destination row.
        route_row, route_row_layout = make_row("Sort destination:")
        self.routing_root_edit = QLineEdit(self.routing_preview_root)
        self.routing_root_edit.setObjectName("ScanPathField")
        self.routing_root_edit.setReadOnly(True)
        self.routing_root_edit.setFixedHeight(action_height)
        self.routing_root_edit.setPlaceholderText("Choose a separate folder for routing preview / sorted files…")
        self.routing_root_edit.setToolTip("Routing Preview calculates destinations here without moving files; Run Routing moves files here.")
        route_row_layout.addWidget(self.routing_root_edit, 1)
        routing_browse = QPushButton("Choose folder…")
        routing_browse.setFixedSize(action_width, action_height)
        routing_browse.setToolTip("Choose the destination root used by Routing Preview and Run Routing.")
        routing_browse.clicked.connect(self._choose_routing_preview_root)
        route_row_layout.addWidget(routing_browse)
        self.routing_root_open_btn = QPushButton("Open folder")
        self.routing_root_open_btn.setFixedSize(110, action_height)
        self.routing_root_open_btn.setToolTip("Open the currently selected Sort destination in Windows File Explorer.")
        self.routing_root_open_btn.clicked.connect(self._open_routing_preview_root)
        route_row_layout.addWidget(self.routing_root_open_btn)
        source_layout.addWidget(route_row)

        destination_hint = QLabel("Routing Preview moves nothing. Run Routing physically moves files into this destination after confirmation.")
        destination_hint.setObjectName("ScanHint")
        destination_hint.setWordWrap(True)
        destination_hint.setContentsMargins(label_width + 10, 0, 0, 0)
        source_layout.addWidget(destination_hint)

        # Windows testing showed that a draggable embedded log boundary is not
        # reliable on all setups. RC5 uses a larger fixed/default output area
        # instead, and reclaims vertical space with one compact action row below.
        scan_setup_section = QWidget()
        scan_setup_section.setObjectName("ScanSetupSection")
        scan_setup_layout = QVBoxLayout(scan_setup_section)
        scan_setup_layout.setContentsMargins(0, 0, 0, 0)
        scan_setup_layout.setSpacing(8)
        scan_setup_layout.addWidget(source_card)

        # Compact progress card: keep the useful information, but don't spend
        # a large block of vertical space on it.
        progress_card = QFrame()
        progress_card.setObjectName("ScanCard")
        progress_card.setMaximumHeight(112)
        progress_layout = QVBoxLayout(progress_card)
        progress_layout.setContentsMargins(12, 7, 12, 8)
        progress_layout.setSpacing(4)

        progress_top = QHBoxLayout()
        progress_top.setSpacing(8)
        progress_heading = QLabel("Scan progress")
        progress_heading.setObjectName("ScanSectionTitle")
        progress_top.addWidget(progress_heading)
        self.scan_status_label = QLabel("Ready — choose a source folder and Master Database")
        self.scan_status_label.setObjectName("ScanStatus")
        progress_top.addWidget(self.scan_status_label, 1)
        self.scan_progress_label = QLabel("0 / 0")
        self.scan_progress_label.setObjectName("ScanProgressLabel")
        progress_top.addWidget(self.scan_progress_label)
        progress_layout.addLayout(progress_top)

        self.scan_progress = QProgressBar()
        self.scan_progress.setRange(0, 100)
        self.scan_progress.setValue(0)
        self.scan_progress.setTextVisible(True)
        self.scan_progress.setFixedHeight(18)
        progress_layout.addWidget(self.scan_progress)

        current_row = QHBoxLayout()
        current_row.setSpacing(6)
        current_row.addWidget(QLabel("Current file:"))
        self.scan_current_file_label = QLabel("—")
        self.scan_current_file_label.setObjectName("ScanCurrentFile")
        self.scan_current_file_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        current_row.addWidget(self.scan_current_file_label, 1)
        progress_layout.addLayout(current_row)
        scan_setup_layout.addWidget(progress_card)
        scan_setup_section.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        layout.addWidget(scan_setup_section)

        log_card = QFrame()
        log_card.setObjectName("ScanCard")
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(12, 10, 12, 11)
        log_layout.setSpacing(7)
        log_top = QHBoxLayout()
        log_heading = QLabel("Scanner log")
        log_heading.setObjectName("ScanSectionTitle")
        log_top.addWidget(log_heading)
        log_top.addStretch(1)

        self.scan_open_log_btn = QPushButton("Open Log")
        self.scan_open_log_btn.setEnabled(False)
        self.scan_open_log_btn.setToolTip("Open the saved Health_Scan_Log.txt from the most recent scan.")
        self.scan_open_log_btn.clicked.connect(self._open_last_scan_log)
        log_top.addWidget(self.scan_open_log_btn)

        self.scan_copy_log_btn = QPushButton("Copy Log")
        self.scan_copy_log_btn.setToolTip("Copy the complete scanner log currently shown here to the clipboard.")
        self.scan_copy_log_btn.clicked.connect(self._copy_scan_log)
        log_top.addWidget(self.scan_copy_log_btn)

        self.scan_expand_log_btn = QPushButton("Expand Log")
        self.scan_expand_log_btn.setToolTip("Open the current scanner log in a larger, resizable window.")
        self.scan_expand_log_btn.clicked.connect(self._open_scan_log_window)
        log_top.addWidget(self.scan_expand_log_btn)

        self.scan_open_preview_btn = QPushButton("Open Routing Preview")
        self.scan_open_preview_btn.setEnabled(False)
        self.scan_open_preview_btn.setToolTip("Open the saved Routing_Preview.txt from the most recent dry run.")
        self.scan_open_preview_btn.clicked.connect(self._open_last_routing_preview)
        log_top.addWidget(self.scan_open_preview_btn)

        self.scan_recover_routing_btn = QPushButton("Recover Last Routing")
        self.scan_recover_routing_btn.setEnabled(False)
        self.scan_recover_routing_btn.setToolTip(
            "Move files from the latest routing run back to their recorded original paths after SHA-256 verification."
        )
        self.scan_recover_routing_btn.clicked.connect(self._recover_last_routing)
        log_top.addWidget(self.scan_recover_routing_btn)

        log_layout.addLayout(log_top)
        self.scan_output_tabs = QTabWidget()
        self.scan_output_tabs.setObjectName("ScanOutputTabs")

        log_page = QWidget()
        log_page.setObjectName("ScanOutputPage")
        log_page_layout = QVBoxLayout(log_page)
        log_page_layout.setContentsMargins(0, 0, 0, 0)
        self.scan_log = QPlainTextEdit()
        self.scan_log.setObjectName("ScanLog")
        self.scan_log.setReadOnly(True)
        if IS_PUBLIC_BUILD:
            ready_text = (
                "CAS CURATOR - CONSERVATIVE CAS SCANNER\n\n"
                "This release reads direct CASP metadata and GEOM face counts. Filenames are never used to classify content.\n"
                "Run Scan generates CASP IDs, supported CC Type and Highest Polys, extracts dedicated THUM previews when available, and records package hash/scan time.\n"
                "Whether a package contains or overlaps with EA resources is not assessed. Skin type, general IMG texture purpose, script/gameplay and Health/Risk are not supported scan results.\n"
                "Run Scan creates a mandatory MDB backup before scanner writes.\n"
                "Preview Routing before any move. Run Routing requires explicit acknowledgement and creates a recovery manifest.\n\n"
                "Historical scanner-derived MDB classifications remain withheld unless independently validated."
            )
        else:
            ready_text = (
                "SCANNER READY\n\n"
                "Core: DEV/PUBLIC SCRIPTCLASS TYPEDEF MATCH PERMANENT - 2026-09-12\n"
                "Start Scan updates the selected MDB but never moves source files.\n"
                "Routing Preview performs the same health/classification scan, calculates the exact route destinations, "
                "moves nothing, and skips MDB write-back.\n\n"
                "Run Routing physically moves source files. Preview first and review the destinations before moving a large library."
            )
        self.scan_log.setPlainText(ready_text)
        log_page_layout.addWidget(self.scan_log)
        self.scan_output_tabs.addTab(log_page, "Scanner log")

        preview_page = QWidget()
        preview_page.setObjectName("ScanOutputPage")
        preview_page_layout = QVBoxLayout(preview_page)
        preview_page_layout.setContentsMargins(0, 0, 0, 0)
        self.routing_preview_table = QTableWidget(0, 5)
        self.routing_preview_table.setObjectName("RoutingPreviewTable")
        self.routing_preview_table.setHorizontalHeaderLabels(["File", "Category", "Duplicate", "Destination (inside sort folder)", "Status"])
        self.routing_preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.routing_preview_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.routing_preview_table.setAlternatingRowColors(False)
        self.routing_preview_table.verticalHeader().setVisible(False)
        self.routing_preview_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.routing_preview_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.routing_preview_table.customContextMenuRequested.connect(self._show_routing_preview_context_menu)
        self.routing_preview_table.cellDoubleClicked.connect(self._copy_routing_destination_from_cell)
        header = self.routing_preview_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.routing_preview_table.setColumnWidth(0, 420)
        self.routing_preview_table.setColumnWidth(3, 620)
        self.routing_preview_table.setToolTip(
            "Destination shows the route relative to the selected sort folder. Hover to see the full path; double-click or right-click to copy the full path."
        )
        preview_page_layout.addWidget(self.routing_preview_table)
        self.scan_output_tabs.addTab(preview_page, "Routing preview")

        self.scan_output_tabs.setMinimumHeight(300)
        log_layout.addWidget(self.scan_output_tabs, 1)
        log_card.setMinimumHeight(360)
        layout.addWidget(log_card, 1)

        workflow_hint = QLabel("Recommended workflow:  1. Run Scan  →  2. Preview Routing  →  3. Run Routing")
        workflow_hint.setObjectName("ScanHint")
        workflow_hint.setWordWrap(True)
        workflow_hint.setToolTip("Run Scan adds or refreshes MDB information. Preview Routing shows destinations without moving files. Run Routing moves files after confirmation.")
        layout.addWidget(workflow_hint)

        button_row = QHBoxLayout()
        button_row.setSpacing(8)
        button_row.addStretch(1)
        self.scan_cancel_btn = QPushButton("Cancel")
        self.scan_cancel_btn.setEnabled(False)
        self.scan_cancel_btn.setToolTip("Cancellation is disabled while the scanner may be writing the MDB, to avoid interrupting a save.")
        self.routing_preview_btn = QPushButton("Preview Routing" if IS_PUBLIC_BUILD else "Preview Routing")
        self.routing_preview_btn.setEnabled(False)
        self.routing_preview_btn.setToolTip(
            "Calculates where every supported file would be routed. Source files, destination folder and MDB remain unchanged."
        )
        self.routing_preview_btn.clicked.connect(self._start_routing_preview)
        self.scan_start_btn = QPushButton("Run Scan" if IS_PUBLIC_BUILD else "Start Scan")
        self.scan_start_btn.setObjectName("ScanStartButton")
        self.scan_start_btn.setEnabled(False)
        self.scan_start_btn.setToolTip(("Scans the selected CAS CC source folder and updates the selected MDB. Source files are not moved." if IS_PUBLIC_BUILD else "Runs the DEV scanner core against the selected source folder and MDB. Source files are not moved."))
        self.scan_start_btn.clicked.connect(self._start_scan_test)

        self.real_routing_check = QPushButton("☐  I understand: Run Routing will MOVE source files" if IS_PUBLIC_BUILD else "☐  I understand: routing MOVES source files")
        self.real_routing_check.setObjectName("RoutingAckButton")
        self.real_routing_check.setCheckable(True)
        self.real_routing_check.setMinimumHeight(34)
        ack_width = QFontMetrics(self.real_routing_check.font()).horizontalAdvance("☑  I understand: Run Routing will MOVE source files") + 34
        self.real_routing_check.setMinimumWidth(max(380, ack_width))
        self.real_routing_check.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.real_routing_check.setToolTip(
            "Required safety acknowledgement. Run Routing physically moves files from the source folder into the selected sort destination."
        )
        self.real_routing_check.toggled.connect(self._on_routing_ack_toggled)
        self.real_routing_btn = QPushButton("Run Routing")
        self.real_routing_btn.setEnabled(False)
        self.real_routing_btn.setToolTip(
            "Physically moves files into the selected sort destination and updates the selected MDB. Preview first when changing routing setup."
        )
        self.real_routing_btn.clicked.connect(self._start_real_routing_test)

        button_row.addWidget(self.scan_cancel_btn)
        button_row.addWidget(self.scan_start_btn)
        button_row.addWidget(self.routing_preview_btn)
        button_row.addWidget(self.real_routing_check, 1)
        button_row.addWidget(self.real_routing_btn)
        layout.addLayout(button_row)

        scroll.setWidget(page)
        return scroll

    def _build_random_sim_page(self):
        """DEV-only Random Sim Studio: generator, probability editor and exporter."""
        scroll = QScrollArea()
        scroll.setObjectName("RandomSimScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        page = QWidget()
        page.setObjectName("RandomSimPage")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        title_row = QHBoxLayout()
        title = QLabel("Random Sim Studio — DEV experiment")
        title.setObjectName("EditTitle")
        title_row.addWidget(title)
        title_row.addStretch(1)
        guide_btn = QPushButton("Random Sim guide…")
        guide_btn.clicked.connect(self._random_sim_show_probability_guide)
        title_row.addWidget(guide_btn)
        layout.addLayout(title_row)

        intro = QLabel(
            "Uses the currently opened Master Database and only scanner-validated CAS classifications. "
            "Choose category counts/chances, optionally tune type and individual-item probability, generate a set, lock pieces, and reroll. "
            "Collection Exporter copies package files only. Scanner logic and the MDB are never changed by Random Sim Studio. "
            "Exported Sim usage is tracked in a separate DEV sidecar index."
        )
        intro.setObjectName("EditHelp")
        intro.setWordWrap(True)
        layout.addWidget(intro)

        self.random_sim_status = QLabel("Open a Master Database in Library first.")
        self.random_sim_status.setObjectName("StatsStrong")
        self.random_sim_status.setWordWrap(True)
        layout.addWidget(self.random_sim_status)

        config_card = QFrame()
        config_card.setObjectName("EditCard")
        config_layout = QVBoxLayout(config_card)
        config_layout.setContentsMargins(9, 9, 9, 9)
        config_layout.setSpacing(6)

        config_heading = QLabel("Category setup")
        config_heading.setObjectName("EditTitle")
        config_layout.addWidget(config_heading)

        config_help = QLabel(
            "Count 0 disables a category. Count 1 chooses one package; Count 2 allows two. "
            "Chance % controls whether an unlocked category participates in a generation pass. Locked pieces are always kept."
        )
        config_help.setWordWrap(True)
        config_help.setObjectName("EditHelp")
        config_layout.addWidget(config_help)

        self.random_sim_category_table = QTableWidget(0, 4)
        self.random_sim_category_table.setObjectName("RandomSimCategoryTable")
        self.random_sim_category_table.setHorizontalHeaderLabels(["Category", "Available", "Count", "Chance %"])
        self.random_sim_category_table.verticalHeader().setVisible(False)
        self.random_sim_category_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.random_sim_category_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.random_sim_category_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.random_sim_category_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.random_sim_category_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.random_sim_category_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.random_sim_category_table.setMinimumHeight(180)
        self.random_sim_category_table.setMaximumHeight(320)
        config_layout.addWidget(self.random_sim_category_table)

        config_buttons = QHBoxLayout()
        self.random_sim_refresh_btn = QPushButton("Refresh from MDB")
        self.random_sim_refresh_btn.clicked.connect(self._refresh_random_sim_data)
        self.random_sim_reset_exclusions_btn = QPushButton("Reset exclusions…")
        self.random_sim_reset_exclusions_btn.clicked.connect(self._random_sim_reset_exclusions)
        config_buttons.addWidget(self.random_sim_refresh_btn)
        config_buttons.addWidget(self.random_sim_reset_exclusions_btn)
        config_buttons.addStretch(1)
        config_layout.addLayout(config_buttons)
        layout.addWidget(config_card)

        probability_card = QFrame()
        probability_card.setObjectName("EditCard")
        probability_layout = QVBoxLayout(probability_card)
        probability_layout.setContentsMargins(9, 9, 9, 9)
        probability_layout.setSpacing(6)

        probability_heading = QLabel("Probability & item preview — DEV phase 4")
        probability_heading.setObjectName("EditTitle")
        probability_layout.addWidget(probability_heading)

        probability_help = QLabel(
            "Leave rules at Weight 1 for normal random selection. Weight 0 effectively prevents a type/item; larger weights make it more common. "
            "Exact % reserves that literal share of a selection draw. Exact values in the same group may total at most 100%; any remaining share goes to Weight-mode options."
        )
        probability_help.setObjectName("EditHelp")
        probability_help.setWordWrap(True)
        probability_layout.addWidget(probability_help)

        prob_filters = QHBoxLayout()
        prob_filters.addWidget(QLabel("Category:"))
        self.random_sim_probability_category_combo = QComboBox()
        self.random_sim_probability_category_combo.setMinimumWidth(170)
        self.random_sim_probability_category_combo.currentIndexChanged.connect(self._random_sim_probability_category_changed)
        prob_filters.addWidget(self.random_sim_probability_category_combo)
        prob_filters.addWidget(QLabel("Item subtype:"))
        self.random_sim_probability_subtype_combo = QComboBox()
        self.random_sim_probability_subtype_combo.setMinimumWidth(170)
        self.random_sim_probability_subtype_combo.currentIndexChanged.connect(self._random_sim_probability_subtype_changed)
        prob_filters.addWidget(self.random_sim_probability_subtype_combo)
        reset_prob_btn = QPushButton("Reset this category…")
        reset_prob_btn.clicked.connect(self._random_sim_reset_probability_category)
        prob_filters.addWidget(reset_prob_btn)
        prob_filters.addStretch(1)
        probability_layout.addLayout(prob_filters)

        probability_splitter = QSplitter(Qt.Horizontal)
        probability_splitter.setChildrenCollapsible(False)

        editor_side = QWidget()
        editor_layout = QVBoxLayout(editor_side)
        editor_layout.setContentsMargins(0, 0, 0, 0)
        editor_layout.setSpacing(6)

        type_label = QLabel("Type / subtype rules")
        type_label.setObjectName("StatsStrong")
        editor_layout.addWidget(type_label)
        self.random_sim_type_probability_table = QTableWidget(0, 4)
        self.random_sim_type_probability_table.setHorizontalHeaderLabels(["Type / subtype", "Available", "Mode", "Value"])
        self.random_sim_type_probability_table.verticalHeader().setVisible(False)
        self.random_sim_type_probability_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.random_sim_type_probability_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.random_sim_type_probability_table.setSelectionMode(QTableWidget.SingleSelection)
        self.random_sim_type_probability_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.random_sim_type_probability_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.random_sim_type_probability_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.random_sim_type_probability_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.random_sim_type_probability_table.setMinimumHeight(130)
        self.random_sim_type_probability_table.setMaximumHeight(220)
        self.random_sim_type_probability_table.itemSelectionChanged.connect(self._random_sim_type_rule_selected)
        editor_layout.addWidget(self.random_sim_type_probability_table)

        item_label = QLabel("Individual CC rules")
        item_label.setObjectName("StatsStrong")
        editor_layout.addWidget(item_label)
        self.random_sim_item_probability_table = QTableWidget(0, 6)
        self.random_sim_item_probability_table.setHorizontalHeaderLabels(["Package", "Subtype", "Mode", "Value", "Health", "Risk"])
        self.random_sim_item_probability_table.verticalHeader().setVisible(False)
        self.random_sim_item_probability_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.random_sim_item_probability_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.random_sim_item_probability_table.setSelectionMode(QTableWidget.SingleSelection)
        self.random_sim_item_probability_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.random_sim_item_probability_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.random_sim_item_probability_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.random_sim_item_probability_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.random_sim_item_probability_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.random_sim_item_probability_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.random_sim_item_probability_table.setMinimumHeight(220)
        self.random_sim_item_probability_table.itemSelectionChanged.connect(self._random_sim_preview_item_selection)
        editor_layout.addWidget(self.random_sim_item_probability_table, 1)
        probability_splitter.addWidget(editor_side)

        preview_side = QFrame()
        preview_side.setObjectName("ThumbnailCard")
        preview_layout = QVBoxLayout(preview_side)
        preview_layout.setContentsMargins(8, 8, 8, 8)
        preview_layout.setSpacing(6)
        preview_title = QLabel("Selected CC preview")
        preview_title.setObjectName("StatsStrong")
        preview_layout.addWidget(preview_title)
        self.random_sim_preview_label = ResponsiveThumbnailLabel("Select an item to preview")
        self.random_sim_preview_label.setObjectName("ThumbnailImage")
        self.random_sim_preview_label.setAlignment(Qt.AlignCenter)
        self.random_sim_preview_label.setMinimumSize(260, 180)
        self.random_sim_preview_label.setMaximumHeight(360)
        self.random_sim_preview_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        preview_layout.addWidget(self.random_sim_preview_label, 1)
        self.random_sim_preview_info = QTextBrowser()
        self.random_sim_preview_info.setOpenExternalLinks(False)
        self.random_sim_preview_info.setMinimumWidth(280)
        self.random_sim_preview_info.setMaximumHeight(210)
        preview_layout.addWidget(self.random_sim_preview_info)
        probability_splitter.addWidget(preview_side)
        probability_splitter.setStretchFactor(0, 3)
        probability_splitter.setStretchFactor(1, 2)
        probability_splitter.setSizes([850, 420])
        probability_layout.addWidget(probability_splitter, 1)
        layout.addWidget(probability_card)

        result_card = QFrame()
        result_card.setObjectName("EditCard")
        result_layout = QVBoxLayout(result_card)
        result_layout.setContentsMargins(9, 9, 9, 9)
        result_layout.setSpacing(6)

        result_heading = QLabel("Generated CC set")
        result_heading.setObjectName("EditTitle")
        result_layout.addWidget(result_heading)

        self.random_sim_result_table = QTableWidget(0, 6)
        self.random_sim_result_table.setObjectName("RandomSimResultTable")
        self.random_sim_result_table.setHorizontalHeaderLabels(
            ["Lock", "Category", "Package", "Health", "Risk", "Source"]
        )
        self.random_sim_result_table.verticalHeader().setVisible(False)
        self.random_sim_result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.random_sim_result_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.random_sim_result_table.setSelectionMode(QTableWidget.SingleSelection)
        self.random_sim_result_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.random_sim_result_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.random_sim_result_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.random_sim_result_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.random_sim_result_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.random_sim_result_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.random_sim_result_table.setMinimumHeight(230)
        self.random_sim_result_table.itemSelectionChanged.connect(self._random_sim_preview_generated_selection)
        result_layout.addWidget(self.random_sim_result_table, 1)

        result_buttons = QHBoxLayout()
        self.random_sim_generate_btn = QPushButton("Generate / reroll unlocked")
        self.random_sim_generate_btn.setObjectName("SaveButton")
        self.random_sim_generate_btn.clicked.connect(self._random_sim_generate)
        self.random_sim_reroll_btn = QPushButton("Reroll selected")
        self.random_sim_reroll_btn.clicked.connect(self._random_sim_reroll_selected)
        self.random_sim_exclude_btn = QPushButton("Exclude selected")
        self.random_sim_exclude_btn.clicked.connect(self._random_sim_exclude_selected)
        self.random_sim_clear_btn = QPushButton("Clear set")
        self.random_sim_clear_btn.clicked.connect(self._random_sim_clear_selection)
        result_buttons.addWidget(self.random_sim_generate_btn)
        result_buttons.addWidget(self.random_sim_reroll_btn)
        result_buttons.addWidget(self.random_sim_exclude_btn)
        result_buttons.addStretch(1)
        result_buttons.addWidget(self.random_sim_clear_btn)
        result_layout.addLayout(result_buttons)

        phase_note = QLabel(
            "Locks override probability so a piece you lock is never discarded by a chance roll. Exclusion still overrides every probability setting. "
            "The selected item preview above also works for rows in this generated set."
        )
        phase_note.setObjectName("EditHelp")
        phase_note.setWordWrap(True)
        result_layout.addWidget(phase_note)
        layout.addWidget(result_card, 1)

        export_card = QFrame()
        export_card.setObjectName("EditCard")
        export_layout = QVBoxLayout(export_card)
        export_layout.setContentsMargins(9, 9, 9, 9)
        export_layout.setSpacing(6)

        export_heading = QLabel("Collection Exporter — DEV phase 2")
        export_heading.setObjectName("EditTitle")
        export_layout.addWidget(export_heading)

        export_help = QLabel(
            "Copies the generated .package files into a new timestamped folder and writes JSON + TXT manifests. "
            "The exporter de-duplicates identical package hashes, safely renames filename collisions, verifies recorded SHA256 values, "
            "and only follows dependencies that are explicitly declared and exactly matched in the open MDB. It never guesses dependencies. "
            "The JSON manifest also records the active Random Sim probability settings for reproducibility. "
            "After a successful export, original CC records gain separate Random Sim usage links; existing Status values such as Used in Game stay unchanged."
        )
        export_help.setObjectName("EditHelp")
        export_help.setWordWrap(True)
        export_layout.addWidget(export_help)

        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("Set name:"))
        self.random_sim_set_name_edit = QLineEdit("Random Sim")
        self.random_sim_set_name_edit.setPlaceholderText("Random Sim")
        name_row.addWidget(self.random_sim_set_name_edit, 1)
        export_layout.addLayout(name_row)

        root_row = QHBoxLayout()
        root_row.addWidget(QLabel("Export root:"))
        self.random_sim_export_root_edit = QLineEdit(self.random_sim_export_root)
        self.random_sim_export_root_edit.setReadOnly(True)
        self.random_sim_export_root_edit.setPlaceholderText("Choose a folder for generated Sim collections…")
        root_row.addWidget(self.random_sim_export_root_edit, 1)
        self.random_sim_export_browse_btn = QPushButton("Browse…")
        self.random_sim_export_browse_btn.clicked.connect(self._random_sim_choose_export_root)
        root_row.addWidget(self.random_sim_export_browse_btn)
        export_layout.addLayout(root_row)

        export_buttons = QHBoxLayout()
        self.random_sim_export_btn = QPushButton("Export generated set")
        self.random_sim_export_btn.setObjectName("SaveButton")
        self.random_sim_export_btn.setEnabled(False)
        self.random_sim_export_btn.clicked.connect(self._random_sim_export_collection)
        self.random_sim_open_export_btn = QPushButton("Open last export")
        self.random_sim_open_export_btn.setEnabled(False)
        self.random_sim_open_export_btn.clicked.connect(self._random_sim_open_last_export)
        export_buttons.addWidget(self.random_sim_export_btn)
        export_buttons.addWidget(self.random_sim_open_export_btn)
        export_buttons.addStretch(1)
        export_layout.addLayout(export_buttons)

        future_note = QLabel(
            "Successful exports are saved into the DEV Random Sim history below. "
            "Playable .sim creation remains a later research experiment."
        )
        future_note.setObjectName("EditHelp")
        future_note.setWordWrap(True)
        export_layout.addWidget(future_note)
        layout.addWidget(export_card)

        history_card = QFrame()
        history_card.setObjectName("EditCard")
        history_layout = QVBoxLayout(history_card)
        history_layout.setContentsMargins(9, 9, 9, 9)
        history_layout.setSpacing(6)

        history_heading_row = QHBoxLayout()
        history_heading = QLabel("Random Sim history & CC gallery — DEV phase 5")
        history_heading.setObjectName("EditTitle")
        history_heading_row.addWidget(history_heading)
        history_heading_row.addStretch(1)
        self.random_sim_history_refresh_btn = QPushButton("Refresh history")
        self.random_sim_history_refresh_btn.clicked.connect(self._random_sim_refresh_history)
        history_heading_row.addWidget(self.random_sim_history_refresh_btn)
        history_layout.addLayout(history_heading_row)

        history_help = QLabel(
            "Every successfully exported Random Sim gets a persistent profile. Select one on the left to see all directly selected CC as a visual gallery. "
            "Each gallery card links back to the original Library record in the currently opened MDB when that package can still be matched. Explicit dependency-only packages remain separate from selected CC."
        )
        history_help.setObjectName("EditHelp")
        history_help.setWordWrap(True)
        history_layout.addWidget(history_help)

        history_splitter = QSplitter(Qt.Horizontal)
        history_splitter.setChildrenCollapsible(False)

        history_list_side = QWidget()
        history_list_layout = QVBoxLayout(history_list_side)
        history_list_layout.setContentsMargins(0, 0, 0, 0)
        history_list_layout.setSpacing(6)
        self.random_sim_history_table = QTableWidget(0, 4)
        self.random_sim_history_table.setHorizontalHeaderLabels(["Sim", "Created", "Selected CC", "Dependencies"])
        self.random_sim_history_table.verticalHeader().setVisible(False)
        self.random_sim_history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.random_sim_history_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.random_sim_history_table.setSelectionMode(QTableWidget.SingleSelection)
        self.random_sim_history_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.random_sim_history_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.random_sim_history_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.random_sim_history_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.random_sim_history_table.setMinimumWidth(420)
        self.random_sim_history_table.setMinimumHeight(310)
        self.random_sim_history_table.itemSelectionChanged.connect(self._random_sim_history_selected)
        history_list_layout.addWidget(self.random_sim_history_table, 1)

        history_buttons = QHBoxLayout()
        self.random_sim_history_open_folder_btn = QPushButton("Open Sim folder")
        self.random_sim_history_open_folder_btn.setEnabled(False)
        self.random_sim_history_open_folder_btn.clicked.connect(self._random_sim_history_open_folder)
        self.random_sim_history_open_manifest_btn = QPushButton("Open manifest")
        self.random_sim_history_open_manifest_btn.setEnabled(False)
        self.random_sim_history_open_manifest_btn.clicked.connect(self._random_sim_history_open_manifest)
        history_buttons.addWidget(self.random_sim_history_open_folder_btn)
        history_buttons.addWidget(self.random_sim_history_open_manifest_btn)
        history_buttons.addStretch(1)
        history_list_layout.addLayout(history_buttons)
        history_splitter.addWidget(history_list_side)

        history_detail = QWidget()
        history_detail_layout = QVBoxLayout(history_detail)
        history_detail_layout.setContentsMargins(0, 0, 0, 0)
        history_detail_layout.setSpacing(6)
        self.random_sim_history_profile = QTextBrowser()
        self.random_sim_history_profile.setOpenExternalLinks(False)
        self.random_sim_history_profile.setMaximumHeight(190)
        self.random_sim_history_profile.setHtml("<i>No exported Random Sim selected.</i>")
        history_detail_layout.addWidget(self.random_sim_history_profile)

        self.random_sim_history_gallery_status = QLabel("No exported Random Sim selected.")
        self.random_sim_history_gallery_status.setObjectName("StatsStrong")
        self.random_sim_history_gallery_status.setWordWrap(True)
        history_detail_layout.addWidget(self.random_sim_history_gallery_status)

        self.random_sim_history_gallery_scroll = QScrollArea()
        self.random_sim_history_gallery_scroll.setWidgetResizable(True)
        self.random_sim_history_gallery_scroll.setFrameShape(QFrame.NoFrame)
        self.random_sim_history_gallery_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.random_sim_history_gallery_widget = QWidget()
        self.random_sim_history_gallery_grid = QGridLayout(self.random_sim_history_gallery_widget)
        self.random_sim_history_gallery_grid.setContentsMargins(0, 0, 0, 0)
        self.random_sim_history_gallery_grid.setHorizontalSpacing(8)
        self.random_sim_history_gallery_grid.setVerticalSpacing(8)
        self.random_sim_history_gallery_grid.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.random_sim_history_gallery_scroll.setWidget(self.random_sim_history_gallery_widget)
        self.random_sim_history_gallery_scroll.setMinimumHeight(360)
        history_detail_layout.addWidget(self.random_sim_history_gallery_scroll, 1)
        history_splitter.addWidget(history_detail)
        history_splitter.setStretchFactor(0, 1)
        history_splitter.setStretchFactor(1, 2)
        history_splitter.setSizes([470, 900])
        history_layout.addWidget(history_splitter, 1)
        layout.addWidget(history_card)

        next_note = QLabel(
            "Next planned DEV stage: Color/Theme Studio improvements — separate Lock to Slot vs Keep in Shuffle colour locks, plus persistent custom-theme storage across CAS Curator updates."
        )
        next_note.setObjectName("EditHelp")
        next_note.setWordWrap(True)
        layout.addWidget(next_note)

        scroll.setWidget(page)
        QTimer.singleShot(0, self._random_sim_refresh_history)
        return scroll

    def _random_sim_show_probability_guide(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Random Sim Studio Guide — Probability & Preview")
        dialog.resize(780, 640)
        layout = QVBoxLayout(dialog)
        browser = QTextBrowser()
        browser.setHtml(
            """
            <h2>Random Sim Studio — Probability guide</h2>
            <h3>1. Category count and chance</h3>
            <p><b>Count</b> is the maximum number of pieces selected for that category (0, 1 or 2). "
            "<b>Chance %</b> decides whether an unlocked category participates in that generation pass. "
            "A locked piece is preserved even if the category chance roll fails.</p>
            <h3>2. Weight mode</h3>
            <p>Weight is relative. <b>1</b> is normal, <b>0</b> prevents that type/item from being chosen, "
            "0.5 is half-normal, 2 is twice-normal, 3 is very common, and 5 is a strong favourite. "
            "Only Weight-mode options share the probability that remains after Exact % rules.</p>
            <h3>3. Exact % mode</h3>
            <p>Exact % reserves a literal percentage band of one selection draw. In the same group, exact values may total at most 100%. "
            "For example, if Long Hair is Exact 70% and Short Hair is Exact 20%, the remaining 10% is shared by any types still in Weight mode. "
            "If every option is Exact and the total is below 100%, the unused remainder can deliberately produce no pick.</p>
            <p>Individual-item Exact % works the same way inside the chosen type/subtype. For Count 2, rules are applied again for the second open slot without reusing the first package.</p>
            <h3>4. Exclusion and locks</h3>
            <p><b>Exclude selected</b> is absolute and removes the CC from the random pool regardless of probability. "
            "Locks are also absolute for the current generated set: a locked piece stays while unlocked pieces reroll.</p>
            <h3>5. Preview</h3>
            <p>Select an individual CC row to see its existing thumbnail (when available), CASP type/subtype, CASP instance IDs, Health/Risk, source path, and Random Sim usage count. "
            "Selecting a row in the Generated CC set uses the same preview panel.</p>
            <h3>6. Export, usage links and history</h3>
            <p>A successful Collection Export creates a persistent Random Sim profile. The History &amp; CC Gallery section shows every directly selected CC item, using its saved thumbnail when available and CASP/type metadata as a fallback. "
            "Click a gallery card's package button to jump back to the original Library record in the currently opened MDB. Dependency-only packages are counted separately and are not presented as directly selected Sim CC.</p>
            <p>The original CC record also keeps its separate Random Sim usage links. This does not replace ordinary Library Status such as <b>Used in Game</b>.</p>
            <h3>7. Persistence and safety</h3>
            <p>Probability rules are stored in CAS Curator DEV settings, while Random Sim export/history links are stored in a separate DEV sidecar registry — not in the Master Database. Scanner logic and source packages are never edited. "
            "The Collection Exporter only copies files, and its JSON manifest records the active generation settings.</p>
            """
        )
        layout.addWidget(browser, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _random_sim_persist_probability_rules(self):
        self.settings.setValue(
            "random_sim/category_chances_json",
            json.dumps(self.random_sim_category_chances, ensure_ascii=False, sort_keys=True),
        )
        self.settings.setValue(
            "random_sim/type_rules_json",
            json.dumps(self.random_sim_type_rules, ensure_ascii=False, sort_keys=True),
        )
        self.settings.setValue(
            "random_sim/item_rules_json",
            json.dumps(self.random_sim_item_rules, ensure_ascii=False, sort_keys=True),
        )

    def _random_sim_set_category_chance(self, category, value):
        category = str(category or "")
        value = max(0.0, min(100.0, float(value)))
        if abs(value - 100.0) < 0.0001:
            self.random_sim_category_chances.pop(category, None)
        else:
            self.random_sim_category_chances[category] = value
        self._random_sim_persist_probability_rules()

    def _random_sim_current_category_chances(self):
        chances = {}
        for category, spin in self.random_sim_category_chance_boxes.items():
            chances[category] = float(spin.value())
        for category, value in self.random_sim_category_chances.items():
            if category not in chances:
                try:
                    chances[category] = float(value)
                except (TypeError, ValueError):
                    pass
        return chances

    def _random_sim_rule_widgets(self, table, row, store_name, key):
        store = self.random_sim_type_rules if store_name == "type" else self.random_sim_item_rules
        rule = store.get(key) if isinstance(store.get(key), dict) else {}
        mode_value = str(rule.get("mode", "weight") or "weight").casefold()
        value_default = 0.0 if mode_value == "exact" else 1.0
        try:
            numeric = float(rule.get("value", value_default))
        except (TypeError, ValueError):
            numeric = value_default

        mode = QComboBox()
        mode.addItems(["Weight", "Exact %"])
        mode.setCurrentIndex(1 if mode_value == "exact" else 0)
        mode.setMinimumWidth(92)
        table.setCellWidget(row, 2, mode)

        value = QDoubleSpinBox()
        value.setDecimals(2)
        value.setRange(0.0, 100.0)
        value.setValue(max(0.0, min(100.0, numeric)))
        value.setMinimumWidth(88)
        table.setCellWidget(row, 3, value)

        def commit(*_args):
            selected_mode = "exact" if mode.currentText().startswith("Exact") else "weight"
            selected_value = float(value.value())
            value.setSuffix(" %" if selected_mode == "exact" else "")
            value.setSingleStep(1.0 if selected_mode == "exact" else 0.25)
            if selected_mode == "weight" and abs(selected_value - 1.0) < 0.0001:
                store.pop(key, None)
            else:
                store[key] = {"mode": selected_mode, "value": selected_value}
            self._random_sim_persist_probability_rules()

        mode.currentTextChanged.connect(commit)
        value.valueChanged.connect(commit)
        # Apply display affordances after signals are connected without changing data.
        value.setSuffix(" %" if mode_value == "exact" else "")
        value.setSingleStep(1.0 if mode_value == "exact" else 0.25)

    def _random_sim_refresh_probability_editor(self):
        if not hasattr(self, "random_sim_probability_category_combo"):
            return
        combo = self.random_sim_probability_category_combo
        previous = combo.currentData()
        combo.blockSignals(True)
        combo.clear()
        for category in self.random_sim_grouped:
            combo.addItem(category_label(category), category)
        if previous:
            index = combo.findData(previous)
            if index >= 0:
                combo.setCurrentIndex(index)
        combo.blockSignals(False)
        self._random_sim_probability_category_changed()

    def _random_sim_probability_category_changed(self, *_args):
        if not hasattr(self, "random_sim_type_probability_table"):
            return
        category = self.random_sim_probability_category_combo.currentData()
        records = list(self.random_sim_grouped.get(category, ())) if category else []

        by_subtype = {}
        labels = {}
        for record in records:
            label = subtype_label(record)
            skey = label.casefold()
            labels.setdefault(skey, label)
            by_subtype.setdefault(skey, []).append(record)

        subtype_combo = self.random_sim_probability_subtype_combo
        previous_subtype = subtype_combo.currentData()
        subtype_combo.blockSignals(True)
        subtype_combo.clear()
        subtype_combo.addItem("All subtypes", "")
        for skey in sorted(by_subtype, key=lambda key: labels[key].casefold()):
            subtype_combo.addItem(labels[skey], skey)
        if previous_subtype:
            idx = subtype_combo.findData(previous_subtype)
            if idx >= 0:
                subtype_combo.setCurrentIndex(idx)
        subtype_combo.blockSignals(False)

        table = self.random_sim_type_probability_table
        table.setRowCount(0)
        for skey in sorted(by_subtype, key=lambda key: labels[key].casefold()):
            row = table.rowCount()
            table.insertRow(row)
            item = QTableWidgetItem(labels[skey])
            item.setData(Qt.UserRole, skey)
            table.setItem(row, 0, item)
            available = QTableWidgetItem(str(len(by_subtype[skey])))
            available.setTextAlignment(Qt.AlignCenter)
            table.setItem(row, 1, available)
            rule_key = type_rule_key(str(category or ""), labels[skey])
            self._random_sim_rule_widgets(table, row, "type", rule_key)

        self._random_sim_refresh_probability_item_table()

    def _random_sim_probability_subtype_changed(self, *_args):
        self._random_sim_refresh_probability_item_table()

    def _random_sim_type_rule_selected(self):
        rows = self.random_sim_type_probability_table.selectionModel().selectedRows()
        if not rows:
            return
        item = self.random_sim_type_probability_table.item(rows[0].row(), 0)
        if item is None:
            return
        skey = item.data(Qt.UserRole)
        index = self.random_sim_probability_subtype_combo.findData(skey)
        if index >= 0:
            self.random_sim_probability_subtype_combo.setCurrentIndex(index)

    def _random_sim_refresh_probability_item_table(self):
        if not hasattr(self, "random_sim_item_probability_table"):
            return
        category = self.random_sim_probability_category_combo.currentData()
        subtype_filter = str(self.random_sim_probability_subtype_combo.currentData() or "")
        records = list(self.random_sim_grouped.get(category, ())) if category else []
        if subtype_filter:
            records = [record for record in records if subtype_label(record).casefold() == subtype_filter]

        table = self.random_sim_item_probability_table
        table.setRowCount(0)
        self.random_sim_probability_item_rows = []
        for record in records:
            row = table.rowCount()
            table.insertRow(row)
            package_item = QTableWidgetItem(str(record.get("Package", "") or ""))
            package_item.setData(Qt.UserRole, record_key(record))
            table.setItem(row, 0, package_item)
            table.setItem(row, 1, QTableWidgetItem(subtype_label(record)))
            self._random_sim_rule_widgets(table, row, "item", record_key(record))
            table.setItem(row, 4, QTableWidgetItem(str(record.get("Health Status", "") or "—")))
            table.setItem(row, 5, QTableWidgetItem(str(record.get("Risk", "") or "—")))
            self.random_sim_probability_item_rows.append(record)

        if records:
            table.selectRow(0)
        else:
            self.random_sim_preview_label.clear_thumbnail("No candidate items")
            self.random_sim_preview_info.setHtml("<span>No candidate CC in this filter.</span>")

    def _random_sim_preview_item_selection(self):
        rows = self.random_sim_item_probability_table.selectionModel().selectedRows()
        if not rows:
            return
        row = rows[0].row()
        if 0 <= row < len(getattr(self, "random_sim_probability_item_rows", [])):
            self._random_sim_preview_record(self.random_sim_probability_item_rows[row])

    def _random_sim_preview_generated_selection(self):
        rows = self.random_sim_result_table.selectionModel().selectedRows()
        if not rows:
            return
        row = rows[0].row()
        if 0 <= row < len(self.random_sim_result_rows):
            self._random_sim_preview_record(self.random_sim_result_rows[row]["record"])

    def _random_sim_preview_record(self, record):
        if not record or not hasattr(self, "random_sim_preview_label"):
            return
        effective = dict(record)
        row_num = int(record.get("_ExcelRow", "0") or 0)
        if row_num and row_num in self.pending_changes and "Image" in self.pending_changes[row_num]:
            effective["_ThumbnailTarget"] = self.pending_changes[row_num].get("Image", "")
        target = effective.get("_ThumbnailTarget", "") or effective.get("Image", "")
        path = target_to_local_path(target) if target else None
        if path is not None and path.exists():
            pix = QPixmap(str(path))
            if not pix.isNull():
                self.random_sim_preview_label.set_thumbnail(pix)
            else:
                self.random_sim_preview_label.clear_thumbnail("Thumbnail could not be opened")
        else:
            self.random_sim_preview_label.clear_thumbnail("No thumbnail found")

        usages = selected_usages_for_record(record, self.random_sim_usage_registry)
        fields = [
            ("Package", record.get("Package", "")),
            ("CC Type", record.get("CC Type", "")),
            ("Subtype", subtype_label(record)),
            ("CASP Instance IDs", record.get("CASP Instance IDs", "")),
            ("Health", record.get("Health Status", "")),
            ("Risk", record.get("Risk", "")),
            ("Used by Random Sims", str(len(usages))),
            ("Source", record.get("Original Source Path", "")),
        ]
        lines = []
        for label, value in fields:
            text = str(value or "—")
            lines.append(f"<div style='margin-bottom:5px'><b>{html.escape(label)}:</b> {html.escape(text)}</div>")
        self.random_sim_preview_info.setHtml("".join(lines))

    def _random_sim_reset_probability_category(self):
        category = self.random_sim_probability_category_combo.currentData()
        if not category:
            return
        records = list(self.random_sim_grouped.get(category, ()))
        if not records:
            return
        answer = QMessageBox.question(
            self,
            "Reset Random Sim probability",
            f"Reset category chance, type/subtype rules and individual-item rules for {category_label(category)}?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        self.random_sim_category_chances.pop(category, None)
        prefix = f"{str(category).casefold()}\x1f"
        self.random_sim_type_rules = {
            key: value for key, value in self.random_sim_type_rules.items() if not str(key).startswith(prefix)
        }
        item_keys = {record_key(record) for record in records}
        self.random_sim_item_rules = {
            key: value for key, value in self.random_sim_item_rules.items() if key not in item_keys
        }
        self._random_sim_persist_probability_rules()
        self._refresh_random_sim_data()

    def _random_sim_validate_probability(self):
        errors = validate_probability_rules(
            self.random_sim_grouped,
            self.random_sim_type_rules,
            self.random_sim_item_rules,
        )
        if not errors:
            return True
        preview = "\n".join(f"• {text}" for text in errors[:8])
        if len(errors) > 8:
            preview += f"\n• …and {len(errors) - 8} more"
        QMessageBox.warning(
            self,
            "Random Sim probability conflict",
            "Exact % values in the same selection group cannot total more than 100%.\n\n"
            + preview
            + "\n\nAdjust those rules before generating.",
        )
        return False

    def _random_sim_generation_settings(self):
        return {
            "category_counts": self._random_sim_current_counts(),
            "category_chances": self._random_sim_current_category_chances(),
            "type_rules": dict(self.random_sim_type_rules),
            "item_rules": dict(self.random_sim_item_rules),
            "excluded_keys": sorted(self.random_sim_excluded_keys),
            "semantics": "category chance -> type/subtype exact-or-weight draw -> item exact-or-weight draw; locks override probability",
        }

    def _random_sim_current_counts(self):
        counts = {}
        for category, combo in self.random_sim_count_boxes.items():
            try:
                counts[category] = int(combo.currentText())
            except (TypeError, ValueError):
                counts[category] = 0
        return counts

    def _random_sim_checked_keys(self):
        checked = set()
        for row_index, row_info in enumerate(self.random_sim_result_rows):
            item = self.random_sim_result_table.item(row_index, 0)
            if item is not None and item.checkState() == Qt.Checked:
                key = record_key(row_info["record"])
                if key:
                    checked.add(key)
        return checked

    def _random_sim_locked_map(self, checked_keys=None, skip_row=None, preserve_all_others=False):
        checked_keys = checked_keys if checked_keys is not None else self._random_sim_checked_keys()
        locked = {}
        for row_index, row_info in enumerate(self.random_sim_result_rows):
            if row_index == skip_row:
                continue
            key = record_key(row_info["record"])
            if not key:
                continue
            if preserve_all_others or key in checked_keys:
                locked.setdefault(row_info["category"], []).append(key)
        return locked

    def _refresh_random_sim_data(self):
        if not IS_DEV_BUILD or not hasattr(self, "random_sim_category_table"):
            return

        previous_counts = self._random_sim_current_counts()
        previous_chances = self._random_sim_current_category_chances()
        self.random_sim_count_boxes = {}
        self.random_sim_category_chance_boxes = {}
        self.random_sim_category_table.setRowCount(0)

        if not self.mdb:
            self.random_sim_grouped = {}
            self.random_sim_status.setText("Open a Master Database in Library first.")
            self.random_sim_generate_btn.setEnabled(False)
            self.random_sim_reroll_btn.setEnabled(False)
            self.random_sim_exclude_btn.setEnabled(False)
            self._random_sim_refresh_probability_editor()
            return

        self.random_sim_grouped = group_eligible_records(
            self.mdb.records, self.random_sim_excluded_keys
        )
        eligible_total = sum(len(rows) for rows in self.random_sim_grouped.values())
        self.random_sim_status.setText(
            f"Eligible CAS packages: {eligible_total:,} across {len(self.random_sim_grouped):,} categories | "
            f"Excluded from random: {len(self.random_sim_excluded_keys):,}"
        )

        for row, (category, records) in enumerate(self.random_sim_grouped.items()):
            self.random_sim_category_table.insertRow(row)
            name_item = QTableWidgetItem(category_label(category))
            name_item.setData(Qt.UserRole, category)
            self.random_sim_category_table.setItem(row, 0, name_item)
            available_item = QTableWidgetItem(str(len(records)))
            available_item.setTextAlignment(Qt.AlignCenter)
            self.random_sim_category_table.setItem(row, 1, available_item)

            combo = QComboBox()
            combo.addItems(["0", "1", "2"])
            wanted = previous_counts.get(category, default_count_for_category(category))
            combo.setCurrentText(str(max(0, min(2, int(wanted)))))
            combo.setMinimumWidth(58)
            self.random_sim_category_table.setCellWidget(row, 2, combo)
            self.random_sim_count_boxes[category] = combo

            chance_spin = QDoubleSpinBox()
            chance_spin.setDecimals(0)
            chance_spin.setRange(0.0, 100.0)
            chance_spin.setSuffix(" %")
            chance_spin.setSingleStep(5.0)
            chance_spin.setMinimumWidth(82)
            try:
                chance_value = float(previous_chances.get(category, self.random_sim_category_chances.get(category, 100.0)))
            except (TypeError, ValueError):
                chance_value = 100.0
            chance_spin.setValue(max(0.0, min(100.0, chance_value)))
            chance_spin.valueChanged.connect(
                lambda value, cat=category: self._random_sim_set_category_chance(cat, value)
            )
            self.random_sim_category_table.setCellWidget(row, 3, chance_spin)
            self.random_sim_category_chance_boxes[category] = chance_spin

        self._random_sim_refresh_probability_editor()

        enabled = bool(self.random_sim_grouped)
        self.random_sim_generate_btn.setEnabled(enabled)
        self.random_sim_reroll_btn.setEnabled(enabled)
        self.random_sim_exclude_btn.setEnabled(enabled)

        valid_keys = {
            record_key(rec)
            for records in self.random_sim_grouped.values()
            for rec in records
        }
        filtered = {}
        for category, records in self.random_sim_selection.items():
            kept = [rec for rec in records if record_key(rec) in valid_keys]
            if kept:
                filtered[category] = kept
        self.random_sim_selection = filtered
        self._render_random_sim_selection()

    def _random_sim_generate(self):
        if not self.mdb:
            QMessageBox.information(self, "Random Sim Studio", "Open a Master Database in Library first.")
            return
        if not self.random_sim_grouped:
            self._refresh_random_sim_data()
        if not self.random_sim_grouped:
            QMessageBox.information(self, "Random Sim Studio", "No eligible validated CAS packages are available.")
            return
        if not self._random_sim_validate_probability():
            return
        checked = self._random_sim_checked_keys()
        locked = self._random_sim_locked_map(checked_keys=checked)
        self.random_sim_selection = generate_selection(
            self.random_sim_grouped,
            self._random_sim_current_counts(),
            locked,
            category_chances=self._random_sim_current_category_chances(),
            type_rules=self.random_sim_type_rules,
            item_rules=self.random_sim_item_rules,
        )
        self._render_random_sim_selection(checked)

    def _random_sim_reroll_selected(self):
        rows = self.random_sim_result_table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.information(self, "Random Sim Studio", "Select one generated package to reroll.")
            return
        row_index = rows[0].row()
        if row_index < 0 or row_index >= len(self.random_sim_result_rows):
            return

        target = self.random_sim_result_rows[row_index]
        target_key = record_key(target["record"])
        checked = self._random_sim_checked_keys()
        checked.discard(target_key)
        locked = self._random_sim_locked_map(
            checked_keys=checked,
            skip_row=row_index,
            preserve_all_others=True,
        )
        if not self._random_sim_validate_probability():
            return
        reroll_chances = self._random_sim_current_category_chances()
        # Reroll selected is an explicit request for a replacement in this already-active
        # category, so do not let the category-level inclusion chance blank the slot.
        reroll_chances[target["category"]] = 100.0
        self.random_sim_selection = generate_selection(
            self.random_sim_grouped,
            self._random_sim_current_counts(),
            locked,
            {target["category"]: [target_key]} if target_key else {},
            category_chances=reroll_chances,
            type_rules=self.random_sim_type_rules,
            item_rules=self.random_sim_item_rules,
        )
        self._render_random_sim_selection(checked)

    def _random_sim_exclude_selected(self):
        rows = self.random_sim_result_table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.information(self, "Random Sim Studio", "Select a generated package to exclude.")
            return
        row_index = rows[0].row()
        if row_index < 0 or row_index >= len(self.random_sim_result_rows):
            return
        record = self.random_sim_result_rows[row_index]["record"]
        key = record_key(record)
        if not key:
            return
        self.random_sim_excluded_keys.add(key)
        self.settings.setValue("random_sim/excluded_keys", sorted(self.random_sim_excluded_keys))
        self._refresh_random_sim_data()
        self.statusBar().showMessage(
            f"Excluded from Random Sim Studio: {record.get('Package', 'package')}", 5000
        )

    def _random_sim_reset_exclusions(self):
        if not self.random_sim_excluded_keys:
            QMessageBox.information(self, "Random Sim Studio", "There are no Random Sim exclusions to reset.")
            return
        answer = QMessageBox.question(
            self,
            "Reset Random Sim exclusions",
            f"Allow all {len(self.random_sim_excluded_keys):,} excluded package(s) to be chosen again?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        self.random_sim_excluded_keys.clear()
        self.settings.remove("random_sim/excluded_keys")
        self._refresh_random_sim_data()

    def _random_sim_clear_selection(self):
        self.random_sim_selection = {}
        self.random_sim_result_rows = []
        self.random_sim_result_table.setRowCount(0)
        if hasattr(self, "random_sim_export_btn"):
            self.random_sim_export_btn.setEnabled(False)
        if self.mdb:
            eligible_total = sum(len(rows) for rows in self.random_sim_grouped.values())
            self.random_sim_status.setText(
                f"Eligible CAS packages: {eligible_total:,} across {len(self.random_sim_grouped):,} categories | "
                f"Excluded from random: {len(self.random_sim_excluded_keys):,}"
            )

    def _render_random_sim_selection(self, checked_keys=None):
        checked_keys = set(checked_keys or ())
        self.random_sim_result_rows = []
        self.random_sim_result_table.setRowCount(0)

        category_order = list(self.random_sim_grouped)
        category_order.extend(
            category for category in self.random_sim_selection if category not in category_order
        )
        for category in category_order:
            for record in self.random_sim_selection.get(category, []):
                row = self.random_sim_result_table.rowCount()
                self.random_sim_result_table.insertRow(row)

                lock_item = QTableWidgetItem()
                lock_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                lock_item.setCheckState(
                    Qt.Checked if record_key(record) in checked_keys else Qt.Unchecked
                )
                self.random_sim_result_table.setItem(row, 0, lock_item)
                self.random_sim_result_table.setItem(row, 1, QTableWidgetItem(category_label(category)))
                self.random_sim_result_table.setItem(row, 2, QTableWidgetItem(str(record.get("Package", ""))))
                self.random_sim_result_table.setItem(row, 3, QTableWidgetItem(str(record.get("Health Status", "") or "—")))
                self.random_sim_result_table.setItem(row, 4, QTableWidgetItem(str(record.get("Risk", "") or "—")))
                self.random_sim_result_table.setItem(row, 5, QTableWidgetItem(str(record.get("Original Source Path", "") or "—")))
                self.random_sim_result_rows.append({"category": category, "record": record})

        if self.mdb:
            eligible_total = sum(len(rows) for rows in self.random_sim_grouped.values())
            selected_total = len(self.random_sim_result_rows)
            self.random_sim_status.setText(
                f"Generated pieces: {selected_total:,} | Eligible CAS packages: {eligible_total:,} | "
                f"Excluded from random: {len(self.random_sim_excluded_keys):,}"
            )
        if hasattr(self, "random_sim_export_btn"):
            self.random_sim_export_btn.setEnabled(bool(self.random_sim_result_rows))

    def _random_sim_choose_export_root(self):
        start = self.random_sim_export_root or str(self.settings.value("last_dir", str(Path.home())))
        folder = QFileDialog.getExistingDirectory(self, "Choose Random Sim collection export folder", start)
        if not folder:
            return False
        self.random_sim_export_root = str(Path(folder).resolve())
        self.settings.setValue("random_sim/export_root", self.random_sim_export_root)
        self.random_sim_export_root_edit.setText(self.random_sim_export_root)
        return True

    def _random_sim_export_collection(self):
        if not self.mdb or not self.random_sim_result_rows:
            QMessageBox.information(self, "Collection Exporter", "Generate a Random Sim CC set first.")
            return
        if not self.random_sim_export_root:
            if not self._random_sim_choose_export_root():
                return

        plan = build_export_plan(self.random_sim_result_rows, self.mdb.records)
        warnings = list(plan.get("warnings", []) or [])
        if warnings:
            preview = []
            for warning in warnings[:6]:
                preview.append(
                    f"• {warning.get('package', '')}: {warning.get('reference', '')} — {warning.get('reason', '')}"
                )
            if len(warnings) > 6:
                preview.append(f"• …and {len(warnings) - 6} more (all will be written to the manifest).")
            answer = QMessageBox.question(
                self,
                "Unresolved dependency notes",
                "Some explicitly declared dependencies could not be auto-copied:\n\n"
                + "\n".join(preview)
                + "\n\nExport the resolved collection anyway?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return

        try:
            folder, manifest = export_collection(
                plan,
                self.random_sim_export_root,
                self.random_sim_set_name_edit.text(),
                VERSION,
                self.mdb.path,
                generation_settings=self._random_sim_generation_settings(),
            )
        except ExportPreflightError as exc:
            details = "\n".join(f"• {problem}" for problem in exc.problems[:10])
            if len(exc.problems) > 10:
                details += f"\n• …and {len(exc.problems) - 10} more."
            QMessageBox.warning(
                self,
                "Collection export blocked",
                "No collection was committed because the preflight check failed:\n\n" + details,
            )
            return
        except Exception as exc:
            QMessageBox.critical(self, "Collection export failed", f"The collection was not completed.\n\n{exc}")
            return

        self.random_sim_last_export = folder
        self.random_sim_open_export_btn.setEnabled(True)
        try:
            self.random_sim_usage_registry = register_random_sim_export(
                self.random_sim_usage_registry, manifest, folder
            )
            save_usage_registry(self.random_sim_usage_registry, self.random_sim_usage_registry_path)
            self._random_sim_refresh_history()
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Collection exported — usage link not saved",
                "The CC collection was exported successfully, but CAS Curator could not save the "
                f"Random Sim usage index. The MDB and source CC were not changed.\n\n{exc}",
            )
        if self.current_source_row is not None:
            current = self.model.record(self.current_source_row)
            if current:
                self.show_record(current)
        self.statusBar().showMessage(f"Random Sim collection exported: {folder}", 8000)
        QMessageBox.information(
            self,
            "Collection exported",
            f"Created:\n{folder}\n\n"
            f"Copied files: {manifest.get('copied_file_count', 0)}\n"
            f"Selected packages: {manifest.get('unique_selected_package_count', 0)}\n"
            f"Declared dependencies copied: {manifest.get('dependency_package_count', 0)}\n\n"
            "JSON and TXT manifests are inside the folder.",
        )

    def _random_sim_open_last_export(self):
        folder = self.random_sim_last_export
        if folder is None or not folder.is_dir():
            QMessageBox.information(self, "Collection Exporter", "There is no completed export folder to open yet.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))

    def _random_sim_current_history_sim_id(self):
        if not hasattr(self, "random_sim_history_table"):
            return ""
        rows = self.random_sim_history_table.selectionModel().selectedRows()
        if not rows:
            return ""
        item = self.random_sim_history_table.item(rows[0].row(), 0)
        return str(item.data(Qt.UserRole) or "") if item is not None else ""

    def _random_sim_clear_history_gallery(self):
        if not hasattr(self, "random_sim_history_gallery_grid"):
            return
        grid = self.random_sim_history_gallery_grid
        while grid.count():
            item = grid.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def _random_sim_refresh_history(self):
        if not IS_DEV_BUILD or not hasattr(self, "random_sim_history_table"):
            return
        previous_sim_id = self._random_sim_current_history_sim_id()
        self.random_sim_usage_registry = load_usage_registry(self.random_sim_usage_registry_path)
        rows = sim_history(self.random_sim_usage_registry)
        self.random_sim_history_rows = rows
        table = self.random_sim_history_table
        table.blockSignals(True)
        table.setRowCount(0)
        wanted_row = -1
        for row_index, sim in enumerate(rows):
            table.insertRow(row_index)
            name_item = QTableWidgetItem(str(sim.get("name", "") or "Random Sim"))
            sim_id = str(sim.get("sim_id", "") or "")
            name_item.setData(Qt.UserRole, sim_id)
            table.setItem(row_index, 0, name_item)
            created = str(sim.get("created_at", "") or "").replace("T", " ")
            table.setItem(row_index, 1, QTableWidgetItem(created or "—"))
            selected_item = QTableWidgetItem(str(int(sim.get("selected_count", 0) or 0)))
            selected_item.setTextAlignment(Qt.AlignCenter)
            table.setItem(row_index, 2, selected_item)
            dependency_item = QTableWidgetItem(str(int(sim.get("dependency_count", 0) or 0)))
            dependency_item.setTextAlignment(Qt.AlignCenter)
            table.setItem(row_index, 3, dependency_item)
            if previous_sim_id and sim_id == previous_sim_id:
                wanted_row = row_index
        table.blockSignals(False)

        if rows:
            table.selectRow(wanted_row if wanted_row >= 0 else 0)
            self._random_sim_history_selected()
        else:
            self.random_sim_history_current = None
            self.random_sim_history_profile.setHtml(
                "<i>No exported Random Sims are recorded yet. Export a generated set to create the first profile.</i>"
            )
            self.random_sim_history_gallery_status.setText("No saved Random Sim gallery yet.")
            self._random_sim_clear_history_gallery()
            self.random_sim_history_open_folder_btn.setEnabled(False)
            self.random_sim_history_open_manifest_btn.setEnabled(False)

    def _random_sim_history_selected(self):
        sim_id = self._random_sim_current_history_sim_id()
        sim = next(
            (row for row in getattr(self, "random_sim_history_rows", []) if str(row.get("sim_id", "") or "") == sim_id),
            None,
        )
        self.random_sim_history_current = sim
        if not sim:
            return

        folder_text = str(sim.get("folder", "") or "").strip()
        manifest_text = str(sim.get("manifest", "") or "").strip()
        folder = Path(folder_text) if folder_text else None
        manifest_path = Path(manifest_text) if manifest_text else None
        selected_count = int(sim.get("selected_count", 0) or 0)
        dependency_count = int(sim.get("dependency_count", 0) or 0)
        generation = sim.get("generation_settings", {})
        generation_note = ""
        if isinstance(generation, dict) and generation:
            category_counts = generation.get("category_counts", {})
            active_categories = 0
            if isinstance(category_counts, dict):
                for value in category_counts.values():
                    try:
                        active_categories += 1 if int(value or 0) > 0 else 0
                    except (TypeError, ValueError):
                        continue
            type_rules = generation.get("type_rules", {})
            item_rules = generation.get("item_rules", {})
            generation_note = (
                f"<br><b>Generation profile:</b> {active_categories} active categories; "
                f"{len(type_rules) if isinstance(type_rules, dict) else 0} type rule(s); "
                f"{len(item_rules) if isinstance(item_rules, dict) else 0} item rule(s)"
            )

        folder_available = bool(folder is not None and folder.is_dir())
        manifest_available = bool(manifest_path is not None and manifest_path.is_file())
        folder_status = "Available" if folder_available else "Missing/moved"
        folder_display = folder_text or "—"
        self.random_sim_history_profile.setHtml(
            f"<h3>{html.escape(str(sim.get('name', '') or 'Random Sim'))}</h3>"
            f"<b>Created:</b> {html.escape(str(sim.get('created_at', '') or '—').replace('T', ' '))}<br>"
            f"<b>Selected CC:</b> {selected_count} &nbsp; <b>Dependencies:</b> {dependency_count}<br>"
            f"<b>CAS Curator:</b> {html.escape(str(sim.get('app_version', '') or '—'))}<br>"
            f"<b>Original MDB:</b> {html.escape(str(sim.get('master_database', '') or '—'))}<br>"
            f"<b>Export folder:</b> {html.escape(folder_display)} ({folder_status})"
            f"{generation_note}"
        )
        self.random_sim_history_open_folder_btn.setEnabled(folder_available)
        self.random_sim_history_open_manifest_btn.setEnabled(manifest_available)
        self._random_sim_render_history_gallery(sim_id)

    def _random_sim_record_for_identity(self, identity):
        wanted = str(identity or "")
        if not wanted or not self.mdb:
            return None
        for row_index, record in enumerate(self.model.records):
            if package_identity(record) == wanted:
                return row_index, record
        return None

    def _random_sim_render_history_gallery(self, sim_id):
        if not hasattr(self, "random_sim_history_gallery_grid"):
            return
        self._random_sim_clear_history_gallery()
        links = usages_for_sim(sim_id, self.random_sim_usage_registry, role="selected")
        dependency_links = usages_for_sim(sim_id, self.random_sim_usage_registry, role="dependency")
        self.random_sim_history_gallery_status.setText(
            f"Selected CC gallery: {len(links)} item(s) | Explicit dependency-only packages: {len(dependency_links)}"
        )
        if not links:
            empty = QLabel("No directly selected CC links were recorded for this Sim.")
            empty.setObjectName("EditHelp")
            empty.setWordWrap(True)
            self.random_sim_history_gallery_grid.addWidget(empty, 0, 0)
            return

        columns = 3
        for index, link in enumerate(links):
            identity = str(link.get("identity", "") or "")
            matched = self._random_sim_record_for_identity(identity)
            record = matched[1] if matched else None
            package = str((record or {}).get("Package", "") or link.get("package", "") or "CC package")
            categories = list(link.get("categories", []) or [])
            category_text = ", ".join(str(value) for value in categories if str(value)) or "Selected CC"

            card = QFrame()
            card.setObjectName("ThumbnailCard")
            card.setMinimumWidth(210)
            card.setMaximumWidth(310)
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(7, 7, 7, 7)
            card_layout.setSpacing(5)

            category_label_widget = QLabel(category_text)
            category_label_widget.setObjectName("StatsStrong")
            category_label_widget.setWordWrap(True)
            card_layout.addWidget(category_label_widget)

            thumb = QLabel()
            thumb.setObjectName("ThumbnailImage")
            thumb.setAlignment(Qt.AlignCenter)
            thumb.setMinimumSize(190, 135)
            thumb.setMaximumHeight(180)
            thumb.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            thumbnail_loaded = False
            if record is not None:
                effective = dict(record)
                row_num = int(record.get("_ExcelRow", "0") or 0)
                if row_num and row_num in self.pending_changes and "Image" in self.pending_changes[row_num]:
                    effective["_ThumbnailTarget"] = self.pending_changes[row_num].get("Image", "")
                target = effective.get("_ThumbnailTarget", "") or effective.get("Image", "")
                path = target_to_local_path(target) if target else None
                if path is not None and path.exists():
                    pix = QPixmap(str(path))
                    if not pix.isNull():
                        thumb.setPixmap(
                            pix.scaled(250, 170, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                        )
                        thumbnail_loaded = True
            if not thumbnail_loaded:
                if record is not None:
                    cc_type = str(record.get("CC Type", "") or category_text)
                    subtype = str(record.get("Subtype", "") or "")
                    casp = str(record.get("CASP Instance IDs", "") or "")
                    casp_short = casp if len(casp) <= 42 else casp[:39] + "…"
                    fallback = "No thumbnail\n" + cc_type
                    if subtype:
                        fallback += "\n" + subtype
                    if casp_short:
                        fallback += "\nCASP: " + casp_short
                    thumb.setText(fallback)
                else:
                    thumb.setText("Original CC not found\nin the current MDB")
            card_layout.addWidget(thumb)

            short_package = package if len(package) <= 38 else package[:35] + "…"
            open_button = QPushButton(short_package)
            open_button.setToolTip(package + ("\nOpen original CC in Library" if matched else "\nNot present in the current MDB"))
            open_button.setEnabled(bool(matched))
            if matched:
                open_button.clicked.connect(
                    lambda checked=False, ident=identity: self._random_sim_open_original_cc(ident)
                )
            card_layout.addWidget(open_button)

            if record is not None:
                subtype = str(record.get("Subtype", "") or "—")
                health = str(record.get("Health Status", "") or "—")
                risk = str(record.get("Risk", "") or "—")
                meta = QLabel(f"{subtype}\nHealth: {health} | Risk: {risk}")
            else:
                meta = QLabel("Saved Sim link preserved; original MDB record is unavailable here.")
            meta.setObjectName("EditHelp")
            meta.setWordWrap(True)
            card_layout.addWidget(meta)

            row = index // columns
            column = index % columns
            self.random_sim_history_gallery_grid.addWidget(card, row, column)

    def _random_sim_open_original_cc(self, identity):
        matched = self._random_sim_record_for_identity(identity)
        if not matched:
            QMessageBox.information(
                self,
                "Random Sim history",
                "That original CC package is not present in the currently opened Master Database.",
            )
            return
        source_row, record = matched
        source_index = self.model.index(source_row, 0)
        proxy_index = self.proxy.mapFromSource(source_index)
        if not proxy_index.isValid():
            self.clear_filters()
            proxy_index = self.proxy.mapFromSource(source_index)
        if not proxy_index.isValid():
            return
        self._set_app_page("library")
        self.table.setCurrentIndex(proxy_index)
        self.table.selectionModel().select(
            proxy_index,
            QItemSelectionModel.ClearAndSelect | QItemSelectionModel.Rows,
        )
        self.table.scrollTo(proxy_index, QTableView.PositionAtCenter)
        if hasattr(self, "gallery"):
            self.gallery.scrollTo(proxy_index, QAbstractItemView.PositionAtCenter)
        self.current_source_row = source_row
        self.show_record(record)
        self.statusBar().showMessage(
            f"Opened original CC from Random Sim history: {record.get('Package', 'package')}", 6000
        )

    def _random_sim_history_open_folder(self):
        sim = getattr(self, "random_sim_history_current", None)
        if not isinstance(sim, dict):
            return
        folder_text = str(sim.get("folder", "") or "").strip()
        folder = Path(folder_text) if folder_text else None
        if folder is None or not folder.is_dir():
            QMessageBox.information(self, "Random Sim history", "That exported Sim folder is missing or has been moved.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))

    def _random_sim_history_open_manifest(self):
        sim = getattr(self, "random_sim_history_current", None)
        if not isinstance(sim, dict):
            return
        manifest_text = str(sim.get("manifest", "") or "").strip()
        manifest_path = Path(manifest_text) if manifest_text else None
        if manifest_path is None or not manifest_path.is_file():
            QMessageBox.information(self, "Random Sim history", "That Sim manifest is missing or has been moved.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(manifest_path)))


    def _set_app_page(self, page: str):
        allowed = {"library", "scan", "guide"}
        if IS_DEV_BUILD:
            allowed.add("random_sim")
        if page not in allowed:
            page = "library"
        self.app_page = page
        is_library = page == "library"
        is_scan = page == "scan"
        is_guide = page == "guide"
        is_random_sim = page == "random_sim"
        self.library_page_btn.setChecked(is_library)
        self.scan_page_btn.setChecked(is_scan)
        self.guide_page_btn.setChecked(is_guide)
        if self.random_sim_page_btn is not None:
            self.random_sim_page_btn.setChecked(is_random_sim)

        # Keep the compact themed header visible on every top-level page.
        # Library-specific rows are shown only in Library mode.
        self.filter_card.setVisible(True)
        for widget in self.library_header_widgets:
            widget.setVisible(is_library)
        self.scan_shell_badge.setVisible(is_scan)
        self.library_mdb_row.setVisible(is_library)
        self.library_filters_row.setVisible(is_library)

        self.splitter.setVisible(is_library)
        self.pinned_bar.setVisible(is_library and bool(self.pinned_packages))
        self.scan_page.setVisible(is_scan)
        self.guide_page.setVisible(is_guide)
        if self.random_sim_page is not None:
            self.random_sim_page.setVisible(is_random_sim)
        if is_scan:
            if self.mdb:
                global_mdb_path = str(self.mdb.path.resolve())
                if self.scan_mdb_path != global_mdb_path:
                    self.scan_mdb_path = global_mdb_path
                    self.scan_mdb_edit.setText(global_mdb_path)
            else:
                self.scan_mdb_path = ""
                self.scan_mdb_override = False
                self.scan_mdb_edit.clear()
            self._update_scan_start_state()
            self.statusBar().showMessage(("CAS Scanner — Run Scan keeps source files in place; Routing Preview is available" if IS_PUBLIC_BUILD else "Scanner — Start Scan keeps source files in place; Routing Preview is available"))
        elif is_random_sim:
            self._refresh_random_sim_data()
            self._random_sim_refresh_history()
            self.statusBar().showMessage("Random Sim Studio — DEV generator, Collection Exporter + Sim history gallery")
        elif is_guide:
            self.statusBar().showMessage("Guide — step-by-step instructions for Library, Profiles and CAS scanning")
        elif self.mdb:
            self.statusBar().showMessage(f"Library: {self.mdb.path.resolve()}")
        else:
            self.statusBar().showMessage("No Master Database loaded — choose File > Open Master Database…")

    def _open_scan_source_folder(self):
        raw = str(self.scan_source_path or "").strip()
        if not raw:
            QMessageBox.information(self, "CC Source Folder", "Choose a CC source folder first.")
            return
        folder = Path(raw)
        if not folder.is_dir():
            QMessageBox.warning(self, "CC Source Folder unavailable", f"The selected source folder is not available:\n\n{raw}")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder.resolve()))):
            QMessageBox.warning(self, "CC Source Folder", f"Windows could not open this folder:\n\n{raw}")

    def _choose_scan_source_folder(self):
        start = self.scan_source_path or str(self.settings.value("last_dir", str(Path.home())))
        folder = QFileDialog.getExistingDirectory(self, "Choose CC scan source folder", start)
        if not folder:
            return
        self.scan_source_path = folder
        self.settings.setValue("scan_source_folder", folder)
        self.scan_source_edit.setText(folder)
        self.scan_log.appendPlainText(f"\nSelected scan source: {folder}")
        self._update_scan_start_state()

    def _open_routing_preview_root(self):
        raw = str(self.routing_preview_root or "").strip()
        if not raw:
            QMessageBox.information(self, "Sort destination", "Choose a Sort destination first.")
            return
        folder = Path(raw)
        if not folder.is_dir():
            QMessageBox.warning(self, "Sort destination unavailable", f"The selected Sort destination is not available:\n\n{raw}")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder.resolve()))):
            QMessageBox.warning(self, "Sort destination", f"Windows could not open this folder:\n\n{raw}")

    def _choose_routing_preview_root(self):
        start = self.routing_preview_root or self.scan_source_path or str(self.settings.value("last_dir", str(Path.home())))
        folder = QFileDialog.getExistingDirectory(self, "Choose routing destination root", start)
        if not folder:
            return
        self.routing_preview_root = str(Path(folder).resolve())
        self.settings.setValue("routing_preview_root", self.routing_preview_root)
        self.routing_root_edit.setText(self.routing_preview_root)
        self.scan_log.appendPlainText(f"Selected routing preview destination: {self.routing_preview_root}")
        self._update_scan_start_state()

    def _scan_use_opened_mdb(self):
        if not self.mdb:
            QMessageBox.information(self, "Scan MDB", "Open a Master Database in the Library first, or use Browse… to choose another MDB.")
            return
        self.scan_mdb_path = str(self.mdb.path.resolve())
        self.scan_mdb_edit.setText(self.scan_mdb_path)
        self.scan_log.appendPlainText(f"Selected opened MDB: {self.scan_mdb_path}")
        self._update_scan_start_state()

    def _choose_scan_mdb(self):
        # Scan no longer owns MDB selection. Keep this legacy callback inert in
        # case an older signal/session tries to invoke it.
        self.statusBar().showMessage("Open, create, or change the Master Database from Library.", 5000)
        return

    def _on_routing_ack_toggled(self, checked: bool):
        if hasattr(self, "real_routing_check"):
            self.real_routing_check.setText(
                ("☑  I understand: Run Routing will MOVE source files" if IS_PUBLIC_BUILD else "☑  I understand: routing MOVES source files")
                if checked
                else ("☐  I understand: Run Routing will MOVE source files" if IS_PUBLIC_BUILD else "☐  I understand: routing MOVES source files")
            )
        self._update_scan_start_state()

    def _update_scan_start_state(self):
        if not hasattr(self, "scan_start_btn"):
            return
        running = self.scan_process is not None and self.scan_process.state() != QProcess.NotRunning
        source_ok = bool(self.scan_source_path and Path(self.scan_source_path).is_dir())
        mdb_ok = bool(self.scan_mdb_path and Path(self.scan_mdb_path).is_file() and self.scan_mdb_path.lower().endswith(".xlsx"))
        routing_ok = bool(self.routing_preview_root and Path(self.routing_preview_root).is_dir())
        preview_mode_ok = self._scan_mode_code() != "THUMBNAIL_BACKFILL" if hasattr(self, "scan_mode_combo") else False
        self.scan_start_btn.setEnabled(source_ok and mdb_ok and not running)
        if hasattr(self, "routing_preview_btn"):
            self.routing_preview_btn.setEnabled(source_ok and mdb_ok and routing_ok and preview_mode_ok and not running)
        if hasattr(self, "real_routing_btn"):
            acknowledged = bool(getattr(self, "real_routing_check", None) and self.real_routing_check.isChecked())
            real_mode_ok = self._scan_mode_code() in ("INDIVIDUAL", "BIG_REVIEW_BATCH") if hasattr(self, "scan_mode_combo") else False
            self.real_routing_btn.setEnabled(source_ok and mdb_ok and routing_ok and real_mode_ok and acknowledged and not running)

    def _scan_mode_code(self) -> str:
        label = self.scan_mode_combo.currentText()
        if label.startswith("Thumbnail"):
            return "THUMBNAIL_BACKFILL"
        if label.startswith("Big Review"):
            return "BIG_REVIEW_BATCH"
        return "INDIVIDUAL"

    def _start_scan_test(self):
        self._start_scan_operation(preview_routing=False)

    def _start_routing_preview(self):
        self._start_scan_operation(preview_routing=True)

    def _start_real_routing_test(self):
        self._start_scan_operation(real_routing=True)

    def _create_scan_mdb_backup(self, target: Path) -> Path:
        """Create a full pre-scan workbook backup beside the selected MDB."""
        backup_dir = target.parent / "MDB_Backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
        backup_path = backup_dir / f"{target.stem}_scan_backup_{stamp}{target.suffix}"
        shutil.copy2(target, backup_path)
        return backup_path

    def _prepare_scan_mdb_compatibility(self, target: Path) -> bool:
        """Stop Scan and offer a separate compatible MDB copy when needed.

        Some workbooks saved by other spreadsheet applications contain OOXML
        extensions that the scanner's openpyxl runtime cannot parse.  Never
        rewrite the user's original MDB here: create and verify a separate copy,
        switch CAS Curator to it, then make the user deliberately run Scan again.
        """
        try:
            report = inspect_mdb_compatibility(target)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "MDB compatibility check failed",
                "CAS Curator could not inspect this Master Database safely.\n\n"
                f"MDB:\n{target}\n\nError:\n{exc}",
            )
            return False

        if report.error:
            QMessageBox.critical(
                self,
                "MDB compatibility check failed",
                "CAS Curator could not verify this Master Database before Scan.\n\n"
                f"MDB:\n{target}\n\n{report.error}",
            )
            return False
        if not report.needs_compatible_copy:
            return True

        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("MDB compatibility issue")
        msg.setText("This Master Database contains spreadsheet-specific metadata that prevents the scanner from opening it.")
        details = [
            f"Scanner-blocking fill-style extensions: {report.patternfill_extlst_count:,}",
            f"Scanner-blocking data-validation attributes: {report.data_validation_view_count:,}",
        ]
        if report.vendor_extlst_count:
            details.append(f"Spreadsheet-app extension blocks in the workbook: {report.vendor_extlst_count:,}")
        msg.setInformativeText(
            "\n".join(details)
            + "\n\nCAS Curator can create a separate scanner-compatible copy beside the original MDB. "
              "The original workbook will NOT be changed. The copy removes spreadsheet-app-specific extension blocks that openpyxl cannot preserve, plus the known attributes that block the scanner from opening the workbook, then CAS Curator verifies the copy before using it.\n\n"
              "Keep the original workbook if you rely on application-specific spreadsheet metadata. Standard workbook data remains in the compatible copy.\n\n"
              "Create a compatible copy now?"
        )
        create_btn = msg.addButton("Create compatible copy…", QMessageBox.ButtonRole.AcceptRole)
        cancel_btn = msg.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
        msg.setDefaultButton(create_btn)
        msg.setEscapeButton(cancel_btn)
        msg.exec()
        if msg.clickedButton() != create_btn:
            return False

        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            result = create_compatible_copy(target)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Could not create compatible MDB copy",
                "The original MDB was left untouched.\n\n"
                f"MDB:\n{target}\n\nError:\n{exc}",
            )
            return False
        finally:
            QApplication.restoreOverrideCursor()

        compatible = result.compatible_path.resolve()

        # Verify with the exact Python 3.8/openpyxl runtime used by the frozen
        # scanner, not the GUI's private Python environment.
        python38, python38_error = self._resolve_scanner_python38()
        if not python38:
            try:
                compatible.unlink(missing_ok=True)
            except Exception:
                pass
            QMessageBox.critical(
                self,
                "Could not verify compatible MDB copy",
                "CAS Curator created the temporary compatible copy but could not verify it because the scanner's Python 3.8 runtime is unavailable. "
                "The compatible copy was removed and the original MDB was left untouched.\n\n"
                + python38_error,
            )
            return False
        verify_code = (
            "import sys,warnings; "
            "warnings.filterwarnings('ignore'); "
            "import openpyxl; "
            "wb=openpyxl.load_workbook(sys.argv[1], read_only=False, data_only=False); "
            "wb.close()"
        )
        try:
            verified = subprocess.run(
                [python38, "-c", verify_code, str(compatible)],
                capture_output=True,
                text=True,
                timeout=180,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0,
            )
        except Exception as exc:
            try:
                compatible.unlink(missing_ok=True)
            except Exception:
                pass
            QMessageBox.critical(
                self,
                "Could not verify compatible MDB copy",
                "Verification could not be completed. The compatible copy was removed and the original MDB was left untouched.\n\n"
                f"Error:\n{exc}",
            )
            return False
        if verified.returncode != 0:
            detail = (verified.stderr or verified.stdout or "Unknown openpyxl verification error").strip()
            try:
                compatible.unlink(missing_ok=True)
            except Exception:
                pass
            QMessageBox.critical(
                self,
                "Compatible MDB verification failed",
                "CAS Curator did not switch to the copy because the scanner's Excel library still could not open it. "
                "The failed copy was removed and the original MDB was left untouched.\n\n"
                f"Details:\n{detail[-4000:]}",
            )
            return False

        opened_target = bool(self.mdb and self.mdb.path.resolve() == target)
        self.scan_mdb_path = str(compatible)
        if hasattr(self, "scan_mdb_edit"):
            self.scan_mdb_edit.setText(self.scan_mdb_path)

        if opened_target:
            self.scan_mdb_override = False
            message_tail = "CAS Curator will now open the compatible copy in Library. Run Scan again after it finishes loading."
        else:
            self.scan_mdb_override = True
            message_tail = "The Scan target has been switched to the compatible copy. Run Scan again when ready."

        QMessageBox.information(
            self,
            "Compatible MDB copy created",
            "The original MDB was not changed.\n\n"
            f"Compatible copy:\n{compatible}\n\n"
            f"Removed spreadsheet-app extension blocks from the copy: {result.removed_vendor_extlst:,}\n"
            f"Removed blocking data-validation attributes: {result.removed_data_validation_view:,}\n\n"
            + message_tail,
        )
        self.scan_log.appendPlainText(
            f"GUI: created scanner-compatible MDB copy: {compatible} "
            f"(vendor extension blocks removed={result.removed_vendor_extlst}, "
            f"data-validation blockers removed={result.removed_data_validation_view})."
        )
        if opened_target:
            self.load_file(compatible)
        self._update_scan_start_state()
        return False

    def _reset_scan_stats(self):
        self.scan_stats = {
            "scanned": 0,
            "summary_records": 0,
            "risk_safe": 0,
            "risk_moderate": 0,
            "risk_high": 0,
            "risk_extreme": 0,
            "thumb_extracted": 0,
            "thumb_reused": 0,
            "thumb_no_preview": 0,
            "thumb_errors": 0,
            "moved": 0,
            "duplicates_routed": 0,
            "mdb_packages": 0,
            "mdb_casp": 0,
            "duration": "",
            "preview_routes": 0,
        }

    def _handle_scan_zip_archives(self, source: Path) -> Optional[dict]:
        """Offer explicit Scan-time handling when ZIP archives are present.

        Returns a summary dict when Scan may continue, or None when the operation
        should stop.  This never moves/deletes the ZIP archives.
        """
        zip_paths = sorted(
            (p.resolve() for p in source.rglob("*") if p.is_file() and p.suffix.lower() == ".zip"),
            key=lambda p: str(p).casefold(),
        )
        if not zip_paths:
            return {"action": "none", "zip_count": 0, "extracted_files": 0, "package_files": 0}

        current_packages = sum(1 for p in source.rglob("*") if p.is_file() and p.suffix.lower() == ".package")
        sample = "\n".join(f"• {p.relative_to(source)}" for p in zip_paths[:8])
        if len(zip_paths) > 8:
            sample += f"\n• …and {len(zip_paths) - 8:,} more"

        box = QMessageBox(self)
        box.setIcon(QMessageBox.Information)
        box.setWindowTitle("ZIP archives found")
        box.setTextFormat(Qt.RichText)
        box.setText(f"<b>{len(zip_paths):,} ZIP archive(s) found</b>")
        box.setInformativeText(
            "Extract beside each ZIP so its contents can be included in this Scan?\n"
            "Original ZIPs stay unchanged and existing files are never overwritten.\n\n"
            f"Packages already unpacked: {current_packages:,}\n\n"
            "Found:\n" + sample
        )
        extract_btn = box.addButton("Extract && Continue", QMessageBox.ButtonRole.AcceptRole)
        keep_btn = box.addButton("Keep Packed", QMessageBox.ButtonRole.ActionRole)
        cancel_btn = box.addButton(QMessageBox.Cancel)
        box.setDefaultButton(cancel_btn)
        box.setEscapeButton(cancel_btn)
        box.setWindowModality(Qt.ApplicationModal)
        box.raise_()
        box.activateWindow()
        QApplication.processEvents()
        box.exec()
        chosen = box.clickedButton()

        if chosen == cancel_btn or chosen is None:
            if hasattr(self, "scan_log"):
                self.scan_log.appendPlainText("GUI: Scan cancelled at ZIP archive prompt.")
            return None
        if chosen == keep_btn:
            return {
                "action": "kept",
                "zip_count": len(zip_paths),
                "extracted_files": 0,
                "package_files": 0,
            }
        if chosen != extract_btn:
            return None

        # Preflight every archive before writing anything.  One unsafe/corrupt ZIP
        # or one overwrite conflict blocks the whole batch.
        try:
            plans = _preflight_scan_zip_plans(zip_paths)
        except (OSError, zipfile.BadZipFile, ValueError, FileExistsError) as exc:
            QMessageBox.warning(
                self,
                "ZIP extraction blocked",
                "No ZIPs were extracted. CAS Curator found a safety/integrity problem during preflight:\n\n"
                f"{exc}\n\n"
                "Fix/remove the affected archive, or run Scan again and choose Keep ZIPs packed.",
            )
            return None

        try:
            extracted = _extract_scan_zip_plans(plans)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "ZIP extraction failed",
                "The ZIP batch could not be completed. CAS Curator rolled back files created by this batch where possible.\n\n"
                f"{exc}\n\nScan was not started.",
            )
            return None

        extracted_packages = sum(1 for p in extracted if p.suffix.lower() == ".package")
        if hasattr(self, "scan_log"):
            self.scan_log.appendPlainText(
                f"GUI: ZIP intake extracted {len(extracted):,} file(s) from {len(plans):,} archive(s); "
                f"{extracted_packages:,} package file(s) are now available to this Scan. Originals were kept."
            )
        self.statusBar().showMessage(
            f"ZIP intake: {len(plans):,} extracted; continuing Scan", 5000
        )
        return {
            "action": "extracted",
            "zip_count": len(plans),
            "extracted_files": len(extracted),
            "package_files": extracted_packages,
        }

    def _handle_scan_sims3packs(self, source: Path) -> Optional[dict]:
        """Analyse Sims3Packs found in the Scan tree and optionally extract safe CAS copies.

        This is intentionally conservative. Only containers whose manifest Type is
        CASpart are eligible for extraction. Lots, worlds, Sims/households,
        patterns, objects and unknown types remain packed for the existing
        SIMS3PACK routing path. Originals are never modified or deleted.
        """
        sims3packs = sorted(
            (p.resolve() for p in source.rglob("*") if p.is_file() and p.suffix.lower() == ".sims3pack"),
            key=lambda p: str(p).casefold(),
        )
        if not sims3packs:
            return {
                "action": "none",
                "count": 0,
                "eligible": 0,
                "protected": 0,
                "matched": 0,
                "extracted_packages": 0,
                "errors": 0,
            }

        standalone_packages = sorted(
            (p.resolve() for p in source.rglob("*") if p.is_file() and p.suffix.lower() == ".package"),
            key=lambda p: str(p).casefold(),
        )
        try:
            existing_hashes = _sims3pack_package_hashes(standalone_packages)
        except OSError as exc:
            QMessageBox.warning(
                self,
                "Sims3Pack analysis blocked",
                "CAS Curator could not hash the existing standalone packages safely. No Sims3Packs were changed.\n\n"
                f"Error:\n{exc}",
            )
            return None

        plans = []
        errors = []
        for path in sims3packs:
            try:
                plan = inspect_sims3pack(path)
                annotate_package_alternatives(plan, existing_hashes)
                plans.append(plan)
            except Exception as exc:
                errors.append((path, str(exc)))

        eligible = [p for p in plans if p.get("status") == "eligible"]
        protected = [p for p in plans if p.get("status") != "eligible"]
        matched_entries = sum(int(p.get("matched_package_count", 0) or 0) for p in eligible)
        missing_entries = sum(int(p.get("missing_package_count", 0) or 0) for p in eligible)
        total_cas_parts = sum(int(p.get("package_count", 0) or 0) for p in eligible)
        multipart_eligible = [p for p in eligible if int(p.get("package_count", 0) or 0) > 1]
        eligible_with_missing = [p for p in eligible if int(p.get("missing_package_count", 0) or 0) > 0]

        lines = []
        for plan in plans[:10]:
            try:
                rel = Path(plan["path"]).relative_to(source)
            except Exception:
                rel = Path(plan["path"]).name
            ptype = str(plan.get("type") or "Unknown")
            if plan.get("status") == "eligible":
                part_count = int(plan.get("package_count", 0) or 0)
                missing = int(plan.get("missing_package_count", 0) or 0)
                matched = int(plan.get("matched_package_count", 0) or 0)
                if part_count > 1:
                    detail = f"CASpart set — {part_count} package parts"
                    if missing:
                        detail += f"; {missing} missing"
                    if matched:
                        detail += f"; {matched} already present"
                    if not missing:
                        detail += "; complete standalone set already present"
                elif missing:
                    detail = f"CASpart — {missing} package copy available"
                    if matched:
                        detail += f", {matched} already present"
                else:
                    detail = f"CASpart — equivalent standalone package already present ({matched})"
            else:
                detail = f"kept packed — {ptype}"
            lines.append(f"• {rel} — {detail}")
        if len(plans) > 10:
            lines.append(f"• …and {len(plans) - 10:,} more analysed Sims3Pack(s)")
        for path, err in errors[:3]:
            try:
                rel = path.relative_to(source)
            except Exception:
                rel = path.name
            lines.append(f"• {rel} — unreadable; kept packed ({err})")
        if len(errors) > 3:
            lines.append(f"• …and {len(errors) - 3:,} more unreadable Sims3Pack(s)")

        box = QMessageBox(self)
        box.setIcon(QMessageBox.Information if not errors else QMessageBox.Warning)
        box.setWindowTitle("Sims3Packs found")
        box.setTextFormat(Qt.RichText)
        box.setText(f"<b>{len(sims3packs):,} Sims3Pack file(s) found</b>")
        summary_lines = [
            f"CAS packages available to extract: {missing_entries:,}",
            f"Equivalent standalone packages already present: {matched_entries:,}",
            f"Protected / unknown kept packed: {len(protected) + len(errors):,}",
        ]
        if multipart_eligible:
            summary_lines.append(f"Multi-part CAS sets: {len(multipart_eligible):,}")
        box.setInformativeText(
            "\n".join(summary_lines)
            + "\n\nOnly CASpart content is extractable. Protected content stays packed. "
              "Original Sims3Pack files stay unchanged.\n\n"
            + "\n".join(lines)
        )

        extract_btn = None
        if eligible_with_missing:
            extract_btn = box.addButton("Extract CAS && Continue", QMessageBox.ButtonRole.AcceptRole)
        keep_btn = box.addButton("Keep Packed && Continue", QMessageBox.ButtonRole.ActionRole)
        cancel_btn = box.addButton(QMessageBox.Cancel)
        box.setDefaultButton(keep_btn)
        box.setEscapeButton(cancel_btn)
        box.setWindowModality(Qt.ApplicationModal)
        box.raise_()
        box.activateWindow()
        QApplication.processEvents()
        box.exec()
        chosen = box.clickedButton()

        if chosen == cancel_btn or chosen is None:
            if hasattr(self, "scan_log"):
                self.scan_log.appendPlainText("GUI: Scan cancelled at Sims3Pack analysis prompt.")
            return None
        if chosen == keep_btn or extract_btn is None:
            return {
                "action": "analysed",
                "count": len(sims3packs),
                "eligible": len(eligible),
                "protected": len(protected) + len(errors),
                "matched": matched_entries,
                "cas_parts": total_cas_parts,
                "multipart_sets": len(multipart_eligible),
                "extracted_packages": 0,
                "errors": len(errors),
            }
        if chosen != extract_btn:
            return None

        try:
            preflight_missing_cas_batch(eligible_with_missing)
            extracted = extract_missing_cas_batch(eligible_with_missing)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Sims3Pack extraction blocked",
                "No protected Sims3Pack was changed. CAS Curator could not safely complete the CAS extraction batch.\n\n"
                f"{exc}\n\n"
                "Fix the conflict/problem, or run Scan again and choose Analyse only / Keep packed.",
            )
            return None

        if hasattr(self, "scan_log"):
            self.scan_log.appendPlainText(
                f"GUI: Sims3Pack intake extracted {len(extracted):,} missing CAS package copy/copies "
                f"from {len(eligible_with_missing):,} Sims3Pack(s); originals were kept."
            )
        self.statusBar().showMessage(
            f"Sims3Pack intake: {len(extracted):,} CAS package copy/copies extracted; continuing Scan", 6000
        )
        return {
            "action": "extracted",
            "count": len(sims3packs),
            "eligible": len(eligible),
            "protected": len(protected) + len(errors),
            "matched": matched_entries,
            "cas_parts": total_cas_parts,
            "multipart_sets": len(multipart_eligible),
            "extracted_packages": len(extracted),
            "errors": len(errors),
        }

    def _routing_scan_state_summary(self, source: Path, target_mdb: Path) -> Dict[str, object]:
        """Compare current source packages with MDB filename/hash identity.

        This is a read-only routing preflight. It does not classify packages or
        change scanner/Routing decisions. A matching filename with no usable
        stored hash is treated as needing a fresh scan because current identity
        cannot be verified safely.
        """
        fresh = MDBWorkbook().load(target_mdb)
        records_by_name: Dict[str, List[Dict[str, str]]] = {}
        for rec in fresh.records:
            name = Path(str(rec.get("Package", "") or "").strip()).name
            if name:
                records_by_name.setdefault(name.casefold(), []).append(rec)

        ready: List[str] = []
        changed: List[str] = []
        missing: List[str] = []
        errors: List[str] = []
        package_paths = sorted(
            (p for p in source.rglob("*") if p.is_file() and p.suffix.lower() == ".package"),
            key=lambda p: str(p).casefold(),
        )
        for path in package_paths:
            candidates = records_by_name.get(path.name.casefold(), [])
            if not candidates:
                missing.append(path.name)
                continue
            stored_hashes = {
                str(rec.get("Package Hash", "") or "").strip().lower()
                for rec in candidates
                if str(rec.get("Package Hash", "") or "").strip()
            }
            if not stored_hashes:
                changed.append(path.name)
                continue
            try:
                digest = self._sha256_file(path).lower()
            except OSError as exc:
                errors.append(f"{path.name}: {exc}")
                changed.append(path.name)
                continue
            if digest in stored_hashes:
                ready.append(path.name)
            else:
                changed.append(path.name)
        return {
            "total": len(package_paths),
            "ready": ready,
            "changed": changed,
            "missing": missing,
            "errors": errors,
        }

    def _routing_scan_state_decision(self, source: Path, target_mdb: Path, real_routing: bool) -> str:
        """Return continue / scan / cancel after the routing identity preflight."""
        try:
            summary = self._routing_scan_state_summary(source, target_mdb)
        except Exception as exc:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Could not verify scan state")
            msg.setText("CAS Curator could not compare the source packages with the selected Master Database.")
            msg.setInformativeText(
                f"{exc}\n\nRun Scan first is recommended before routing. You can still continue deliberately, or cancel."
            )
            scan_btn = msg.addButton("Run Scan First", QMessageBox.ButtonRole.AcceptRole)
            continue_btn = msg.addButton("Continue Routing Anyway", QMessageBox.ButtonRole.DestructiveRole)
            cancel_btn = msg.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
            msg.setDefaultButton(scan_btn)
            msg.exec()
            if msg.clickedButton() == scan_btn:
                return "scan"
            if msg.clickedButton() == continue_btn:
                return "continue_confirmed"
            return "cancel"

        ready = list(summary.get("ready", []))
        changed = list(summary.get("changed", []))
        missing = list(summary.get("missing", []))
        errors = list(summary.get("errors", []))
        if hasattr(self, "scan_log"):
            self.scan_log.appendPlainText(
                f"Routing scan-state check: Ready {len(ready):,}; Changed since scan {len(changed):,}; Not in MDB {len(missing):,}."
            )
        if not changed and not missing:
            return "continue"

        def sample_lines(label: str, names: List[str]) -> str:
            if not names:
                return ""
            shown = names[:6]
            text = f"\n{label}:\n" + "\n".join(f"• {name}" for name in shown)
            if len(names) > len(shown):
                text += f"\n• …and {len(names) - len(shown):,} more"
            return text

        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Scan recommended before routing")
        msg.setText("Some source packages are not current in this Master Database.")
        operation_note = (
            "Continue Routing Anyway will physically move files using the available routing evidence. "
            "Changed/unknown rows may keep stale metadata or may not receive complete File Location tracking."
            if real_routing else
            "Continue Routing Anyway will only create a dry-run preview; no source files will move and the MDB will not be written."
        )
        detail = (
            f"Packages checked: {int(summary.get('total', 0)):,}\n"
            f"Ready: {len(ready):,}\n"
            f"Changed since scan / not verifiable: {len(changed):,}\n"
            f"Not in MDB: {len(missing):,}\n\n"
            "Run Scan First is recommended so scanner-derived metadata and package identity are current.\n"
            + operation_note
            + sample_lines("Changed since scan", changed)
            + sample_lines("Not in MDB", missing)
        )
        if errors:
            detail += sample_lines("Could not hash", errors)
        msg.setInformativeText(detail)
        scan_btn = msg.addButton("Run Scan First", QMessageBox.ButtonRole.AcceptRole)
        continue_btn = msg.addButton("Continue Routing Anyway", QMessageBox.ButtonRole.DestructiveRole)
        cancel_btn = msg.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
        msg.setDefaultButton(scan_btn)
        msg.exec()
        if msg.clickedButton() == scan_btn:
            return "scan"
        if msg.clickedButton() == continue_btn:
            return "continue_confirmed"
        return "cancel"

    def _start_scan_operation(self, preview_routing: bool = False, real_routing: bool = False):
        action_name = "Run Routing" if real_routing else ("Preview Routing" if preview_routing else ("Run Scan" if IS_PUBLIC_BUILD else "Start Scan"))
        if hasattr(self, "scan_log"):
            self.scan_log.appendPlainText(f"\nGUI: {action_name} clicked; validating source and MDB…")
        self._update_scan_start_state()
        relevant_button = self.real_routing_btn if real_routing else (self.routing_preview_btn if preview_routing else self.scan_start_btn)
        if not relevant_button.isEnabled():
            if hasattr(self, "scan_log"):
                if real_routing:
                    self.scan_log.appendPlainText(
                        "GUI: Run Routing cannot start. Check source, MDB, destination and the acknowledgement checkbox."
                        if IS_PUBLIC_BUILD else
                        "GUI: Run Routing cannot start. Check source, MDB, destination, Individual/Big Review mode, and the acknowledgement checkbox."
                    )
                elif preview_routing:
                    self.scan_log.appendPlainText(
                        "GUI: routing preview cannot start because source, MDB, or preview destination is not valid."
                    )
                else:
                    self.scan_log.appendPlainText(
                        "GUI: scan cannot start because the source folder or MDB is not currently valid."
                    )
            self.raise_()
            self.activateWindow()
            QMessageBox.information(
                self,
                "Run Routing" if real_routing else ("Routing preview" if preview_routing else "Scanner"),
                ("Choose a valid source folder, MDB and destination, then tick the routing acknowledgement first."
                 if IS_PUBLIC_BUILD else
                 "Choose a valid source folder, MDB, destination, Individual or Big Review mode, and tick the routing acknowledgement first.")
                if real_routing
                else ("Choose a valid source folder, MDB, and routing preview destination first."
                      if preview_routing
                      else "Choose a valid source folder and MDB first."),
            )
            return

        if preview_routing and self._scan_mode_code() == "THUMBNAIL_BACKFILL":
            QMessageBox.information(self, "Routing preview", "Routing preview is available for Individual and Big Review scans, not Thumbnail Backfill.")
            return
        if real_routing and self._scan_mode_code() not in ("INDIVIDUAL", "BIG_REVIEW_BATCH"):
            QMessageBox.information(self, "Run Routing", "Routing is available for Individual and Big Review modes, not Thumbnail Backfill.")
            return

        source = Path(self.scan_source_path).resolve()
        target = Path(self.scan_mdb_path).resolve()
        routing_root = Path(self.routing_preview_root).resolve() if (preview_routing or real_routing) else None
        if not source.is_dir() or not target.is_file() or ((preview_routing or real_routing) and not routing_root.is_dir()):
            QMessageBox.warning(
                self,
                "Run Routing" if real_routing else ("Routing preview" if preview_routing else "Scanner"),
                "The selected source folder, MDB, or destination no longer exists."
                if (preview_routing or real_routing)
                else "The selected source folder or MDB no longer exists.",
            )
            self._update_scan_start_state()
            return

        if (not preview_routing) and self.mdb and self.mdb.path.resolve() == target and self.pending_changes:
            QMessageBox.warning(
                self,
                "Unsaved library edits",
                "The scanner target is the MDB currently open in Library and it has unsaved manual edits.\n\n"
                "Save or revert those edits before starting the scanner."
            )
            return

        zip_intake = {"action": "none", "zip_count": 0, "extracted_files": 0, "package_files": 0}
        sims3pack_intake = {
            "action": "none", "count": 0, "eligible": 0, "protected": 0,
            "matched": 0, "extracted_packages": 0, "errors": 0,
        }
        if not preview_routing and not real_routing:
            if not self._prepare_scan_mdb_compatibility(target):
                return
            zip_intake = self._handle_scan_zip_archives(source)
            if zip_intake is None:
                return
            # ZIP extraction may reveal Sims3Packs and/or matching standalone
            # packages, so analyse Sims3Packs only after the ZIP intake step.
            sims3pack_intake = self._handle_scan_sims3packs(source)
            if sims3pack_intake is None:
                return

        # Recount after optional ZIP/Sims3Pack extraction so newly unpacked
        # packages are included immediately in this same Scan.
        package_count = sum(1 for p in source.rglob("*") if p.is_file() and p.suffix.lower() == ".package")
        if self._scan_mode_code() != "THUMBNAIL_BACKFILL" and package_count == 0:
            QMessageBox.information(
                self,
                "Run Routing" if real_routing else ("Routing preview" if preview_routing else "Scanner"),
                "No .package files were found in the selected source folder.",
            )
            return

        routing_preflight_confirmed = False
        if preview_routing or real_routing:
            scan_state_decision = self._routing_scan_state_decision(source, target, real_routing)
            if scan_state_decision == "scan":
                QTimer.singleShot(0, self._start_scan_test)
                return
            if scan_state_decision == "continue_confirmed":
                # The attention dialog already stated exactly what this operation will do.
                # Treat Continue Routing Anyway as the final confirmation so users never
                # have to approve two consecutive proceed dialogs.
                routing_preflight_confirmed = True
            elif scan_state_decision != "continue":
                return

        if real_routing:
            try:
                if source == routing_root or source in routing_root.parents or routing_root in source.parents:
                    QMessageBox.warning(
                        self,
                        "Unsafe routing destination",
                        "For safety, the source and destination must be completely separate folders.\n\n"
                        f"Source:\n{source}\n\nDestination:\n{routing_root}"
                    )
                    return
            except Exception:
                pass
            try:
                destination_items = [p for p in routing_root.iterdir()]
            except Exception as exc:
                QMessageBox.warning(self, "Run Routing", f"Could not inspect the destination folder:\n{exc}")
                return
            destination_item_count = len(destination_items)

        if hasattr(self, "scan_log"):
            self.scan_log.appendPlainText("GUI: validation passed; preparing operation confirmation…")

        self.show()
        self.raise_()
        self.activateWindow()
        QApplication.processEvents()

        if routing_preflight_confirmed:
            if hasattr(self, "scan_log"):
                self.scan_log.appendPlainText(
                    "GUI: scan-state warning accepted; Continue Routing Anyway is serving as the final confirmation."
                )
        elif real_routing or preview_routing:
            # Routing changes files (or previews a destination plan), so keep one
            # deliberate confirmation. A normal Run Scan no longer gets an extra
            # Yes/No box after the user already clicked Run Scan.
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning if real_routing else QMessageBox.Information)
            msg.setTextFormat(Qt.RichText)
            if real_routing:
                msg.setWindowTitle("Run Routing?")
                msg.setText(f"<b>Move {package_count:,} package(s) now?</b>")
                msg.setInformativeText(
                    f"From: {source}\nTo: {routing_root}\n\n"
                    "Source files will be moved, not copied. Existing destination files are never overwritten. "
                    "The MDB will be updated after successful routing."
                )
            else:
                msg.setWindowTitle("Preview Routing?")
                msg.setText(f"<b>Preview routes for {package_count:,} package(s)?</b>")
                msg.setInformativeText(
                    f"Source: {source}\nDestination preview: {routing_root}\n\n"
                    "Dry run only — no source files move, no destination files are created, and the MDB is not written."
                )
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg.setDefaultButton(QMessageBox.No)
            msg.setWindowModality(Qt.ApplicationModal)
            msg.raise_()
            msg.activateWindow()
            QApplication.processEvents()
            result = msg.exec()
            if result != QMessageBox.Yes:
                if hasattr(self, "scan_log"):
                    self.scan_log.appendPlainText("GUI: operation cancelled at confirmation.")
                return
        elif hasattr(self, "scan_log"):
            self.scan_log.appendPlainText("GUI: validation passed; Run Scan starts directly (no redundant confirmation).")

        if getattr(sys, "frozen", False):
            runner_py = Path(sys.executable).resolve().parent / "scanner" / "CASCurator_ScannerRuntime.exe"
        else:
            runner_py = Path(__file__).resolve().parent / "scanner_integration_runner.py"
        if not runner_py.exists():
            QMessageBox.critical(self, "Scanner", f"Missing integration runner:\n{runner_py}")
            return

        python38, python38_error = self._resolve_scanner_python38()
        if not python38:
            QMessageBox.critical(
                self,
                "Python 3.8 scanner runtime not found",
                "The integrated scanner needs the same Python 3.8 runtime used by the validated scanner.\n\n"
                + python38_error,
            )
            return

        self.last_scan_backup_path = None
        if not preview_routing:
            try:
                self.last_scan_backup_path = self._create_scan_mdb_backup(target)
            except Exception as exc:
                QMessageBox.critical(
                    self,
                    "MDB backup failed",
                    "The scan was NOT started because the safety backup could not be created.\n\n"
                    f"MDB:\n{target}\n\nError:\n{exc}",
                )
                return

        self.scan_run_kind = "routing" if real_routing else ("preview" if preview_routing else "scan")
        self.last_scan_output_dir = None
        self.last_scan_log_path = None
        self.last_routing_preview_path = None
        self.last_sorted_root = routing_root if real_routing else None
        self.last_move_manifest_path = None
        self.last_recovery_manifest_path = None
        self.last_routing_location_tracking = {}
        if hasattr(self, "scan_recover_routing_btn"):
            self.scan_recover_routing_btn.setEnabled(False)
        self.last_scan_summary_text = ""
        self._post_scan_refresh_target = None
        self._reset_scan_stats()
        if hasattr(self, "scan_open_preview_btn"):
            self.scan_open_preview_btn.setEnabled(False)
        if hasattr(self, "scan_open_log_btn"):
            self.scan_open_log_btn.setEnabled(False)
        if hasattr(self, "scan_open_output_btn"):
            self.scan_open_output_btn.setEnabled(False)
        if hasattr(self, "routing_preview_table"):
            self.routing_preview_table.setRowCount(0)

        self.scan_log.clear()
        if IS_PUBLIC_BUILD:
            self.scan_log.appendPlainText(
                "CAS CURATOR - CONSERVATIVE CAS SCANNER\n\n"
                "This release reads direct CASP metadata and GEOM face counts. Filenames are never used to classify supported content types.\n"
                "Run Scan generates CASP IDs, supported CC Type and Highest Polys, extracts dedicated THUM previews when available, and records package hash/scan time.\n"
                "Whether a package contains or overlaps with EA resources is not assessed. Skin type, general IMG texture purpose, script/gameplay and Health/Risk are not supported scan results.\n"
                "Run Scan creates a mandatory MDB backup before scanner writes.\n"
                "Preview Routing before any move. Run Routing requires explicit acknowledgement and creates a recovery manifest.\n\n"
                "Experimental features can display provisional information for review, but they never control Routing automatically.\n"
            )
        if real_routing:
            self.scan_log.appendPlainText("STARTING ROUTING")
        elif preview_routing:
            self.scan_log.appendPlainText("STARTING ROUTING PREVIEW DRY RUN")
        else:
            self.scan_log.appendPlainText("STARTING SCAN")
        self.scan_log.appendPlainText(f"Source: {source}")
        self.scan_log.appendPlainText(f"MDB: {target}")
        if not IS_PUBLIC_BUILD:
            self.scan_log.appendPlainText(f"Mode: {self._scan_mode_code()}")
        self.scan_log.appendPlainText(f"Python 3.8: {python38}")
        self.scan_log.appendPlainText("Launch method: direct process (no cmd.exe / batch wrapper)")
        if self.last_scan_backup_path:
            self.scan_log.appendPlainText(f"MDB safety backup: {self.last_scan_backup_path}")
        if real_routing:
            self.scan_log.appendPlainText(f"REAL routing destination: {routing_root}")
            self.scan_log.appendPlainText("Safety: ROUTING ENABLED — source files WILL MOVE; MDB write ENABLED")
        elif preview_routing:
            self.scan_log.appendPlainText(f"Preview destination: {routing_root}")
            self.scan_log.appendPlainText("Safety: DRY RUN — source moves DISABLED; MDB write DISABLED")
        else:
            self.scan_log.appendPlainText("Safety: source-file moving/routing DISABLED")
            if zip_intake.get("action") == "extracted":
                self.scan_log.appendPlainText(
                    f"ZIP intake: extracted {zip_intake['zip_count']:,} archive(s) beside their ZIPs "
                    f"({zip_intake['extracted_files']:,} file(s), {zip_intake['package_files']:,} package file(s)); original ZIPs kept"
                )
            elif zip_intake.get("action") == "kept":
                self.scan_log.appendPlainText(
                    f"ZIP intake: kept {zip_intake['zip_count']:,} archive(s) packed; Routing may send them to ARCHIVES later"
                )
            if sims3pack_intake.get("action") == "extracted":
                self.scan_log.appendPlainText(
                    f"Sims3Pack intake: analysed {sims3pack_intake['count']:,}; extracted "
                    f"{sims3pack_intake['extracted_packages']:,} missing CAS package copy/copies; "
                    f"equivalent standalone packages already present={sims3pack_intake['matched']:,}; "
                    f"CAS package parts={sims3pack_intake.get('cas_parts', 0):,}; multi-part sets={sims3pack_intake.get('multipart_sets', 0):,}; "
                    f"protected/unreadable kept packed={sims3pack_intake['protected']:,}"
                )
            elif sims3pack_intake.get("action") == "analysed":
                self.scan_log.appendPlainText(
                    f"Sims3Pack intake: analysed {sims3pack_intake['count']:,}; no package copies extracted; "
                    f"equivalent standalone packages already present={sims3pack_intake['matched']:,}; "
                    f"CAS package parts={sims3pack_intake.get('cas_parts', 0):,}; multi-part sets={sims3pack_intake.get('multipart_sets', 0):,}; "
                    f"protected/unreadable kept packed={sims3pack_intake['protected']:,}"
                )
        self.scan_log.appendPlainText("")
        self.scan_total_packages = package_count
        self.scan_current_package = 0
        self.scan_progress.setRange(0, max(package_count + 1, 1))
        self.scan_progress.setValue(0)
        self.scan_progress_label.setText(f"0 / {package_count}")
        self.scan_current_file_label.setText("Starting scanner…")
        self.scan_status_label.setText(
            "Starting routing…" if real_routing else ("Starting routing preview dry run…" if preview_routing else "Starting Python 3.8 scanner process…")
        )
        if hasattr(self, "scan_output_tabs"):
            self.scan_output_tabs.setCurrentIndex(0)

        proc = QProcess(self)
        self.scan_process = proc
        proc.setProcessChannelMode(QProcess.MergedChannels)
        proc.readyReadStandardOutput.connect(self._scan_read_output)
        proc.finished.connect(self._scan_process_finished)
        proc.errorOccurred.connect(self._scan_process_error)
        proc.setWorkingDirectory(str(Path(__file__).resolve().parent))

        args = [
            "-u", str(runner_py),
            "--source", str(source),
            "--mdb", str(target),
            "--mode", self._scan_mode_code(),
            "--edition", APP_EDITION,
        ]
        if preview_routing:
            preview_workdir = Path(__file__).resolve().parent / "routing_preview_runs"
            args.extend([
                "--preview-routing",
                "--routing-root", str(routing_root),
                "--preview-workdir", str(preview_workdir),
            ])
        elif real_routing:
            args.extend([
                "--real-routing",
                "--routing-root", str(routing_root),
            ])

        self.scan_start_btn.setEnabled(False)
        if hasattr(self, "routing_preview_btn"):
            self.routing_preview_btn.setEnabled(False)
        if hasattr(self, "real_routing_btn"):
            self.real_routing_btn.setEnabled(False)
        self.scan_cancel_btn.setEnabled(False)
        proc.start(str(python38), args)

    def _resolve_scanner_python38(self):
        """Resolve the real Python 3.8 executable without shell quoting."""
        if getattr(sys, "frozen", False):
            scanner_exe = Path(sys.executable).resolve().parent / "scanner" / "CASCurator_ScannerRuntime.exe"
            if scanner_exe.is_file():
                return scanner_exe, ""
            return None, "The bundled CAS Curator scanner runtime is missing:\n" + str(scanner_exe)
        probe = (
            "import sys; print(sys.executable); "
            "raise SystemExit(0 if sys.version_info[:2] == (3, 8) else 8)"
        )
        attempts = []
        candidates = []

        py_launcher = shutil.which("py")
        if py_launcher:
            candidates.append(("py -3.8", [py_launcher, "-3.8", "-c", probe]))

        pymanager = shutil.which("pymanager")
        if pymanager:
            candidates.append(("pymanager -V:3.8", [pymanager, "exec", "-V:3.8", "-c", probe]))

        python38_cmd = shutil.which("python3.8")
        if python38_cmd:
            candidates.append(("python3.8", [python38_cmd, "-c", probe]))

        for label, command in candidates:
            try:
                completed = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=12,
                    shell=False,
                )
                output = (completed.stdout or "").strip().splitlines()
                if completed.returncode == 0 and output:
                    exe = Path(output[-1].strip().strip('"'))
                    if exe.is_file():
                        return exe, ""
                detail = (completed.stderr or completed.stdout or "no output").strip()
                attempts.append(f"{label}: exit {completed.returncode} — {detail}")
            except Exception as exc:
                attempts.append(f"{label}: {exc}")

        if not candidates:
            attempts.append("No py, pymanager, or python3.8 command was found on PATH.")

        return None, (
            "Could not resolve a usable Python 3.8 executable.\n\n"
            + "\n".join(attempts)
        )

    def _scan_read_output(self):
        if not self.scan_process:
            return
        raw = bytes(self.scan_process.readAllStandardOutput()).decode("utf-8", errors="replace")
        if not raw:
            return
        self._scan_output_buffer += raw.replace("\r\n", "\n").replace("\r", "\n")
        while "\n" in self._scan_output_buffer:
            line, self._scan_output_buffer = self._scan_output_buffer.split("\n", 1)
            self._handle_scan_output_line(line.rstrip())

    def _handle_scan_output_line(self, line: str):
        display_line = line
        if IS_PUBLIC_BUILD:
            # Keep internal edition/build identifiers out of the user-facing log.
            # The scanner core itself is unchanged; this is presentation only.
            if line.startswith("SCANNER EDITION:") or line.startswith("BUILD: DEV/PUBLIC"):
                display_line = ""
            elif "PUBLIC" in line:
                display_line = line.replace("PUBLIC", "release")

        # The frozen scanner core reports passthrough destinations using legacy
        # movement wording even during a normal Run Scan, where source moving is
        # disabled.  Correct only the GUI presentation; Routing itself and the
        # scanner core remain unchanged.
        if self.scan_run_kind == "scan":
            stripped = line.strip()
            if stripped.startswith("Sims3Pack files found:") and "moved intact to SIMS3PACK" in stripped:
                prefix = line[:len(line) - len(line.lstrip())]
                count_text = stripped.split("Sims3Pack files found:", 1)[1].split("->", 1)[0].strip()
                display_line = f"{prefix}Sims3Pack files found: {count_text} — preserved for later SIMS3PACK routing"
            elif stripped == "Non-package files were moved intact; no archives were extracted.":
                prefix = line[:len(line) - len(line.lstrip())]
                display_line = f"{prefix}Non-package files were preserved during Run Scan; passthrough movement occurs only during Run Routing."
        if display_line:
            self.scan_log.appendPlainText(display_line)
            bar = self.scan_log.verticalScrollBar()
            bar.setValue(bar.maximum())

        self._update_scan_stats_from_line(line)

        if line.startswith("ROUTING PREVIEW ITEM\t"):
            parts = line.split("\t", 5)
            if len(parts) == 6 and hasattr(self, "routing_preview_table"):
                _prefix, category, duplicate_flag, source_path, destination, note = parts
                row = self.routing_preview_table.rowCount()
                self.routing_preview_table.insertRow(row)
                source_item = QTableWidgetItem(Path(source_path).name)
                source_item.setToolTip(source_path)
                category_item = QTableWidgetItem(category)
                duplicate_item = QTableWidgetItem(duplicate_flag)
                destination_item = QTableWidgetItem(self._routing_preview_display_path(destination))
                destination_item.setData(Qt.UserRole, destination)
                destination_item.setToolTip(destination)
                note_item = QTableWidgetItem(note)
                self.routing_preview_table.setItem(row, 0, source_item)
                self.routing_preview_table.setItem(row, 1, category_item)
                self.routing_preview_table.setItem(row, 2, duplicate_item)
                self.routing_preview_table.setItem(row, 3, destination_item)
                self.routing_preview_table.setItem(row, 4, note_item)
                self.scan_status_label.setText(f"Routing preview: {row + 1} destination(s) calculated…")
                if hasattr(self, "scan_output_tabs"):
                    self.scan_output_tabs.setCurrentIndex(1)
            return

        if line.startswith("Move manifest saved:"):
            raw_path = line.split(":", 1)[1].strip()
            if raw_path:
                self.last_move_manifest_path = Path(raw_path)
            return

        if line.startswith("Routing preview saved:"):
            raw_path = line.split(":", 1)[1].strip()
            if raw_path:
                preview_path = Path(raw_path)
                self.last_routing_preview_path = preview_path
                self.last_scan_output_dir = preview_path.parent
                if hasattr(self, "scan_open_preview_btn"):
                    self.scan_open_preview_btn.setEnabled(preview_path.exists())
                if hasattr(self, "scan_open_output_btn"):
                    self.scan_open_output_btn.setEnabled(self.last_scan_output_dir.exists())
            self.scan_status_label.setText("Routing preview calculated — finishing dry run…")
            return

        if line.startswith("ROUTING PREVIEW COUNT:"):
            try:
                preview_count = int(line.split(":", 1)[1].strip())
                self.scan_status_label.setText(f"Routing preview: {preview_count} file(s) calculated")
            except Exception:
                pass
            return

        match = re.search(r"TOTAL PACKAGES TO SCAN:\s*(\d+)", line, re.IGNORECASE)
        if match:
            self.scan_total_packages = int(match.group(1))
            self.scan_progress.setRange(0, max(self.scan_total_packages + 1, 1))
            self.scan_progress_label.setText(f"0 / {self.scan_total_packages}")
            self.scan_status_label.setText("Scanner running…")
            return

        match = re.match(r"Scanning\s+(\d+)\s*/\s*(\d+)\s*:\s*(.+)", line, re.IGNORECASE)
        if match:
            current = int(match.group(1))
            total = int(match.group(2))
            filename = match.group(3).strip()
            self.scan_current_package = current
            self.scan_total_packages = total
            self.scan_progress.setRange(0, max(total + 1, 1))
            self.scan_progress.setValue(current)
            self.scan_progress_label.setText(f"{current} / {total}")
            self.scan_current_file_label.setText(filename)
            self.scan_status_label.setText("Scanning packages…")
            return

        if line.startswith("MOVED:") or line.startswith("MOVED AFTER RETRY:"):
            moved_path = line.split(":", 1)[1].strip() if ":" in line else ""
            if moved_path:
                self.scan_status_label.setText(f"Routing files… {Path(moved_path).name}")
            return

        if line.startswith("MASTER DB: starting update"):
            self.scan_progress.setRange(0, 0)
            self.scan_progress_label.setText("Finalising…")
            self.scan_status_label.setText("Updating Master Database…")
        elif line.startswith("MASTER DB: saving workbook"):
            self.scan_progress.setRange(0, 0)
            self.scan_progress_label.setText("Finalising…")
            self.scan_status_label.setText("Saving Master Database — do not close the app…")
        elif line.startswith("MASTER DB: save complete"):
            self.scan_progress.setRange(0, 0)
            self.scan_progress_label.setText("Finalising…")
            self.scan_status_label.setText("Master Database saved — writing results…")
        elif line.startswith("MASTER DB: DRY RUN - update skipped"):
            self.scan_status_label.setText("Dry run: MDB write skipped — finishing…")
        elif line.startswith("Finished! Scan summary saved:"):
            if self.scan_run_kind == "preview":
                self.scan_status_label.setText("Routing preview finished…")
            else:
                self.scan_status_label.setText("Scan finished — refreshing Library…")
            raw_path = line.split(":", 1)[1].strip()
            if raw_path:
                summary_path = Path(raw_path)
                self.last_scan_output_dir = summary_path.parent
                if hasattr(self, "scan_open_output_btn"):
                    self.scan_open_output_btn.setEnabled(self.last_scan_output_dir.exists())
                candidate_log = self.last_scan_output_dir / "Health_Scan_Log.txt"
                if candidate_log.exists():
                    self.last_scan_log_path = candidate_log
                    if hasattr(self, "scan_open_log_btn"):
                        self.scan_open_log_btn.setEnabled(True)
        elif line.startswith("Console log saved:"):
            raw_path = line.split(":", 1)[1].strip()
            if raw_path:
                self.last_scan_log_path = Path(raw_path)
                if self.last_scan_log_path.parent.exists():
                    self.last_scan_output_dir = self.last_scan_log_path.parent
                if hasattr(self, "scan_open_log_btn"):
                    self.scan_open_log_btn.setEnabled(self.last_scan_log_path.exists())
                if hasattr(self, "scan_open_output_btn") and self.last_scan_output_dir:
                    self.scan_open_output_btn.setEnabled(self.last_scan_output_dir.exists())

    def _update_scan_stats_from_line(self, line: str):
        if not line:
            return
        stats = self.scan_stats

        m = re.search(r"TOTAL PACKAGES TO SCAN:\s*(\d+)", line, re.IGNORECASE)
        if m:
            stats["scanned"] = int(m.group(1))

        if line.startswith("FINAL SUMMARY:"):
            stats["summary_records"] = int(stats.get("summary_records", 0) or 0) + 1
            m = re.search(r"\s(EXTREME|HIGH|MODERATE|SAFE)\s+\d+\s*/\s*\d+\s*$", line, re.IGNORECASE)
            if m:
                key = "risk_" + m.group(1).lower()
                stats[key] = int(stats.get(key, 0) or 0) + 1

        m = re.search(
            r"THUMBNAIL SUMMARY:\s*extracted=(\d+)\s+reused=(\d+)\s+no dedicated preview=(\d+)\s+errors=(\d+)",
            line,
            re.IGNORECASE,
        )
        if m:
            stats["thumb_extracted"] = int(m.group(1))
            stats["thumb_reused"] = int(m.group(2))
            stats["thumb_no_preview"] = int(m.group(3))
            stats["thumb_errors"] = int(m.group(4))

        m = re.search(r"EXACT SHA256 DUPLICATES ROUTED:\s*(\d+)", line, re.IGNORECASE)
        if m:
            stats["duplicates_routed"] = int(m.group(1))

        if line.startswith("MOVED:") or line.startswith("MOVED AFTER RETRY:"):
            stats["moved"] = int(stats.get("moved", 0) or 0) + 1
        if line.startswith("SKIPPED DUPLICATE:"):
            stats["skipped_existing"] = int(stats.get("skipped_existing", 0) or 0) + 1

        m = re.search(r"MASTER DB:\s*starting update for\s*(\d+)\s*packages\s*/\s*(\d+)\s*CASP records", line, re.IGNORECASE)
        if m:
            # Fallback/input counts. A newer scanner core later prints the
            # number of distinct workbook rows actually touched; that later
            # message deliberately overwrites these values.
            stats["mdb_packages"] = int(m.group(1))
            stats["mdb_casp"] = int(m.group(2))

        m = re.search(
            r"MASTER DB:\s*distinct rows updated:\s*Packages=(\d+)\s*\|\s*CASP Master=(\d+)",
            line,
            re.IGNORECASE,
        )
        if m:
            stats["mdb_packages"] = int(m.group(1))
            stats["mdb_casp"] = int(m.group(2))

        m = re.search(r"ROUTING PREVIEW COUNT:\s*(\d+)", line, re.IGNORECASE)
        if m:
            stats["preview_routes"] = int(m.group(1))

        if line.lower().startswith("scan time:"):
            stats["duration"] = line.split(":", 1)[1].strip()

    def _build_scan_completion_summary(self) -> str:
        stats = self.scan_stats
        if self.scan_run_kind == "preview":
            title = "Routing Preview completed"
        elif self.scan_run_kind == "routing":
            title = "Routing completed"
        else:
            title = "Scan completed"

        lines = [title, ""]
        scanned_count = max(int(stats.get("scanned", 0) or 0), int(stats.get("summary_records", 0) or 0))
        lines.append((f"Packages scanned: {scanned_count:,}" if self.scan_run_kind == "preview" else f"Files scanned: {scanned_count:,}"))
        if self.scan_run_kind == "preview":
            lines.append(f"Routes previewed: {int(stats.get('preview_routes', 0) or 0):,} (includes pass-through non-package files)")
        else:
            if not IS_PUBLIC_BUILD:
                lines.append(
                    "Risk: "
                    f"Safe {int(stats.get('risk_safe', 0) or 0):,} | "
                    f"Moderate {int(stats.get('risk_moderate', 0) or 0):,} | "
                    f"High {int(stats.get('risk_high', 0) or 0):,} | "
                    f"Extreme {int(stats.get('risk_extreme', 0) or 0):,}"
                )
            lines.append(
                "Thumbnails: "
                f"extracted {int(stats.get('thumb_extracted', 0) or 0):,} | "
                f"reused {int(stats.get('thumb_reused', 0) or 0):,} | "
                f"no preview {int(stats.get('thumb_no_preview', 0) or 0):,} | "
                f"errors {int(stats.get('thumb_errors', 0) or 0):,}"
            )
            lines.append(f"Duplicates routed: {int(stats.get('duplicates_routed', 0) or 0):,}")
            if self.scan_run_kind == "routing":
                lines.append(f"Files moved: {int(stats.get('moved', 0) or 0):,}")
                skipped_existing = int(stats.get('skipped_existing', 0) or 0)
                if skipped_existing:
                    lines.append(f"Existing destination files skipped: {skipped_existing:,}")
            lines.append(
                f"MDB rows updated: Packages {int(stats.get('mdb_packages', 0) or 0):,} | "
                f"CASP Master {int(stats.get('mdb_casp', 0) or 0):,}"
            )

        if stats.get("duration"):
            lines.append(f"Duration: {stats['duration']}")

        lines.append("")
        if self.last_scan_backup_path:
            lines.append(f"MDB backup: {self.last_scan_backup_path}")
        elif self.scan_run_kind == "preview":
            lines.append("MDB backup: not needed (dry run; MDB unchanged)")
        if self.last_move_manifest_path:
            lines.append(f"Move manifest: {self.last_move_manifest_path}")
        if self.last_scan_log_path:
            lines.append(f"Log: {self.last_scan_log_path}")
        if self.last_scan_output_dir:
            lines.append(f"Results: {self.last_scan_output_dir}")
        if self.last_sorted_root and self.scan_run_kind == "routing":
            lines.append(f"Sorted files: {self.last_sorted_root}")
        if self.scan_run_kind == "routing" and self.last_routing_location_tracking:
            tracking = self.last_routing_location_tracking
            tracked = int(tracking.get("tracked", 0) or 0)
            skipped = int(tracking.get("skipped", 0) or 0)
            new_locations = int(tracking.get("new_locations", 0) or 0)
            if tracked:
                lines.append(
                    f"Package locations tracked: {tracked:,} | "
                    f"new File Locations {new_locations:,} | skipped {skipped:,}"
                )
            elif tracking.get("error"):
                lines.append(f"Package location tracking: WARNING — {tracking.get('error')}")
            elif skipped:
                lines.append(f"Package location tracking: 0 tracked | skipped {skipped:,}")
            backup = str(tracking.get("backup", "") or "").strip()
            if backup:
                lines.append(f"Location-tracking MDB backup: {backup}")
        return "\n".join(lines)

    def _open_scan_backup(self):
        path = self.last_scan_backup_path
        if not path or not path.exists():
            QMessageBox.information(self, "MDB backup", "No pre-scan MDB backup is available for the latest run.")
            return
        folder = path.parent
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder))):
            QMessageBox.warning(self, "MDB backup", f"Could not open:\n{folder}")

    def _open_move_manifest(self):
        path = self.last_move_manifest_path
        if not path or not path.exists():
            QMessageBox.information(self, "Move manifest", "No move manifest is available for the latest run.")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            QMessageBox.warning(self, "Move manifest", f"Could not open:\n{path}")

    def _copy_scan_completion_summary(self, button=None):
        text = self.last_scan_summary_text or self._build_scan_completion_summary()
        QApplication.clipboard().setText(text)
        self.statusBar().showMessage("Scan completion summary copied to clipboard", 3000)
        if button is not None:
            original = button.text()
            button.setText("Copied ✓")
            button.setEnabled(False)
            QTimer.singleShot(1100, lambda b=button, t=original: (b.setText(t), b.setEnabled(True)))

    def _set_scan_progress_complete(self):
        self.scan_progress.setRange(0, 100)
        self.scan_progress.setValue(100)
        self.scan_progress_label.setText("Complete")
        self.scan_current_file_label.setText("Complete")
        if self.scan_run_kind == "preview":
            self.scan_status_label.setText("Routing preview complete")
        elif self.scan_run_kind == "routing":
            self.scan_status_label.setText("Routing complete")
        else:
            self.scan_status_label.setText("Scan complete")

    def _finalise_scan_without_popup(self):
        """Finish a scan quietly in the embedded Scan page.

        The older large completion window duplicated the embedded Scanner Log and
        made completion feel more disruptive than useful. Keep the summary in the
        embedded log and leave Open Log / Expand Log available on demand.
        """
        self.last_scan_summary_text = self._build_scan_completion_summary()
        if hasattr(self, "scan_log"):
            self.scan_log.appendPlainText("\n" + self.last_scan_summary_text)

        refresh_target = self._post_scan_refresh_target
        self._post_scan_refresh_target = None
        if refresh_target and Path(refresh_target).is_file():
            self.scan_progress.setRange(0, 0)
            self.scan_progress_label.setText("Refreshing Library…")
            self.scan_current_file_label.setText("Refreshing Library…")
            self.scan_status_label.setText("Scan work finished — refreshing Library…")
            self._scan_completion_waiting_for_library = True
            self.scan_log.appendPlainText("GUI: refreshing the open Library now…")
            QTimer.singleShot(100, lambda p=Path(refresh_target): self.load_file(p))
        else:
            self._scan_completion_waiting_for_library = False
            self._set_scan_progress_complete()

    def _show_scan_completion_dialog(self):
        self.last_scan_summary_text = self._build_scan_completion_summary()
        dialog = QDialog(self)
        dialog.setWindowTitle(
            "Routing Preview Complete" if self.scan_run_kind == "preview"
            else ("Routing Complete" if self.scan_run_kind == "routing" else "Scan Complete")
        )
        dialog.setModal(True)
        dialog.setMinimumSize(720, 450)
        dialog.setSizeGripEnabled(True)
        dialog.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        dialog.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
        dialog.resize(900, 600)
        layout = QVBoxLayout(dialog)

        heading = QLabel(dialog.windowTitle())
        heading.setStyleSheet("font-size: 14pt; font-weight: 700;")
        layout.addWidget(heading)

        summary = QPlainTextEdit()
        summary.setReadOnly(True)
        summary.setMinimumSize(500, 280)
        summary.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        summary.setPlainText(self.last_scan_summary_text)
        layout.addWidget(summary, 1)

        row1 = QHBoxLayout()
        if self.last_scan_log_path and self.last_scan_log_path.exists():
            btn = QPushButton("Open Log")
            btn.clicked.connect(self._open_last_scan_log)
            row1.addWidget(btn)
        if self.last_scan_output_dir and self.last_scan_output_dir.exists():
            btn = QPushButton("Open Results")
            btn.clicked.connect(self._open_results_folder)
            row1.addWidget(btn)
        if self.scan_run_kind == "routing" and self.last_sorted_root and self.last_sorted_root.exists():
            btn = QPushButton("Open Sorted Folder")
            btn.clicked.connect(self._open_sorted_folder)
            row1.addWidget(btn)
        if self.last_scan_backup_path and self.last_scan_backup_path.exists():
            btn = QPushButton("Open MDB Backup")
            btn.clicked.connect(self._open_scan_backup)
            row1.addWidget(btn)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        if self.last_move_manifest_path and self.last_move_manifest_path.exists():
            btn = QPushButton("Open Move Manifest")
            btn.clicked.connect(self._open_move_manifest)
            row2.addWidget(btn)
        if self.last_routing_preview_path and self.last_routing_preview_path.exists():
            btn = QPushButton("Open Routing Preview")
            btn.clicked.connect(self._open_last_routing_preview)
            row2.addWidget(btn)
        copy_btn = QPushButton("Copy Summary")
        copy_btn.clicked.connect(lambda _checked=False, b=copy_btn: self._copy_scan_completion_summary(b))
        row2.addWidget(copy_btn)
        row2.addStretch(1)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        row2.addWidget(close_btn)
        layout.addLayout(row2)

        dialog.exec()

        refresh_target = self._post_scan_refresh_target
        self._post_scan_refresh_target = None
        if refresh_target and Path(refresh_target).is_file():
            self.scan_log.appendPlainText("GUI: refreshing the open Library now…")
            QTimer.singleShot(100, lambda p=Path(refresh_target): self.load_file(p))

    def _sha256_file(self, path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    def _routing_location_name(self, folder: Path, routing_root: Path, registry: List[tuple[str, str]]) -> str:
        """Return a stable friendly File Location name for an exact routed folder.

        Existing registry entries that already point to the folder always win.
        New auto-created names include the selected routing-root name plus the
        relative routed subfolder so Manage Locations stays understandable.
        """
        try:
            folder_resolved = folder.resolve()
        except Exception:
            folder_resolved = folder.absolute()

        for saved_name, saved_path in registry:
            try:
                if Path(saved_path).resolve() == folder_resolved:
                    return str(saved_name).strip()
            except Exception:
                continue

        try:
            root_resolved = routing_root.resolve()
        except Exception:
            root_resolved = routing_root.absolute()
        try:
            rel = folder_resolved.relative_to(root_resolved)
            rel_text = str(rel).strip()
        except Exception:
            rel_text = folder_resolved.name

        root_label = routing_root.name.strip() or "Routing"
        base = f"{root_label}: {rel_text}" if rel_text not in {"", "."} else f"{root_label}: root"
        existing_names = {str(name).casefold(): str(path) for name, path in registry}
        candidate = base
        counter = 2
        while candidate.casefold() in existing_names:
            try:
                if Path(existing_names[candidate.casefold()]).resolve() == folder_resolved:
                    return candidate
            except Exception:
                pass
            candidate = f"{base} ({counter})"
            counter += 1
        return candidate

    def _track_package_paths_in_mdb(
        self,
        target_mdb: Path,
        package_paths: List[Path],
        location_root: Path,
        log_label: str,
    ) -> Dict[str, object]:
        """Track current physical package folders in File Location, fail-closed."""
        fresh = MDBWorkbook().load(target_mdb)
        records_by_name: Dict[str, List[Dict[str, str]]] = {}
        for rec in fresh.records:
            package_name = Path(str(rec.get("Package", "") or "").strip()).name
            if package_name:
                records_by_name.setdefault(package_name.casefold(), []).append(rec)

        proposed_by_row: Dict[int, List[Path]] = {}
        skipped = 0
        skip_reasons: List[str] = []
        for current_path in package_paths:
            current = Path(current_path)
            if current.suffix.lower() != ".package" or not current.is_file():
                continue
            digest = self._sha256_file(current).lower()
            candidates = list(records_by_name.get(current.name.casefold(), []))
            if not candidates:
                skipped += 1
                skip_reasons.append(f"no MDB row: {current.name}")
                continue

            hash_matches = [
                rec for rec in candidates
                if str(rec.get("Package Hash", "") or "").strip().lower() == digest
            ]
            if len(hash_matches) == 1:
                chosen = hash_matches[0]
            elif len(candidates) == 1:
                stored_hash = str(candidates[0].get("Package Hash", "") or "").strip().lower()
                if stored_hash and stored_hash != digest:
                    skipped += 1
                    skip_reasons.append(f"hash mismatch: {current.name}")
                    continue
                chosen = candidates[0]
            else:
                skipped += 1
                skip_reasons.append(f"ambiguous MDB rows: {current.name}")
                continue

            try:
                excel_row = int(chosen.get("_ExcelRow", "0") or 0)
            except (TypeError, ValueError):
                excel_row = 0
            if excel_row <= 1:
                skipped += 1
                skip_reasons.append(f"missing MDB row number: {current.name}")
                continue
            proposed_by_row.setdefault(excel_row, []).append(current.parent)

        # One MDB row must resolve to exactly one current folder. Exact duplicate
        # copies can otherwise point the same logical row at two physical files.
        row_destinations: Dict[int, Path] = {}
        for excel_row, parents in proposed_by_row.items():
            unique: List[Path] = []
            seen = set()
            for parent in parents:
                try:
                    key = str(parent.resolve()).casefold()
                except Exception:
                    key = str(parent.absolute()).casefold()
                if key not in seen:
                    seen.add(key)
                    unique.append(parent)
            if len(unique) != 1:
                skipped += 1
                skip_reasons.append(f"multiple physical copies for MDB row {excel_row}")
                continue
            row_destinations[excel_row] = unique[0]

        registry = list(getattr(fresh, "file_locations", []) or [])
        original_registry_len = len(registry)
        changes: Dict[int, Dict[str, str]] = {}
        for excel_row, parent in row_destinations.items():
            friendly = self._routing_location_name(parent, location_root, registry)
            if not any(str(name).casefold() == friendly.casefold() for name, _path in registry):
                registry.append((friendly, str(parent.resolve())))
            changes[excel_row] = {"File Location": friendly}

        if not changes:
            self.scan_log.appendPlainText(
                f"GUI: {log_label} File Location tracking made no MDB changes; skipped {skipped} package mapping(s)."
            )
            return {
                "tracked": 0,
                "skipped": skipped,
                "new_locations": 0,
                "reasons": skip_reasons[:12],
            }

        result = MDBWorkbook().save_manual_changes(
            fresh,
            changes,
            file_locations=registry if len(registry) != original_registry_len else None,
        )
        tracked = {
            "tracked": len(changes),
            "skipped": skipped,
            "new_locations": len(registry) - original_registry_len,
            "backup": str(result.backup_path),
            "reasons": skip_reasons[:12],
        }
        self.scan_log.appendPlainText(
            f"GUI: {log_label} File Location tracking updated {len(changes)} package row(s) "
            f"and added {len(registry) - original_registry_len} exact folder location(s)."
        )
        if skipped:
            self.scan_log.appendPlainText(
                f"GUI: {log_label} File Location tracking skipped {skipped} ambiguous/unmatched mapping(s); no guesses were made."
            )
        self.scan_log.appendPlainText(f"GUI: {log_label} File Location MDB backup: {result.backup_path}")
        return tracked

    def _post_routing_track_package_locations(self, target_mdb: Path) -> None:
        """Record exact post-routing package folders from the completed manifest."""
        self.last_routing_location_tracking = {}
        manifest = self.last_move_manifest_path
        routing_root = self.last_sorted_root
        if not manifest or not manifest.is_file() or not routing_root:
            return
        try:
            lines = manifest.read_text(encoding="utf-8", errors="replace").splitlines()
            destinations: List[Path] = []
            for line in lines[1:]:
                parts = line.split("\t")
                if len(parts) < 4:
                    continue
                _source_text, destination_text, _category, status = parts[:4]
                if status not in {"MOVED", "MOVED_AFTER_RETRY"}:
                    continue
                destination = Path(destination_text)
                if destination.suffix.lower() == ".package" and destination.is_file():
                    destinations.append(destination)
            self.last_routing_location_tracking = self._track_package_paths_in_mdb(
                target_mdb,
                destinations,
                routing_root,
                "post-routing",
            )
        except Exception as exc:
            self.last_routing_location_tracking = {"error": str(exc), "tracked": 0, "skipped": 0, "new_locations": 0}
            self.scan_log.appendPlainText(
                f"GUI WARNING: routing succeeded, but post-routing File Location tracking could not be saved: {exc}"
            )

    def _prepare_routing_recovery_manifest(self):
        """Create a fail-closed hash manifest for the latest successful routing run.

        The shared scanner core's text Move_Manifest records source/destination
        paths. PUBLIC recovery adds SHA-256 fingerprints *after* the move so a
        later recovery can verify every destination before moving anything back.
        """
        manifest = self.last_move_manifest_path
        if not manifest or not manifest.is_file():
            self.scan_log.appendPlainText("GUI: routing completed but no move manifest was found; recovery is unavailable.")
            return
        try:
            lines = manifest.read_text(encoding="utf-8", errors="replace").splitlines()
            records = []
            for line in lines[1:]:
                parts = line.split("\t")
                if len(parts) < 4:
                    continue
                source_text, destination_text, category, status = parts[:4]
                if status not in {"MOVED", "MOVED_AFTER_RETRY"}:
                    continue
                source = Path(source_text)
                destination = Path(destination_text)
                if not destination.is_file():
                    raise RuntimeError(f"Routed file is missing: {destination}")
                records.append({
                    "source": str(source),
                    "destination": str(destination),
                    "category": category,
                    "sha256": self._sha256_file(destination),
                })
            if not records:
                self.scan_log.appendPlainText("GUI: no successfully moved files were found in the move manifest; recovery is not needed.")
                return
            payload = {
                "schema": "CASCuratorPublicRecovery/v1",
                "created": datetime.now().isoformat(timespec="seconds"),
                "move_manifest": str(manifest),
                "mdb_path": str(Path(self.scan_mdb_path).resolve()) if self.scan_mdb_path else "",
                "source_root": self.scan_source_edit.text().strip() if hasattr(self, "scan_source_edit") else "",
                "records": records,
            }
            recovery_path = manifest.with_name("CAS_Curator_Recovery_Manifest.json")
            recovery_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
            self.last_recovery_manifest_path = recovery_path
            if hasattr(self, "scan_recover_routing_btn"):
                self.scan_recover_routing_btn.setEnabled(True)
            self.scan_log.appendPlainText(f"GUI: hash-verified recovery manifest prepared: {recovery_path}")
        except Exception as exc:
            self.last_recovery_manifest_path = None
            if hasattr(self, "scan_recover_routing_btn"):
                self.scan_recover_routing_btn.setEnabled(False)
            self.scan_log.appendPlainText(f"GUI: recovery manifest could not be prepared: {exc}")

    def _recover_last_routing(self):
        path = self.last_recovery_manifest_path
        if not path or not path.is_file():
            QMessageBox.information(self, "Recover Last Routing", "No unrecovered routing manifest is available in this session.")
            return
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            records = list(payload.get("records", []))
        except Exception as exc:
            QMessageBox.critical(self, "Recover Last Routing", f"Could not read the recovery manifest:\n{exc}")
            return
        if not records:
            QMessageBox.information(self, "Recover Last Routing", "The recovery manifest contains no moved files.")
            return

        recovery_mdb_text = str(payload.get("mdb_path", "") or "").strip()
        recovery_mdb = Path(recovery_mdb_text) if recovery_mdb_text else (Path(self.scan_mdb_path) if self.scan_mdb_path else None)
        if self.mdb and recovery_mdb:
            try:
                same_open_mdb = self.mdb.path.resolve() == recovery_mdb.resolve()
            except Exception:
                same_open_mdb = False
            if same_open_mdb and not self._confirm_pending_changes():
                return

        # Fail closed: verify every row before moving the first file.
        failures = []
        verified = []
        for item in records:
            source = Path(str(item.get("source", "")))
            destination = Path(str(item.get("destination", "")))
            expected_hash = str(item.get("sha256", ""))
            if source.exists():
                failures.append(f"Original path already exists: {source}")
                continue
            if not destination.is_file():
                failures.append(f"Routed file is missing: {destination}")
                continue
            try:
                actual_hash = self._sha256_file(destination)
            except Exception as exc:
                failures.append(f"Could not hash {destination}: {exc}")
                continue
            if not expected_hash or actual_hash.lower() != expected_hash.lower():
                failures.append(f"Hash mismatch: {destination}")
                continue
            verified.append((source, destination))
        if failures:
            QMessageBox.critical(
                self,
                "Recovery blocked",
                "Recovery stopped before moving any files because a safety check failed:\n\n" + "\n".join(failures[:12]),
            )
            return

        answer = QMessageBox.question(
            self,
            "Recover Last Routing?",
            "CAS Curator will move the routed files back to their recorded original paths after SHA-256 verification.\n\n"
            "Nothing will be overwritten. Recovery stops if an original path already exists or a hash differs.\n\n"
            f"Files to recover: {len(verified):,}\n\nContinue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return

        moved_back = []
        try:
            for source, destination in verified:
                source.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(destination), str(source))
                moved_back.append((source, destination))
        except Exception as exc:
            # Best-effort rollback of files already restored in this attempt.
            rollback_errors = []
            for source, destination in reversed(moved_back):
                try:
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    if source.exists() and not destination.exists():
                        shutil.move(str(source), str(destination))
                except Exception as rollback_exc:
                    rollback_errors.append(str(rollback_exc))
            detail = f"Recovery failed and was stopped:\n{exc}"
            if rollback_errors:
                detail += "\n\nSome rollback operations also failed:\n" + "\n".join(rollback_errors[:8])
            QMessageBox.critical(self, "Recovery failed", detail)
            return

        recovery_tracking = {}
        if recovery_mdb and recovery_mdb.is_file():
            try:
                restored_packages = [source for source, _destination in moved_back if source.suffix.lower() == ".package" and source.is_file()]
                source_root_text = str(payload.get("source_root", "") or "").strip()
                if source_root_text:
                    source_root = Path(source_root_text)
                elif restored_packages:
                    try:
                        source_root = Path(os.path.commonpath([str(item.parent) for item in restored_packages]))
                    except Exception:
                        source_root = restored_packages[0].parent
                else:
                    source_root = Path.home()
                if restored_packages:
                    recovery_tracking = self._track_package_paths_in_mdb(
                        recovery_mdb, restored_packages, source_root, "post-recovery"
                    )
            except Exception as exc:
                recovery_tracking = {"error": str(exc)}
                self.scan_log.appendPlainText(
                    f"GUI WARNING: files were recovered, but File Location tracking could not be restored: {exc}"
                )

        recovered_path = path.with_name(path.stem + "_RECOVERED.json")
        try:
            path.replace(recovered_path)
        except Exception:
            recovered_path = path
        self.last_recovery_manifest_path = None
        if hasattr(self, "scan_recover_routing_btn"):
            self.scan_recover_routing_btn.setEnabled(False)
        self.scan_log.appendPlainText(f"\nROUTING RECOVERY COMPLETED: {len(moved_back):,} file(s) restored.")
        extra = ""
        if recovery_tracking.get("tracked"):
            extra = f"\n\nFile Location restored for {int(recovery_tracking.get('tracked', 0)):,} package row(s)."
        elif recovery_tracking.get("error"):
            extra = "\n\nThe files were recovered, but File Location tracking could not be updated; see Scanner Log."
        QMessageBox.information(self, "Routing recovered", f"Recovered {len(moved_back):,} file(s) to their original paths.{extra}")
        if self.mdb and recovery_mdb and recovery_mdb.is_file():
            try:
                if self.mdb.path.resolve() == recovery_mdb.resolve():
                    QTimer.singleShot(100, lambda p=recovery_mdb: self.load_file(p))
            except Exception:
                pass

    def _open_scan_log_window(self):
        """Show the current scanner log in a larger, independently resizable window."""
        existing = getattr(self, "_scan_log_dialog", None)
        if existing is not None:
            try:
                if existing.isVisible():
                    existing.showNormal()
                    existing.raise_()
                    existing.activateWindow()
                    return
            except RuntimeError:
                self._scan_log_dialog = None

        # RC4 retains the standalone top-level QMainWindow for Expand Log; the embedded output area now also has its own vertical splitter.
        # The previous QDialog variants still produced a non-draggable resize border
        # on the Windows test machine, so avoid dialog-size-grip/window-flag behaviour
        # entirely and let Windows manage the native resizable main-window frame.
        dialog = QMainWindow()
        dialog.setObjectName("ThemedUtilityWindow")
        dialog.setWindowTitle("CAS Curator — Scanner Log")
        dialog.setWindowIcon(self.windowIcon())
        dialog.setAttribute(Qt.WA_DeleteOnClose, True)
        dialog.setStyleSheet(self._dialog_stylesheet())
        dialog.setMinimumSize(460, 300)
        dialog.resize(1000, 680)

        central = QWidget(dialog)
        central.setObjectName("ThemedUtilityPanel")
        dialog.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        heading = QLabel("Scanner Log", central)
        heading_font = heading.font()
        heading_font.setPointSizeF(12.0)
        heading_font.setBold(True)
        heading.setFont(heading_font)
        layout.addWidget(heading)
        log_view = QPlainTextEdit(central)
        log_view.setObjectName("ExpandedScannerLog")
        log_view.setReadOnly(True)
        log_view.setLineWrapMode(QPlainTextEdit.NoWrap)
        log_view.setMinimumSize(260, 160)
        log_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        log_view.setPlainText(self.scan_log.toPlainText() if hasattr(self, "scan_log") else "")
        layout.addWidget(log_view, 1)

        buttons = QHBoxLayout()
        copy_btn = QPushButton("Copy Log", central)
        copy_btn.clicked.connect(lambda _checked=False: QApplication.clipboard().setText(log_view.toPlainText()))
        buttons.addWidget(copy_btn)
        buttons.addStretch(1)
        close_btn = QPushButton("Close", central)
        close_btn.clicked.connect(dialog.close)
        buttons.addWidget(close_btn)
        layout.addLayout(buttons)

        # Keep the pop-out current while a scan is running, without forcing the
        # user's scroll position back to the end when they inspect older lines.
        def refresh_log_view():
            scrollbar = log_view.verticalScrollBar()
            old_value = scrollbar.value()
            was_at_bottom = old_value >= scrollbar.maximum() - 1
            log_view.setPlainText(self.scan_log.toPlainText())
            scrollbar.setValue(scrollbar.maximum() if was_at_bottom else min(old_value, scrollbar.maximum()))

        self.scan_log.textChanged.connect(refresh_log_view)

        def cleanup_log_window(*_args):
            try:
                self.scan_log.textChanged.disconnect(refresh_log_view)
            except (RuntimeError, TypeError):
                pass
            if getattr(self, "_scan_log_dialog", None) is dialog:
                self._scan_log_dialog = None

        dialog.destroyed.connect(cleanup_log_window)
        self._scan_log_dialog = dialog
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def _copy_scan_log(self):
        text = self.scan_log.toPlainText() if hasattr(self, "scan_log") else ""
        QApplication.clipboard().setText(text)
        self.statusBar().showMessage("Scanner log copied to clipboard", 3000)
        if hasattr(self, "scan_copy_log_btn"):
            btn = self.scan_copy_log_btn
            btn.setDown(True)
            btn.setText("Copied ✓")
            btn.setEnabled(False)
            QTimer.singleShot(140, lambda b=btn: b.setDown(False))
            QTimer.singleShot(1200, lambda b=btn: self._reset_copy_log_button(b))

    def _reset_copy_log_button(self, button):
        if button is not None:
            button.setText("Copy Log")
            button.setEnabled(True)

    def _copy_filename_text(self, filename: str, source_button=None):
        filename = (filename or "").strip()
        if not filename:
            self.statusBar().showMessage("No file name is selected", 2500)
            return
        QApplication.clipboard().setText(filename)
        self.statusBar().showMessage(f"Copied file name: {filename}", 3000)
        if source_button is not None:
            original = source_button.text()
            source_button.setDown(True)
            source_button.setText("Copied ✓")
            source_button.setEnabled(False)
            QTimer.singleShot(140, lambda b=source_button: b.setDown(False))
            QTimer.singleShot(1100, lambda b=source_button, t=original: self._reset_copy_filename_button(b, t))

    def _reset_copy_filename_button(self, button, text: str):
        if button is not None:
            button.setText(text)
            button.setEnabled(self.current_source_row is not None)

    def _copy_current_filename(self):
        if self.current_source_row is None:
            self.statusBar().showMessage("Select a package first", 2500)
            return
        rec = self.model.record(self.current_source_row)
        self._copy_filename_text(rec.get("Package", ""), self.copy_filename_btn)

    def _show_table_context_menu(self, pos):
        proxy_index = self.table.indexAt(pos)
        if not proxy_index.isValid():
            return
        selection = self.table.selectionModel()
        already_selected = bool(selection and selection.isRowSelected(proxy_index.row(), QModelIndex()))
        if not already_selected:
            self.table.selectRow(proxy_index.row())
        source_index = self.proxy.mapToSource(proxy_index)
        rec = self.model.record(source_index.row())
        if not rec:
            return
        filename = rec.get("Package", "")
        try:
            excel_row = int(rec.get("_ExcelRow", "0") or 0)
        except (TypeError, ValueError):
            excel_row = 0
        menu = QMenu(self.table)
        file_location_name = self.model.effective_value(source_index.row(), "File Location").strip()
        open_location_action = None
        if file_location_name:
            open_location_action = menu.addAction("Open File Location")
        creator = self.model.effective_value(source_index.row(), "Creator").strip()
        search_web_action = menu.addAction("Search Web for Package")
        search_web_action.setEnabled(bool(str(filename or "").strip()))
        search_creator_action = None
        if creator:
            search_creator_action = menu.addAction("Search Web for Creator")
        menu.addSeparator()
        selected_rows = self._selected_source_rows() or [source_index.row()]
        selected_count = len(selected_rows)
        set_creator_action = menu.addAction(
            "Set Creator…" if selected_count == 1 else f"Set Creator for {selected_count} Selected…"
        )
        move_location_action = menu.addAction(
            "Move Package to Location…" if selected_count == 1 else f"Move {selected_count} Selected Packages to Location…"
        )
        copy_dropin_action = menu.addAction(
            "Copy Package to CC Magic Import…" if selected_count == 1 else f"Copy {selected_count} Selected Packages to CC Magic Import…"
        )
        menu.addSeparator()
        copy_action = menu.addAction("Copy file name")
        menu.addSeparator()
        delete_action = menu.addAction("Delete Record…")
        delete_action.setEnabled(bool(self.mdb and excel_row > 1 and excel_row not in self.pending_deletions))
        chosen = menu.exec(self.table.viewport().mapToGlobal(pos))
        if open_location_action is not None and chosen == open_location_action:
            self._open_file_location_for_source_row(source_index.row())
        elif chosen == search_web_action:
            self._search_package_web(filename, creator)
        elif search_creator_action is not None and chosen == search_creator_action:
            self._search_creator_web(creator)
        elif chosen == set_creator_action:
            self.set_creator_for_selected(selected_rows)
        elif chosen == move_location_action:
            self.move_selected_to_location(selected_rows)
        elif chosen == copy_dropin_action:
            self.copy_selected_to_cc_magic_dropin(selected_rows)
        elif chosen == copy_action:
            self._copy_filename_text(filename)
        elif chosen == delete_action:
            self._stage_delete_record(rec)

    def set_creator_for_selected(self, source_rows: Optional[List[int]] = None):
        """Stage one manual Creator value across the selected Library rows.

        This is deliberately a Library edit rather than scanner inference: the
        creator is supplied by the user, written only through the normal pending
        edit/save path, and remains fully revertible until Save Changes.
        """
        if not self.mdb:
            QMessageBox.information(self, "Set Creator", "Open a Master Database first.")
            return

        rows = list(source_rows or self._selected_source_rows())
        if not rows and self.current_source_row is not None:
            rows = [self.current_source_row]

        eligible: List[int] = []
        seen = set()
        skipped_deletions = 0
        for source_row in rows:
            try:
                source_row = int(source_row)
            except (TypeError, ValueError):
                continue
            if source_row in seen or not (0 <= source_row < len(self.model.records)):
                continue
            seen.add(source_row)
            rec = self.model.record(source_row)
            if not rec:
                continue
            excel_row = self.model._excel_row(rec)
            if not excel_row:
                continue
            if excel_row in self.pending_deletions:
                skipped_deletions += 1
                continue
            eligible.append(source_row)

        if not eligible:
            QMessageBox.information(self, "Set Creator", "Select one or more package rows first.")
            return

        existing_values = [self.model.effective_value(row, "Creator").strip() for row in eligible]
        nonblank_values = [value for value in existing_values if value]
        unique_nonblank = []
        seen_creator_keys = set()
        for value in nonblank_values:
            key = value.casefold()
            if key not in seen_creator_keys:
                seen_creator_keys.add(key)
                unique_nonblank.append(value)
        initial = unique_nonblank[0] if len(unique_nonblank) == 1 else ""

        creator, ok = QInputDialog.getText(
            self,
            "Set Creator for Selected",
            f"Creator for {len(eligible):,} selected package(s):",
            QLineEdit.Normal,
            initial,
        )
        if not ok:
            return
        creator = re.sub(r"\s+", " ", str(creator or "").strip())
        if not creator:
            QMessageBox.information(
                self,
                "Set Creator",
                "Creator cannot be blank in the bulk-fill action. Edit a Creator cell directly if you deliberately want to clear a value.",
            )
            return

        already_same = sum(1 for value in existing_values if value == creator)
        overwrite_nonblank = sum(1 for value in existing_values if value and value != creator)
        changing = len(eligible) - already_same
        if changing <= 0:
            self.statusBar().showMessage(f"All selected packages already use Creator '{creator}'.", 5000)
            return

        details = [
            f"Set Creator to '{creator}' for {len(eligible):,} selected package(s)?",
            f"{changing:,} row(s) will change; {already_same:,} already have this Creator.",
        ]
        if overwrite_nonblank:
            details.append(f"{overwrite_nonblank:,} existing non-empty Creator value(s) will be replaced.")
        if skipped_deletions:
            details.append(f"{skipped_deletions:,} row(s) staged for deletion will be skipped.")
        details.append("These edits remain staged until you use Save Changes.")
        answer = QMessageBox.question(
            self,
            "Set Creator for Selected",
            "\n\n".join(details),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return

        changed_rows = self.model.stage_fill_value(eligible, "Creator", creator)
        if not changed_rows:
            self.statusBar().showMessage("No Creator values needed changing.", 5000)
            return
        for source_row in changed_rows:
            self.model.notify_row_changed(source_row)
        if self.current_source_row in changed_rows:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)
        self._update_dirty_state()
        self._update_selected_dirty_label()
        self._apply_filters()
        self.statusBar().showMessage(
            f"Set Creator '{creator}' for {len(changed_rows):,} package(s) — unsaved until Save Changes",
            8000,
        )

    def _managed_selected_rows(self, source_rows: Optional[List[int]] = None) -> List[int]:
        rows = list(source_rows or self._selected_source_rows())
        if not rows and self.current_source_row is not None:
            rows = [self.current_source_row]
        result = []
        seen = set()
        for row in rows:
            try:
                row_num = int(row)
            except (TypeError, ValueError):
                continue
            if row_num < 0 or row_num in seen:
                continue
            if self.model.record(row_num):
                seen.add(row_num)
                result.append(row_num)
        return result

    @staticmethod
    def _managed_file_sha256(path: Path) -> str:
        h = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest().lower()

    def _resolve_managed_package_source(self, source_row: int) -> tuple[Path, str]:
        """Resolve one real package file conservatively, preferring the assigned physical File Location."""
        rec = self.model.record(source_row)
        if not rec:
            raise ValueError("The selected Library row is no longer available.")
        filename = Path(str(rec.get("Package", "") or "").strip()).name
        if not filename or not filename.lower().endswith(".package"):
            raise ValueError("This row does not contain a usable .package file name.")
        expected_hash = str(rec.get("Package Hash", "") or "").strip().lower()
        candidates: List[tuple[int, Path]] = []

        location_name = self.model.effective_value(source_row, "File Location").strip()
        if location_name:
            folder_path = self._file_location_path_for_name(location_name)
            if folder_path and Path(folder_path).is_dir():
                folder = Path(folder_path)
                direct = folder / filename
                if direct.is_file():
                    candidates.append((0, direct))
                else:
                    # File Locations may point to a managed root containing creator/set subfolders.
                    # A recursive match is accepted only when identity is unambiguous below.
                    try:
                        for found in folder.rglob(filename):
                            if found.is_file():
                                candidates.append((1, found))
                                if len(candidates) >= 80:
                                    break
                    except OSError:
                        pass

        for raw_line in str(rec.get("Original Source Path", "") or "").splitlines():
            raw_line = raw_line.strip()
            if not raw_line:
                continue
            candidate = Path(raw_line)
            if candidate.is_dir():
                candidate = candidate / filename
            if candidate.is_file():
                candidates.append((2, candidate))

        unique: Dict[str, tuple[int, Path]] = {}
        for priority, candidate in candidates:
            try:
                resolved = candidate.resolve()
            except OSError:
                resolved = candidate.absolute()
            key = os.path.normcase(str(resolved)).casefold()
            previous = unique.get(key)
            if previous is None or priority < previous[0]:
                unique[key] = (priority, resolved)
        candidates = list(unique.values())
        if not candidates:
            raise FileNotFoundError(
                f"Could not locate the original package '{filename}'.\n\n"
                "CC Magic Import is only the destination for a copy. CAS Curator also needs a valid File Location for the original/master package.\n\n"
                "To create a valid File Location, either:\n"
                "• route the package first (Preview Routing → Run Routing), which records its routed folder; or\n"
                "• select the package, open Edit, and assign File Location to a saved managed location that contains the original file. "
                "Use Manage Locations… in Edit first if that folder has not been added yet."
            )

        checked: List[tuple[int, Path, str]] = []
        for priority, candidate in candidates:
            try:
                actual_hash = self._managed_file_sha256(candidate)
            except OSError:
                continue
            checked.append((priority, candidate, actual_hash))

        if expected_hash:
            matching = [(priority, path, actual) for priority, path, actual in checked if actual == expected_hash]
            if not matching:
                raise ValueError(
                    f"Located file(s) named '{filename}', but none match the Package Hash stored in the MDB. No file was moved or copied."
                )
        else:
            matching = checked
        if not matching:
            raise FileNotFoundError(f"Could not read a usable source file for '{filename}'.")

        best_priority = min(item[0] for item in matching)
        best = [item for item in matching if item[0] == best_priority]
        if len(best) != 1:
            paths = "\n".join(f"• {item[1]}" for item in best[:8])
            raise ValueError(
                f"More than one equally valid copy of '{filename}' was found. CAS Curator will not guess which original to use.\n\n{paths}"
            )
        _priority, source_path, actual_hash = best[0]
        return source_path, actual_hash

    def _managed_preflight_sources(self, source_rows: List[int]) -> List[dict]:
        plans = []
        errors = []
        for source_row in source_rows:
            rec = self.model.record(source_row)
            filename = Path(str((rec or {}).get("Package", "") or "")).name or "(unnamed package)"
            try:
                source_path, source_hash = self._resolve_managed_package_source(source_row)
                excel_row = int((rec or {}).get("_ExcelRow", "0") or 0)
                if excel_row <= 1:
                    raise ValueError("The Library row has no safe MDB worksheet identity.")
                plans.append({
                    "source_row": source_row,
                    "excel_row": excel_row,
                    "filename": filename,
                    "source": source_path,
                    "hash": source_hash,
                })
            except Exception as exc:
                errors.append(f"{filename}: {exc}")
        if errors:
            shown = "\n\n".join(errors[:10])
            if len(errors) > 10:
                shown += f"\n\n…and {len(errors) - 10} more problem(s)."
            raise ValueError(shown)
        return plans

    def _rollback_managed_moves(self, moved: List[dict]) -> List[str]:
        problems = []
        for item in reversed(moved):
            source = item["source"]
            destination = item["destination"]
            expected_hash = item["hash"]
            try:
                if source.exists():
                    problems.append(f"Original path already exists during rollback: {source}")
                    continue
                if not destination.is_file():
                    problems.append(f"Moved copy is missing during rollback: {destination}")
                    continue
                shutil.copy2(destination, source)
                if self._managed_file_sha256(source) != expected_hash:
                    source.unlink(missing_ok=True)
                    problems.append(f"Hash verification failed while restoring: {source}")
                    continue
                destination.unlink()
            except Exception as exc:
                problems.append(f"{item['filename']}: {exc}")
        return problems

    def move_selected_to_location(self, source_rows: Optional[List[int]] = None):
        if not self.mdb:
            QMessageBox.information(self, "Move to Location", "Open a Master Database first.")
            return
        rows = self._managed_selected_rows(source_rows)
        if not rows:
            QMessageBox.information(self, "Move to Location", "Select one or more package rows first.")
            return
        if self.scan_process is not None and self.scan_process.state() != QProcess.NotRunning:
            QMessageBox.warning(self, "Move to Location", "Wait for the current Scan/Routing operation to finish before moving package files.")
            return
        if not self._confirm_pending_changes():
            return
        locations = [(name, path) for name, path in self.file_location_registry_current if str(name).strip() and str(path).strip()]
        if not locations:
            QMessageBox.information(
                self, "Move to Location",
                "No physical File Locations are configured yet. Use Manage Locations… in Edit to add a named folder first."
            )
            return
        names = [name for name, _path in sorted(locations, key=lambda item: item[0].casefold())]
        target_name, ok = QInputDialog.getItem(
            self, "Move to Location", "Destination File Location:", names, 0, False
        )
        if not ok or not str(target_name).strip():
            return
        target_name = str(target_name).strip()
        target_path_text = self._file_location_path_for_name(target_name)
        target_root = Path(target_path_text) if target_path_text else Path()
        if not target_path_text or not target_root.is_dir():
            QMessageBox.warning(
                self, "Move to Location",
                f"The saved folder for '{target_name}' is unavailable:\n\n{target_path_text or '(no path saved)'}"
            )
            return
        try:
            plans = self._managed_preflight_sources(rows)
        except Exception as exc:
            QMessageBox.warning(self, "Move preflight failed", str(exc))
            return

        target_keys = set()
        conflicts = []
        for item in plans:
            destination = (target_root / item["filename"]).resolve()
            item["destination"] = destination
            source_key = os.path.normcase(str(item["source"].resolve())).casefold()
            dest_key = os.path.normcase(str(destination)).casefold()
            item["same_path"] = source_key == dest_key
            if dest_key in target_keys and not item["same_path"]:
                conflicts.append(f"Two selected packages would use the same destination name: {item['filename']}")
            target_keys.add(dest_key)
            if destination.exists() and not item["same_path"]:
                try:
                    dest_hash = self._managed_file_sha256(destination)
                    kind = "an exact duplicate" if dest_hash == item["hash"] else "a different file"
                except Exception:
                    kind = "an existing file"
                conflicts.append(f"{destination} already contains {kind}.")
        if conflicts:
            QMessageBox.warning(
                self, "Move blocked",
                "Nothing was moved because the destination is not conflict-free:\n\n" + "\n".join(f"• {c}" for c in conflicts[:12])
            )
            return

        preview = [
            f"Move {len(plans):,} package(s) to '{target_name}'?",
            "",
            f"Destination folder: {target_root}",
            "",
        ]
        for item in plans[:12]:
            if item["same_path"]:
                preview.append(f"• {item['filename']} — already physically in this folder; File Location will be updated")
            else:
                preview.append(f"• {item['source']}\n  → {item['destination']}")
        if len(plans) > 12:
            preview.append(f"…and {len(plans) - 12} more package(s).")
        preview += [
            "",
            "CAS Curator copies each file first, verifies SHA-256, deletes the original only after verification, then updates File Location in the MDB. The MDB writer creates its normal backup. Existing destination files are never overwritten.",
        ]
        answer = QMessageBox.question(
            self, "Preview managed move", "\n".join(preview),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return

        moved: List[dict] = []
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            for item in plans:
                if item["same_path"]:
                    continue
                source = item["source"]
                destination = item["destination"]
                try:
                    shutil.copy2(source, destination)
                    if self._managed_file_sha256(destination) != item["hash"]:
                        destination.unlink(missing_ok=True)
                        raise ValueError("copied file failed SHA-256 verification")
                    try:
                        source.unlink()
                    except Exception:
                        destination.unlink(missing_ok=True)
                        raise
                    moved.append(item)
                except Exception as exc:
                    try:
                        if destination.exists() and source.exists():
                            destination.unlink(missing_ok=True)
                    except Exception:
                        pass
                    rollback_problems = self._rollback_managed_moves(moved)
                    detail = f"Move failed at '{item['filename']}': {exc}\n\nEarlier files were rolled back."
                    if rollback_problems:
                        detail += "\n\nRollback needs attention:\n" + "\n".join(f"• {p}" for p in rollback_problems)
                    raise RuntimeError(detail) from exc

            changes = {item["excel_row"]: {"File Location": target_name} for item in plans}
            try:
                result = MDBWorkbook().save_manual_changes(self.mdb, changes)
            except Exception as exc:
                rollback_problems = self._rollback_managed_moves(moved)
                detail = f"The files were moved, but the MDB update failed: {exc}\n\nCAS Curator attempted to restore the original file locations."
                if rollback_problems:
                    detail += "\n\nRollback needs attention:\n" + "\n".join(f"• {p}" for p in rollback_problems)
                raise RuntimeError(detail) from exc
        except Exception as exc:
            QMessageBox.critical(self, "Managed move failed", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        for item in plans:
            rec = self.model.record(item["source_row"])
            if rec:
                rec["File Location"] = target_name
                self.model.notify_row_changed(item["source_row"])
        self._populate_filters()
        self._apply_filters()
        if self.current_source_row is not None:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)
        already_in_destination = sum(1 for item in plans if item.get("same_path"))
        moved_count = len(plans) - already_in_destination
        move_lines = [f"Moved: {moved_count:,}"]
        if already_in_destination:
            move_lines.append(f"Already in destination: {already_in_destination:,}")
        move_lines += ["", "File Location was updated in the MDB.", f"Backup: {result.backup_path.name}"]
        QMessageBox.information(self, "Managed move complete", "\n".join(move_lines))
        self.statusBar().showMessage(
            f"Managed move: {moved_count:,} moved, {already_in_destination:,} already in '{target_name}'", 9000
        )

    def _cc_magic_import_signature(self):
        """Cheap folder signature used to notice manual CC Magic Import changes."""
        raw = str(self.cc_magic_dropin_path or "").strip()
        if not raw:
            return ("unavailable", "")
        root = Path(raw)
        if not root.is_dir():
            return ("unavailable", os.path.normcase(os.path.abspath(raw)).casefold())
        entries = []
        try:
            for child in root.iterdir():
                if not child.is_file() or child.suffix.lower() != ".package":
                    continue
                try:
                    stat = child.stat()
                    entries.append((child.name.casefold(), int(stat.st_size), int(stat.st_mtime_ns)))
                except OSError:
                    entries.append((child.name.casefold(), -1, -1))
        except OSError:
            return ("unavailable", os.path.normcase(os.path.abspath(raw)).casefold())
        return (os.path.normcase(os.path.abspath(raw)).casefold(), tuple(sorted(entries)))

    def _refresh_cc_magic_import_statuses_if_changed(self):
        if not self.mdb:
            return
        signature = self._cc_magic_import_signature()
        if signature != self._cc_magic_import_last_signature:
            self._refresh_cc_magic_import_statuses()

    def event(self, event):
        result = super().event(event)
        if event.type() == QEvent.WindowActivate:
            QTimer.singleShot(0, self._refresh_cc_magic_import_statuses_if_changed)
        return result

    def _refresh_cc_magic_import_statuses(self, refresh_filters: bool = True):
        """Refresh the live, read-only CC Magic Import status for Library rows.

        This field is never written to the MDB. It reflects the actual configured
        Import folder each time the Library is loaded or a copy/remove action
        completes. Identity is verified by the stored Package Hash whenever a
        same-name Import file exists.
        """
        if not self.mdb or not hasattr(self, "model"):
            return

        import_root_text = str(self.cc_magic_dropin_path or "").strip()
        import_root = Path(import_root_text) if import_root_text else None
        available = bool(import_root and import_root.is_dir())
        import_hashes: Dict[str, Optional[str]] = {}
        if available:
            try:
                for child in import_root.iterdir():
                    if not child.is_file() or child.suffix.lower() != ".package":
                        continue
                    try:
                        import_hashes[child.name.casefold()] = self._managed_file_sha256(child)
                    except OSError:
                        import_hashes[child.name.casefold()] = None
            except OSError:
                available = False
                import_hashes = {}

        for source_row, rec in enumerate(self.model.records):
            filename = Path(str(rec.get("Package", "") or "").strip()).name
            if not available:
                status = "Unavailable"
            elif not filename or not filename.lower().endswith(".package"):
                status = "Not in Import"
            else:
                key = filename.casefold()
                if key not in import_hashes:
                    status = "Not in Import"
                else:
                    actual_hash = import_hashes.get(key)
                    expected_hash = str(rec.get("Package Hash", "") or "").strip().lower()
                    # Older MDB rows may pre-date Package Hash. If a real master
                    # copy can still be resolved outside Import, use its live hash
                    # for this display-only status rather than mislabelling a copy.
                    if actual_hash and not expected_hash:
                        try:
                            master_path, master_hash = self._resolve_managed_package_source(source_row)
                            root_resolved = import_root.resolve() if import_root else None
                            master_resolved = master_path.resolve()
                            inside_import = bool(
                                root_resolved
                                and (master_resolved == root_resolved or root_resolved in master_resolved.parents)
                            )
                            if not inside_import:
                                expected_hash = master_hash
                        except Exception:
                            pass
                    if not actual_hash or not expected_hash:
                        status = "Unverifiable"
                    elif actual_hash == expected_hash:
                        status = "In Import"
                    else:
                        status = "Conflict"
            rec["CC Magic Import"] = status

        self._cc_magic_import_last_signature = self._cc_magic_import_signature()
        self.model.notify_all_changed()
        if refresh_filters and hasattr(self, "filter_boxes"):
            self._populate_filters()
            self._apply_filters()
        if self.current_source_row is not None:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)

    def set_cc_magic_dropin_folder(self) -> bool:
        start = self.cc_magic_dropin_path or str(self.settings.value("last_dir", str(Path.home())) or Path.home())
        folder = QFileDialog.getExistingDirectory(self, "Choose CC Magic Import folder", start)
        if not folder:
            return False
        path = str(Path(folder).resolve())
        self.cc_magic_dropin_path = path
        self.settings.setValue("cc_magic_dropin_folder", path)
        self.settings.sync()
        self._refresh_cc_magic_import_statuses()
        self.statusBar().showMessage(f"CC Magic Import folder set: {path}", 6000)
        return True

    def open_cc_magic_dropin_folder(self):
        path = str(self.cc_magic_dropin_path or "").strip()
        if not path or not Path(path).is_dir():
            answer = QMessageBox.question(
                self, "CC Magic Import",
                "No CC Magic Import folder is configured.\n\n"
                "Set it from Tools → Set CC Magic Import Folder….\n\n"
                "Choose the folder now?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes,
            )
            if answer != QMessageBox.Yes or not self.set_cc_magic_dropin_folder():
                return
            path = self.cc_magic_dropin_path
        self._open_saved_folder(path, "CC Magic Import")

    def copy_selected_to_cc_magic_dropin(self, source_rows: Optional[List[int]] = None):
        if not self.mdb:
            QMessageBox.information(self, "Copy to CC Magic Import", "Open a Master Database first.")
            return
        rows = self._managed_selected_rows(source_rows)
        if not rows:
            QMessageBox.information(self, "Copy to CC Magic Import", "Select one or more package rows first.")
            return
        if self.scan_process is not None and self.scan_process.state() != QProcess.NotRunning:
            QMessageBox.warning(self, "Copy to CC Magic Import", "Wait for the current Scan/Routing operation to finish before copying package files.")
            return
        if not self._confirm_pending_changes():
            return
        if not self.cc_magic_dropin_path or not Path(self.cc_magic_dropin_path).is_dir():
            answer = QMessageBox.question(
                self, "CC Magic Import",
                "No CC Magic Import folder is configured yet.\n\n"
                "Set it from Tools → Set CC Magic Import Folder….\n\n"
                "Choose the folder now?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes,
            )
            if answer != QMessageBox.Yes or not self.set_cc_magic_dropin_folder():
                return
        target_root = Path(self.cc_magic_dropin_path)
        try:
            plans = self._managed_preflight_sources(rows)
        except Exception as exc:
            QMessageBox.warning(self, "Copy preflight failed", str(exc))
            return

        target_keys = set()
        conflicts = []
        to_copy = []
        already_present = []
        for item in plans:
            destination = (target_root / item["filename"]).resolve()
            item["destination"] = destination
            dest_key = os.path.normcase(str(destination)).casefold()
            if dest_key in target_keys:
                conflicts.append(f"Two selected packages would use the same Import-folder name: {item['filename']}")
                continue
            target_keys.add(dest_key)
            if destination.exists():
                try:
                    if self._managed_file_sha256(destination) == item["hash"]:
                        already_present.append(item)
                    else:
                        conflicts.append(f"{destination} already contains a different file.")
                except Exception as exc:
                    conflicts.append(f"Could not verify existing destination '{destination}': {exc}")
            else:
                to_copy.append(item)
        if conflicts:
            QMessageBox.warning(
                self, "Copy blocked",
                "Nothing was copied because the CC Magic Import folder is not conflict-free:\n\n" + "\n".join(f"• {c}" for c in conflicts[:12])
            )
            return
        if not to_copy and already_present:
            QMessageBox.information(
                self, "CC Magic Import",
                f"All {len(already_present):,} selected package(s) are already present in the CC Magic Import folder with matching SHA-256 hashes. Nothing was changed."
            )
            return

        preview = [
            f"Copy {len(to_copy):,} package(s) to CC Magic Import?",
            "",
            f"CC Magic Import folder: {target_root}",
        ]
        if already_present:
            preview.append(f"Exact copies already present and will be left alone: {len(already_present):,}")
        preview.append("")
        for item in to_copy[:12]:
            preview.append(f"• {item['source']}\n  → {item['destination']}")
        if len(to_copy) > 12:
            preview.append(f"…and {len(to_copy) - 12} more package(s).")
        preview += [
            "",
            "Original package files stay in place. CAS Curator verifies SHA-256 after each copy and never overwrites an existing different file. File Location is not changed because CC Magic Import is a deployment copy, not the master storage location.",
        ]
        answer = QMessageBox.question(
            self, "Preview CC Magic copy", "\n".join(preview),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return

        copied = []
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            for item in to_copy:
                destination = item["destination"]
                try:
                    shutil.copy2(item["source"], destination)
                    if self._managed_file_sha256(destination) != item["hash"]:
                        destination.unlink(missing_ok=True)
                        raise ValueError("copied file failed SHA-256 verification")
                    copied.append(item)
                except Exception as exc:
                    try:
                        destination.unlink(missing_ok=True)
                    except Exception:
                        pass
                    cleanup_problems = []
                    for completed in reversed(copied):
                        try:
                            completed["destination"].unlink(missing_ok=True)
                        except Exception as cleanup_exc:
                            cleanup_problems.append(f"{completed['filename']}: {cleanup_exc}")
                    detail = f"Copy failed at '{item['filename']}': {exc}\n\nNew copies from this batch were removed."
                    if cleanup_problems:
                        detail += "\n\nCleanup needs attention:\n" + "\n".join(f"• {p}" for p in cleanup_problems)
                    raise RuntimeError(detail) from exc
        except Exception as exc:
            QMessageBox.critical(self, "CC Magic copy failed", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        self._refresh_cc_magic_import_statuses()
        copy_lines = [f"Copied to CC Magic Import: {len(copied):,}"]
        if already_present:
            copy_lines.append(f"Already present: {len(already_present):,}")
        copy_lines += ["", "Original files stayed in place. File Location was not changed."]
        QMessageBox.information(self, "CC Magic copy complete", "\n".join(copy_lines))
        self.statusBar().showMessage(
            f"CC Magic Import: {len(copied):,} copied, {len(already_present):,} already present", 9000
        )

    def _search_package_web(self, filename: str, creator: str = ""):
        raw_name = str(filename or "").strip()
        if not raw_name:
            self.statusBar().showMessage("No package file name is available to search", 3500)
            return
        base_name = Path(raw_name).name
        clean_name = re.sub(r"\.package$", "", base_name, flags=re.IGNORECASE)
        clean_name = re.sub(r"[_-]+", " ", clean_name)
        clean_name = re.sub(r"\s+", " ", clean_name).strip()
        if not clean_name:
            clean_name = base_name
        clean_creator = re.sub(r"\s+", " ", str(creator or "").strip())
        query_parts = ["Sims 3"]
        if clean_creator:
            query_parts.append(clean_creator)
        query_parts.append(clean_name)
        query = " ".join(query_parts)
        url = QUrl(f"https://www.google.com/search?q={quote_plus(query)}")
        if not QDesktopServices.openUrl(url):
            QMessageBox.warning(
                self,
                "Search Web for Package",
                "CAS Curator could not open the web search in your default browser.",
            )
            return
        shown = f"{clean_creator} — {clean_name}" if clean_creator else clean_name
        self.statusBar().showMessage(f"Searching the web for: {shown}", 5000)

    def _search_creator_web(self, creator: str):
        clean_creator = re.sub(r"\s+", " ", str(creator or "").strip())
        if not clean_creator:
            self.statusBar().showMessage("No Creator is filled in for this package", 3500)
            return
        query = f"Sims 3 {clean_creator}"
        url = QUrl(f"https://www.google.com/search?q={quote_plus(query)}")
        if not QDesktopServices.openUrl(url):
            QMessageBox.warning(
                self,
                "Search Web for Creator",
                "CAS Curator could not open the creator search in your default browser.",
            )
            return
        self.statusBar().showMessage(f"Searching the web for creator: {clean_creator}", 5000)

    def _stage_delete_record(self, rec: Dict[str, str]):
        """Confirm and stage removal of a library record; the actual package file is untouched."""
        if not self.mdb or not rec:
            return
        try:
            excel_row = int(rec.get("_ExcelRow", "0") or 0)
        except (TypeError, ValueError):
            excel_row = 0
        if excel_row <= 1:
            QMessageBox.warning(self, "Delete Record", "This row cannot be deleted safely.")
            return
        if excel_row in self.pending_deletions:
            QMessageBox.information(self, "Delete Record", "This record is already staged for deletion.")
            return
        package_name = str(rec.get("Package", "(unnamed package)") or "(unnamed package)")
        unsaved = self.pending_changes.get(excel_row, {})
        extra = (
            "\n\nUnsaved edits for this record will be discarded if you continue."
            if unsaved else ""
        )
        answer = QMessageBox.question(
            self,
            "Delete Record?",
            f"Remove '{package_name}' from the CAS Curator library database?"
            "\n\nThe record's library fields and Profile assignments will be cleared when you "
            "click Save Changes. A full workbook backup is created before writing."
            "\n\nThe actual .package file will not be deleted, moved or modified. "
            "Scanner history on the Packages and CASP Master sheets will be preserved."
            + extra + "\n\nStage this record for deletion?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            self.statusBar().showMessage("Delete Record cancelled — nothing changed", 5000)
            return
        self.pending_changes.pop(excel_row, None)
        self.pending_deletions.add(excel_row)
        self.model.notify_all_changed()
        if self.current_source_row is not None:
            selected_rec = self.model.record(self.current_source_row)
            if selected_rec and int(selected_rec.get("_ExcelRow", "0") or 0) == excel_row:
                self._set_editor_enabled(False)
        self._update_dirty_state()
        self._update_selected_dirty_label()
        self.statusBar().showMessage(
            f"'{package_name}' is staged for deletion. Use Save Changes to confirm the workbook change, "
            "or Revert All Unsaved Changes to cancel it.",
            10000,
        )

    def _routing_preview_display_path(self, destination: str) -> str:
        """Show a compact route relative to the selected routing root; keep full path in item data/tooltip."""
        raw = str(destination or "").strip()
        if not raw:
            return ""
        try:
            root_text = str(getattr(self, "routing_preview_root", "") or "").strip()
            if root_text:
                root = Path(root_text).resolve()
                dest = Path(raw).resolve()
                try:
                    rel = dest.relative_to(root)
                    return str(rel)
                except ValueError:
                    pass
        except Exception:
            pass
        return raw

    def _copy_routing_destination_from_cell(self, row: int, column: int):
        if column != 3 or not hasattr(self, "routing_preview_table"):
            return
        item = self.routing_preview_table.item(row, 3)
        if not item:
            return
        destination = str(item.data(Qt.UserRole) or item.text()).strip()
        if not destination:
            return
        QApplication.clipboard().setText(destination)
        self.statusBar().showMessage("Routing destination copied to clipboard", 2500)

    def _show_routing_preview_context_menu(self, pos):
        if not hasattr(self, "routing_preview_table"):
            return
        index = self.routing_preview_table.indexAt(pos)
        if not index.isValid():
            return
        row = index.row()
        file_item = self.routing_preview_table.item(row, 0)
        dest_item = self.routing_preview_table.item(row, 3)
        filename = file_item.text().strip() if file_item else ""
        destination = str(dest_item.data(Qt.UserRole) or dest_item.text()).strip() if dest_item else ""
        menu = QMenu(self.routing_preview_table)
        copy_dest = menu.addAction("Copy destination path")
        copy_name = menu.addAction("Copy file name")
        copy_row = menu.addAction("Copy route summary")
        zip_action = None
        source_path = Path(file_item.toolTip()) if file_item and file_item.toolTip() else None
        if source_path and source_path.suffix.lower() == ".zip" and source_path.is_file():
            menu.addSeparator()
            zip_action = menu.addAction("Preview / Extract ZIP…")
        chosen = menu.exec(self.routing_preview_table.viewport().mapToGlobal(pos))
        if zip_action is not None and chosen == zip_action:
            self._show_zip_preview_extract(source_path)
        elif chosen == copy_dest and destination:
            QApplication.clipboard().setText(destination)
            self.statusBar().showMessage("Routing destination copied to clipboard", 2500)
        elif chosen == copy_name and filename:
            QApplication.clipboard().setText(filename)
            self.statusBar().showMessage("File name copied to clipboard", 2500)
        elif chosen == copy_row:
            category_item = self.routing_preview_table.item(row, 1)
            duplicate_item = self.routing_preview_table.item(row, 2)
            status_item = self.routing_preview_table.item(row, 4)
            category = category_item.text().strip() if category_item else ""
            duplicate = duplicate_item.text().strip() if duplicate_item else ""
            status = status_item.text().strip() if status_item else ""
            summary = f"{filename} | {category} | Duplicate: {duplicate} | {destination} | {status}"
            QApplication.clipboard().setText(summary)
            self.statusBar().showMessage("Route summary copied to clipboard", 2500)

    def _choose_zip_preview_extract(self):
        start = self.scan_source_path or str(self.settings.value("last_dir", str(Path.home())))
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose ZIP archive",
            start,
            "ZIP archives (*.zip);;All files (*.*)",
        )
        if path:
            self._show_zip_preview_extract(Path(path))

    def _show_zip_preview_extract(self, zip_path: Path):
        archive = Path(zip_path)
        try:
            plan = _zip_preview_plan(archive)
        except (OSError, zipfile.BadZipFile, ValueError) as exc:
            QMessageBox.warning(self, "ZIP Preview", f"Could not safely preview this ZIP:\n\n{exc}")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("ZIP Preview / Extract")
        dialog.resize(920, 620)
        layout = QVBoxLayout(dialog)

        summary = QLabel(
            f"<b>{html.escape(archive.name)}</b><br>"
            f"{plan['file_count']:,} file(s) · {plan['total_size']:,} bytes<br>"
            + (
                f"Loose/multi-root archive: files will be kept under <b>{html.escape(plan['outer_folder'])}</b>."
                if plan["uses_archive_outer_folder"] else
                f"Existing top-level folder <b>{html.escape(plan['outer_folder'])}</b> will be preserved."
            )
        )
        summary.setWordWrap(True)
        layout.addWidget(summary)

        destination_row = QHBoxLayout()
        destination_row.addWidget(QLabel("Extract into:"))
        destination_edit = QLineEdit(str(archive.parent.resolve()))
        destination_edit.setReadOnly(True)
        destination_row.addWidget(destination_edit, 1)
        choose_destination = QPushButton("Choose folder…")
        destination_row.addWidget(choose_destination)
        layout.addLayout(destination_row)

        table = QTableWidget(0, 3)
        table.setHorizontalHeaderLabels(["ZIP item", "Planned extracted path", "Size"])
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.verticalHeader().setVisible(False)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        for entry in plan["entries"]:
            if entry["is_dir"]:
                continue
            row = table.rowCount()
            table.insertRow(row)
            table.setItem(row, 0, QTableWidgetItem(str(entry["member"])))
            table.setItem(row, 1, QTableWidgetItem(str(entry["relative"])))
            table.setItem(row, 2, QTableWidgetItem(f"{int(entry['size']):,}"))
        layout.addWidget(table, 1)

        safety = QLabel(
            "The original ZIP is never changed or deleted. Extraction refuses unsafe archive paths and will not overwrite an existing file."
        )
        safety.setWordWrap(True)
        layout.addWidget(safety)

        buttons = QDialogButtonBox()
        extract_btn = buttons.addButton("Extract", QDialogButtonBox.AcceptRole)
        close_btn = buttons.addButton(QDialogButtonBox.Close)
        layout.addWidget(buttons)

        def choose_root():
            current = destination_edit.text().strip() or str(archive.parent)
            folder = QFileDialog.getExistingDirectory(dialog, "Choose ZIP extraction destination", current)
            if folder:
                destination_edit.setText(str(Path(folder).resolve()))

        def extract_now():
            raw_root = destination_edit.text().strip()
            if not raw_root:
                QMessageBox.information(dialog, "ZIP Extract", "Choose an extraction destination first.")
                return
            root = Path(raw_root)
            conflicts = _zip_conflicts(plan, root)
            if conflicts:
                shown = "\n".join(conflicts[:8])
                extra = "" if len(conflicts) <= 8 else f"\n…and {len(conflicts) - 8} more"
                QMessageBox.warning(
                    dialog,
                    "ZIP Extract — existing files",
                    "Extraction was stopped because CAS Curator will not overwrite existing files:\n\n"
                    + shown + extra,
                )
                return
            answer = QMessageBox.question(
                dialog,
                "Extract ZIP?",
                f"Extract {plan['file_count']:,} file(s) into:\n\n{root}\n\nThe original ZIP will remain unchanged.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
            try:
                extracted = _extract_zip_plan(plan, root)
            except Exception as exc:
                QMessageBox.critical(dialog, "ZIP Extract failed", f"Nothing was intentionally overwritten.\n\n{exc}")
                return
            self.settings.setValue("last_dir", str(root))
            self.settings.sync()
            self.statusBar().showMessage(f"ZIP extracted: {len(extracted):,} file(s)", 7000)
            QMessageBox.information(
                dialog,
                "ZIP Extract complete",
                f"Extracted {len(extracted):,} file(s).\n\nOriginal ZIP kept:\n{archive}",
            )

        choose_destination.clicked.connect(choose_root)
        extract_btn.clicked.connect(extract_now)
        close_btn.clicked.connect(dialog.reject)
        dialog.exec()

    def _open_last_scan_log(self):
        path = self.last_scan_log_path
        if not path or not path.exists():
            QMessageBox.information(self, "Scanner log", "No saved scanner log is available yet.")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            QMessageBox.warning(self, "Scanner log", f"Could not open:\n{path}")

    def _open_last_routing_preview(self):
        path = self.last_routing_preview_path
        if not path or not path.exists():
            QMessageBox.information(self, "Routing preview", "No saved Routing_Preview.txt is available yet.")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            QMessageBox.warning(self, "Routing preview", f"Could not open:\n{path}")

    def _load_routing_preview_file_if_needed(self):
        path = self.last_routing_preview_path
        if not path or not path.exists() or not hasattr(self, "routing_preview_table"):
            return
        try:
            # Routing_Preview.txt is the authoritative completed dry-run result.
            # Rebuild the table from it even when some live preview rows arrived.
            # Quiet scanner logging can intentionally suppress diagnostic lines,
            # so keeping a partially populated live table caused Big Review to
            # show only the last few routes instead of the complete report.
            self.routing_preview_table.setRowCount(0)
            lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
            for line in lines[1:]:
                parts = line.split("\t", 4)
                if len(parts) != 5:
                    continue
                category, duplicate_flag, source_path, destination, note = parts
                row = self.routing_preview_table.rowCount()
                self.routing_preview_table.insertRow(row)
                source_item = QTableWidgetItem(Path(source_path).name)
                source_item.setToolTip(source_path)
                destination_item = QTableWidgetItem(self._routing_preview_display_path(destination))
                destination_item.setData(Qt.UserRole, destination)
                destination_item.setToolTip(destination)
                self.routing_preview_table.setItem(row, 0, source_item)
                self.routing_preview_table.setItem(row, 1, QTableWidgetItem(category))
                self.routing_preview_table.setItem(row, 2, QTableWidgetItem(duplicate_flag))
                self.routing_preview_table.setItem(row, 3, destination_item)
                self.routing_preview_table.setItem(row, 4, QTableWidgetItem(note))
            if self.routing_preview_table.rowCount():
                self.scan_log.appendPlainText(
                    f"GUI: loaded {self.routing_preview_table.rowCount()} routing preview row(s) from Routing_Preview.txt."
                )
        except Exception as exc:
            self.scan_log.appendPlainText(f"GUI WARNING: could not load Routing_Preview.txt: {exc}")

    def _open_results_folder(self):
        folder = self.last_scan_output_dir
        if not folder or not folder.exists():
            QMessageBox.information(self, "Results folder", "No completed scan results folder is available yet.")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder))):
            QMessageBox.warning(self, "Results folder", f"Could not open:\n{folder}")

    def _open_sorted_folder(self):
        folder = self.last_sorted_root
        if not folder or not folder.exists():
            QMessageBox.information(self, "Sorted folder", "No completed sorted-files folder is available yet.")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder))):
            QMessageBox.warning(self, "Sorted folder", f"Could not open:\n{folder}")

    def _open_last_scan_output_folder(self):
        folder = self.last_sorted_root if (self.scan_run_kind == "routing" and self.last_sorted_root) else self.last_scan_output_dir
        if not folder or not folder.exists():
            QMessageBox.information(self, "Scan folder", "No completed sorted/results folder is available yet.")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder))):
            QMessageBox.warning(self, "Scan folder", f"Could not open:\n{folder}")

    def _scan_process_error(self, _error):
        if not self.scan_process:
            return
        self.scan_log.appendPlainText(f"\nPROCESS ERROR: {self.scan_process.errorString()}")

    def _scan_process_finished(self, exit_code: int, _exit_status):
        if self._scan_output_buffer:
            self._handle_scan_output_line(self._scan_output_buffer.rstrip())
            self._scan_output_buffer = ""

        target = Path(self.scan_mdb_path).resolve() if self.scan_mdb_path else None
        if exit_code == 0:
            self.scan_progress.setRange(0, 0)
            self.scan_progress_label.setText("Finalising…")
            self.scan_current_file_label.setText("Finishing…")
            if self.scan_run_kind == "preview":
                self._load_routing_preview_file_if_needed()
                preview_count = self.routing_preview_table.rowCount() if hasattr(self, "routing_preview_table") else 0
                # The preview table is the authoritative GUI result. Some Qt/process
                # buffering paths can miss the textual ROUTING PREVIEW COUNT line,
                # which previously made the completion dialog incorrectly report 0.
                self.scan_stats["preview_routes"] = preview_count
                if preview_count == 0 and self.scan_total_packages:
                    self.scan_log.appendPlainText(
                        f"GUI WARNING: routing preview returned 0 rows for {self.scan_total_packages} scanned package(s)."
                    )
                self.scan_status_label.setText(f"Routing preview completed — {preview_count} route(s)")
                self.scan_log.appendPlainText("\nGUI: routing preview dry run completed successfully.")
                self.scan_log.appendPlainText("GUI: source files unchanged; destination unchanged; MDB unchanged.")
                self.statusBar().showMessage(f"Routing preview completed — {preview_count} route(s)")
                if hasattr(self, "scan_output_tabs"):
                    self.scan_output_tabs.setCurrentIndex(1)
            elif self.scan_run_kind == "routing":
                self.scan_status_label.setText("Routing completed successfully")
                self.scan_log.appendPlainText("\nGUI: routing completed successfully.")
                self._prepare_routing_recovery_manifest()
                if target and target.is_file():
                    self._post_routing_track_package_locations(target)
                if self.last_sorted_root:
                    self.scan_log.appendPlainText(f"GUI: sorted files destination: {self.last_sorted_root}")
                    if hasattr(self, "scan_open_output_btn"):
                        self.scan_open_output_btn.setEnabled(self.last_sorted_root.exists())
                self.statusBar().showMessage("Routing completed successfully")
                if hasattr(self, "real_routing_check"):
                    self.real_routing_check.setChecked(False)
                if self.mdb and target and self.mdb.path.resolve() == target:
                    self.scan_log.appendPlainText("GUI: Library will refresh automatically before Scan is marked complete…")
                    self._post_scan_refresh_target = target
            else:
                self.scan_status_label.setText("Scan completed successfully")
                self.scan_log.appendPlainText("\nGUI: scanner process completed successfully.")
                self.statusBar().showMessage("Scan completed successfully")
                if self.mdb and target and self.mdb.path.resolve() == target:
                    self.scan_log.appendPlainText("GUI: Library will refresh automatically before Scan is marked complete…")
                    self._post_scan_refresh_target = target

            QTimer.singleShot(0, self._finalise_scan_without_popup)
        else:
            label = "Routing" if self.scan_run_kind == "routing" else ("Routing preview" if self.scan_run_kind == "preview" else "Scanner")
            self.scan_progress.setRange(0, 100)
            self.scan_progress.setValue(0)
            self.scan_progress_label.setText("Stopped")
            self.scan_current_file_label.setText("Stopped")
            self.scan_status_label.setText(f"{label} stopped with error (exit code {exit_code})")
            self.scan_log.appendPlainText(f"\nGUI: scanner process exited with code {exit_code}.")
            self.statusBar().showMessage(f"{label} failed (exit code {exit_code})")

        if self.scan_process:
            self.scan_process.deleteLater()
        self.scan_process = None
        self._update_scan_start_state()

    def choose_file(self):
        if not self._confirm_pending_changes():
            return
        start_dir = self.settings.value("last_dir", str(Path.home()))
        path, _ = QFileDialog.getOpenFileName(self, "Open Master Database", str(start_dir), "Excel Workbook (*.xlsx);;All files (*)")
        if path:
            self.load_file(Path(path))

    def reload(self):
        if self.mdb and self._confirm_pending_changes():
            self.load_file(self.mdb.path)

    def _load_thread_is_running(self) -> bool:
        """Return whether the MDB loader thread is active, tolerating stale Qt wrappers."""
        thread = self._load_thread
        if thread is None:
            return False
        try:
            return thread.isRunning()
        except RuntimeError:
            # PySide can leave the Python wrapper referenced after the underlying
            # C++ QThread has already been deleted. Treat that as finished.
            if self._load_thread is thread:
                self._load_thread = None
                self._load_worker = None
            return False

    def _clear_load_thread_refs(self, thread, worker):
        """Clear loader references only if they still belong to this finished load."""
        if self._load_thread is thread:
            self._load_thread = None
        if self._load_worker is worker:
            self._load_worker = None

    def load_file(self, path: Path):
        if self._load_thread_is_running():
            return
        self.statusBar().showMessage(f"Loading {path.name}…")
        self.setEnabled(False)

        thread = QThread(self)
        worker = LoadWorker(path)
        self._load_thread = thread
        self._load_worker = worker

        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._on_loaded)
        worker.failed.connect(self._on_load_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(lambda t=thread, w=worker: self._clear_load_thread_refs(t, w))
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_load_failed(self, message: str):
        self.setEnabled(True)
        if getattr(self, "_scan_completion_waiting_for_library", False):
            self._scan_completion_waiting_for_library = False
            self.scan_progress.setRange(0, 100)
            self.scan_progress.setValue(99)
            self.scan_progress_label.setText("Refresh failed")
            self.scan_status_label.setText("Scan finished, but Library refresh failed")
        QMessageBox.critical(self, "Could not open MDB", message)
        self.statusBar().showMessage("Could not open MDB")

    def _on_loaded(self, mdb: MDBData):
        self.setEnabled(True)
        self.mdb = mdb
        self.pending_changes.clear()
        self.pending_deletions.clear()
        self.model.set_pending_changes(self.pending_changes)
        self.profile_registry_saved = list(getattr(mdb, "profile_options", []) or [])
        self.profile_registry_current = list(self.profile_registry_saved)
        self.profile_registry_dirty = False
        self.status_registry_dirty = False
        self.file_location_registry_saved = list(getattr(mdb, "file_locations", []) or [])
        self.file_location_registry_current = list(self.file_location_registry_saved)
        self.file_location_registry_dirty = False
        self.merge_profile_registry_saved = list(getattr(mdb, "merge_profile_options", []) or [])
        self.merge_profile_registry_current = list(self.merge_profile_registry_saved)
        self.merge_profile_registry_dirty = False
        self.current_source_row = None
        self.settings.setValue("last_mdb", str(mdb.path.resolve()))
        self.settings.setValue("last_dir", str(mdb.path.parent.resolve()))
        self.settings.sync()
        full_mdb_path = str(mdb.path.resolve())
        if hasattr(self, "mdb_path_label"):
            self.mdb_path_label.setText(full_mdb_path)
            self.mdb_path_label.setToolTip(full_mdb_path)
        self.scan_mdb_path = full_mdb_path
        self.scan_mdb_override = False
        if hasattr(self, "scan_mdb_edit"):
            self.scan_mdb_edit.setText(full_mdb_path)
        self._populate_edit_options()
        self._refresh_cc_magic_profile_choices()
        self._set_editor_enabled(False)
        # Preserve the user's column visibility choice across app restarts and
        # MDB switches. Only apply columns available in the currently loaded MDB.
        available_columns = ["#", "Thumbnail"] + _available_public_columns_for_mdb(mdb)
        saved_columns = self.settings.value("library_visible_columns", None)
        if isinstance(saved_columns, str):
            try:
                parsed_columns = json.loads(saved_columns)
                saved_columns = parsed_columns if isinstance(parsed_columns, list) else None
            except (TypeError, ValueError):
                saved_columns = None
        if isinstance(saved_columns, (list, tuple)):
            selected = [c for c in saved_columns if isinstance(c, str) and c in available_columns]
            if selected:
                preferred = ["#", "Package", "Thumbnail"] + [c for c in _available_public_columns_for_mdb(mdb) if c != "Package"]
                self.visible_columns = sorted(dict.fromkeys(selected), key=lambda x: preferred.index(x) if x in preferred else 9999)
            else:
                self.visible_columns = [c for c in DEFAULT_COLUMNS if c in ("#", "Thumbnail") or c in _available_public_columns_for_mdb(mdb)]
        else:
            self.visible_columns = [c for c in DEFAULT_COLUMNS if c in ("#", "Thumbnail") or c in _available_public_columns_for_mdb(mdb)]
        self.model.set_columns(self.visible_columns)
        self.model.set_records(mdb.records)
        self._refresh_cc_magic_import_statuses(refresh_filters=False)
        self._refresh_cc_magic_profile_choices()
        self._load_pins_for_current_mdb()
        self._refresh_pinned_bar()
        if hasattr(self, "gallery_status_buttons"):
            self._clear_gallery_quick_filters()
        self._build_columns_menu()
        self._update_column_fold_buttons()
        self._populate_filters()
        self._resize_columns()
        self._apply_filters()
        if IS_DEV_BUILD and hasattr(self, "random_sim_category_table"):
            self._refresh_random_sim_data()
            self._random_sim_refresh_history()
        self.packages_label.setText(f"Packages: {len(mdb.records):,}")
        self.thumbs_label.setText(f"Thumbnails: {mdb.thumbnail_count:,}")
        self._update_dirty_state()
        self.statusBar().showMessage(
            f"Loaded: {mdb.path.resolve()} | {len(mdb.records):,} library records | manual editing enabled"
        )
        if getattr(self, "_scan_completion_waiting_for_library", False):
            self._scan_completion_waiting_for_library = False
            self._set_scan_progress_complete()

    def _available_column_fold_fields(self, group_name: str) -> List[str]:
        if not self.mdb:
            return []
        available = set(_available_public_columns_for_mdb(self.mdb))
        return [
            field for field in COLUMN_FOLD_GROUPS.get(group_name, [])
            if field in available and _public_field_visible(field)
        ]

    def _update_column_fold_buttons(self):
        if not hasattr(self, "column_fold_buttons"):
            return
        for group_name, button in self.column_fold_buttons.items():
            fields = self._available_column_fold_fields(group_name)
            button.setEnabled(bool(fields))
            all_visible = bool(fields) and all(field in self.visible_columns for field in fields)
            button.setText(f"{group_name} {'−' if all_visible else '+'}")
            if not fields:
                button.setToolTip(f"No {group_name} columns are currently available. Enable Advanced/Experimental features when applicable.")
            elif all_visible:
                button.setToolTip(f"Hide the available {group_name} column group. Individual choices remain available under Columns.")
            else:
                button.setToolTip(f"Show all currently available {group_name} columns. Individual choices remain available under Columns.")

    def _toggle_column_fold_group(self, group_name: str):
        fields = self._available_column_fold_fields(group_name)
        if not fields:
            self._update_column_fold_buttons()
            return
        all_visible = all(field in self.visible_columns for field in fields)
        if all_visible:
            self.visible_columns = [field for field in self.visible_columns if field not in set(fields)]
        else:
            selected = list(dict.fromkeys(self.visible_columns + fields))
            preferred = ["#", "Package", "Thumbnail"] + [
                c for c in self.mdb.headers if c != "Package" and _public_field_visible(c)
            ]
            self.visible_columns = sorted(
                selected,
                key=lambda x: preferred.index(x) if x in preferred else 9999,
            )
        self.model.set_columns(self.visible_columns)
        self._persist_visible_columns()
        self._resize_columns()
        self._update_column_fold_buttons()

    def _build_columns_menu(self):
        self.columns_menu.clear()
        choose_action = QAction("Choose visible columns…", self)
        choose_action.triggered.connect(self._open_columns_dialog)
        self.columns_menu.addAction(choose_action)
        self.columns_menu.addSeparator()
        reset_action = QAction("Reset to default visible columns", self)
        reset_action.triggered.connect(self._reset_default_columns)
        self.columns_menu.addAction(reset_action)

    def _open_columns_dialog(self):
        if not self.mdb:
            return
        fields = ["#", "Thumbnail"] + _available_public_columns_for_mdb(self.mdb)
        dialog = QDialog(self)
        dialog.setWindowTitle("Choose visible columns")
        dialog.resize(420, 560)
        layout = QVBoxLayout(dialog)
        info = QLabel("Tick the columns you want visible. Scroll for the full list. Advanced and Experimental fields appear here only when enabled in Options.")
        layout.addWidget(info)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(8, 8, 8, 8)
        container_layout.setSpacing(6)

        checks = {}
        for field in fields:
            cb = QCheckBox(_display_field_name(field))
            cb.setChecked(field in self.visible_columns)
            checks[field] = cb
            container_layout.addWidget(cb)
        container_layout.addStretch(1)
        scroll.setWidget(container)
        layout.addWidget(scroll, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        all_btn = QPushButton("Select all")
        default_btn = QPushButton("Default")
        buttons.addButton(all_btn, QDialogButtonBox.ActionRole)
        buttons.addButton(default_btn, QDialogButtonBox.ActionRole)
        all_btn.clicked.connect(lambda: [cb.setChecked(True) for cb in checks.values()])
        default_btn.clicked.connect(lambda: self._apply_default_checks(checks))
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() == QDialog.Accepted:
            selected = [field for field, cb in checks.items() if cb.isChecked()]
            if not selected:
                QMessageBox.information(self, "Columns", "Keep at least one visible column.")
                return
            preferred = ["#", "Package", "Thumbnail"] + ([c for c in _available_public_columns_for_mdb(self.mdb) if c != "Package"] if self.mdb else [])
            self.visible_columns = sorted(selected, key=lambda x: preferred.index(x) if x in preferred else 9999)
            self.model.set_columns(self.visible_columns)
            self._persist_visible_columns()
            self._resize_columns()
            self._update_column_fold_buttons()

    def _persist_visible_columns(self):
        """Persist column visibility independently of the active MDB."""
        self.settings.setValue("library_visible_columns", list(self.visible_columns))
        self.settings.sync()

    def _apply_default_checks(self, checks: Dict[str, QCheckBox]):
        default_set = set(DEFAULT_COLUMNS)
        for field, cb in checks.items():
            cb.setChecked(field in default_set)

    def _reset_default_columns(self):
        if not self.mdb:
            return
        self.visible_columns = [c for c in DEFAULT_COLUMNS if c in ("#", "Thumbnail") or c in _available_public_columns_for_mdb(self.mdb)]
        self.model.set_columns(self.visible_columns)
        self._persist_visible_columns()
        self._resize_columns()
        self._update_column_fold_buttons()

    def _toggle_column(self, field: str, checked: bool):
        if checked and field not in self.visible_columns:
            preferred = ["#", "Package", "Thumbnail"] + ([c for c in _available_public_columns_for_mdb(self.mdb) if c != "Package"] if self.mdb else [])
            self.visible_columns.append(field)
            self.visible_columns.sort(key=lambda x: preferred.index(x) if x in preferred else 9999)
        elif not checked and field in self.visible_columns:
            if len(self.visible_columns) == 1:
                return
            self.visible_columns.remove(field)
        self.model.set_columns(self.visible_columns)
        self._persist_visible_columns()
        self._resize_columns()
        self._update_column_fold_buttons()

    def _resize_columns(self):
        # Keep the curated widths, but never allow a section to be narrower
        # than its rendered header text (plus room for padding/sort indicator).
        # This also behaves better on Windows display scaling >100%.
        header = self.table.horizontalHeader()
        metrics = header.fontMetrics()
        for i, field in enumerate(self.visible_columns):
            base_width = COLUMN_WIDTHS.get(field, 145)
            header_width = metrics.horizontalAdvance(str(field)) + 46
            self.table.setColumnWidth(i, max(base_width, header_width))

    def _available_column_filter_fields(self) -> List[str]:
        """All user-facing data fields currently available in the active mode."""
        if not self.mdb:
            return []
        fields = []
        for field in _available_public_columns_for_mdb(self.mdb):
            if field == "#" or field in PUBLIC_OVERVIEW_ONLY_FIELDS:
                continue
            if field not in fields and _public_field_visible(field):
                fields.append(field)
        # Thumbnail is a live Library column derived from the stored/manual image
        # link, not a physical CC Master header, so add it explicitly.
        if "Thumbnail" not in fields:
            insert_at = fields.index("Package") + 1 if "Package" in fields else 0
            fields.insert(insert_at, "Thumbnail")
        return fields

    def _filter_values_for_field(self, field: str) -> List[str]:
        """Return useful suggestions without restricting free-text filtering."""
        if not self.mdb:
            return []
        if field == "CC Type":
            return _unique_cc_type_values(self.model.records)
        if field == "CC Magic Profiles":
            return self._all_profile_names()
        if field == "CC Magic Import":
            present = {self.model.effective_value(row, field) for row in range(len(self.model.records))}
            order = ["In Import", "Not in Import", "Conflict", "Unverifiable", "Unavailable"]
            return [value for value in order if value in present]
        values = {
            str(self.model.effective_value(row, field) or "").strip()
            for row in range(len(self.model.records))
        }
        values.discard("")
        return sorted(values, key=str.casefold)

    def _sanitize_extra_column_filters(self):
        allowed = set(self._available_column_filter_fields())
        self.extra_column_filters = [
            item for item in self.extra_column_filters
            if str(item.get("field", "") or "") in allowed
        ]

    def _column_filter_chip_text(self, item: Dict[str, str]) -> str:
        field = _display_field_name(str(item.get("field", "") or ""))
        condition = str(item.get("condition", "is") or "is")
        value = str(item.get("value", "") or "")
        if condition in {"is blank", "is not blank"}:
            return f"{field} {condition}"
        shown = value if len(value) <= 42 else value[:39] + "…"
        return f"{field} {condition} {shown}"

    def _refresh_active_filter_chips(self):
        if not hasattr(self, "active_filter_chip_layout"):
            return
        while self.active_filter_chip_layout.count():
            item = self.active_filter_chip_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        for index, item in enumerate(self.extra_column_filters):
            chip = QPushButton(self._column_filter_chip_text(item) + "  ×")
            chip.setObjectName("FilterChip")
            chip.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            chip.setToolTip("Remove this filter")
            chip.clicked.connect(lambda checked=False, i=index: self._remove_column_filter(i))
            self.active_filter_chip_layout.addWidget(chip)
        self.active_filter_chip_layout.addStretch(1)
        self.active_filter_bar.setVisible(bool(self.extra_column_filters))

    def _remove_column_filter(self, index: int):
        if 0 <= index < len(self.extra_column_filters):
            self.extra_column_filters.pop(index)
            self._refresh_active_filter_chips()
            self._apply_filters()

    def _add_column_filter(self):
        if not self.mdb:
            QMessageBox.information(self, "Add Filter", "Open a Master Database first.")
            return
        fields = self._available_column_filter_fields()
        if not fields:
            QMessageBox.information(self, "Add Filter", "No filterable Library fields are currently available.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Add Library Filter")
        self._style_custom_dialog(dialog)
        dialog.setModal(True)
        dialog.setMinimumWidth(440)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        intro = QLabel(
            "Choose any field currently available in the Library. Advanced and Experimental fields appear here only while those options are enabled."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        form = QFormLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        field_box = ArrowComboBox()
        for field in fields:
            field_box.addItem(_display_field_name(field), field)
        condition_box = ArrowComboBox()
        condition_box.addItems(["is", "is not", "contains", "does not contain", "is blank", "is not blank"])
        value_box = ArrowComboBox()
        value_box.setEditable(True)
        value_box.setInsertPolicy(QComboBox.NoInsert)
        value_box.setMinimumWidth(250)
        value_box.setToolTip("Choose an existing value or type your own text.")
        form.addRow("Field:", field_box)
        form.addRow("Condition:", condition_box)
        form.addRow("Value:", value_box)
        layout.addLayout(form)

        def refresh_values(*_args):
            field = str(field_box.currentData() or "")
            values = self._filter_values_for_field(field)
            value_box.blockSignals(True)
            value_box.clear()
            # Large free-text columns can contain thousands of unique strings.
            # Suggestions are capped for a responsive dialog; users can always type.
            value_box.addItems(values[:500])
            if not values:
                value_box.setEditText("")
            value_box.blockSignals(False)

        def refresh_value_enabled(*_args):
            needs_value = condition_box.currentText() not in {"is blank", "is not blank"}
            value_box.setEnabled(needs_value)

        field_box.currentIndexChanged.connect(refresh_values)
        condition_box.currentTextChanged.connect(refresh_value_enabled)
        refresh_values()
        refresh_value_enabled()

        buttons = QDialogButtonBox(QDialogButtonBox.Cancel)
        add_btn = buttons.addButton("Add Filter", QDialogButtonBox.AcceptRole)
        buttons.rejected.connect(dialog.reject)
        add_btn.clicked.connect(dialog.accept)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.Accepted:
            return
        field = str(field_box.currentData() or "")
        condition = condition_box.currentText()
        value = value_box.currentText().strip()
        if condition not in {"is blank", "is not blank"} and not value:
            QMessageBox.information(self, "Add Filter", "Enter or choose a value for this filter.")
            return
        new_filter = {"field": field, "condition": condition, "value": value}
        if new_filter in self.extra_column_filters:
            self.statusBar().showMessage("That filter is already active.", 4000)
            return
        self.extra_column_filters.append(new_filter)
        self._refresh_active_filter_chips()
        self._apply_filters()

    def _populate_filters(self):
        if not self.mdb:
            return
        for field, _ in FILTER_FIELDS:
            box = self.filter_boxes[field]
            current = box.currentText()
            box.blockSignals(True)
            box.clear()
            box.addItem("All")
            values = self._filter_values_for_field(field)
            box.addItems(values)
            if current == "All":
                box.setCurrentText("All")
            elif field == "CC Type":
                wanted = _cc_type_key(current)
                match = next((value for value in values if _cc_type_key(value) == wanted), None)
                box.setCurrentText(match or "All")
            elif current in values:
                box.setCurrentText(current)
            else:
                box.setCurrentText("All")
            box.blockSignals(False)

        self._sanitize_extra_column_filters()
        self._refresh_active_filter_chips()

        if hasattr(self, "gallery_type_filter"):
            current = self.gallery_type_filter.currentText()
            self.gallery_type_filter.blockSignals(True)
            self.gallery_type_filter.clear()
            self.gallery_type_filter.addItem("All")
            type_values = _unique_cc_type_values(self.mdb.records)
            self.gallery_type_filter.addItems(type_values)
            if current == "All":
                self.gallery_type_filter.setCurrentText("All")
            else:
                wanted = _cc_type_key(current)
                match = next((value for value in type_values if _cc_type_key(value) == wanted), None)
                self.gallery_type_filter.setCurrentText(match or "All")
            self.gallery_type_filter.blockSignals(False)

    def _apply_filters(self):
        if IS_PUBLIC_BUILD:
            permanent_fields = {"Status", "CC Type", "Creator", "CC Magic Profiles"}
            filters = {
                field: box.currentText()
                for field, box in self.filter_boxes.items()
                if field in permanent_fields and box.currentText() != "All"
            }
        else:
            filters = {f: b.currentText() for f, b in self.filter_boxes.items() if b.currentText() != "All"}
        self.proxy.set_filters(filters)
        self.proxy.set_column_filters(self.extra_column_filters)
        self.proxy.set_search_text(self.search_edit.text())
        self.shown_label.setText(f"Shown: {self.proxy.rowCount():,}")

    def clear_filters(self):
        self.search_edit.clear()
        for box in self.filter_boxes.values():
            box.setCurrentText("All")
        self.extra_column_filters.clear()
        self._refresh_active_filter_chips()
        self._apply_filters()

    def _show_selected(self, *_):
        selection = self.table.selectionModel()
        proxy_index = selection.currentIndex() if selection else QModelIndex()
        if not proxy_index.isValid():
            indexes = selection.selectedRows() if selection else []
            proxy_index = indexes[0] if indexes else QModelIndex()
        if not proxy_index.isValid():
            self.current_source_row = None
            self._set_editor_enabled(False)
            if hasattr(self, "copy_filename_btn"):
                self.copy_filename_btn.setEnabled(False)
            self._update_pin_button()
            self._refresh_pinned_bar()
            return
        source_index = self.proxy.mapToSource(proxy_index)
        self.current_source_row = source_index.row()
        if hasattr(self, "copy_filename_btn"):
            self.copy_filename_btn.setEnabled(True)
        rec = self.model.record(source_index.row())
        if rec:
            self.show_record(rec)

    def eventFilter(self, watched, event):
        if isinstance(watched, QMessageBox) and event.type() == QEvent.Show:
            self._style_message_box(watched)

        # Respond to the details pane being resized, including window-resize
        # changes where QSplitter.splitterMoved isn't emitted.
        if (hasattr(self, "detail_title_scroll")
                and watched is self.detail_title_scroll.viewport()
                and event.type() == QEvent.Resize):
            new_width = event.size().width()
            if new_width > 0 and new_width != self._detail_title_viewport_width:
                self._detail_title_viewport_width = new_width
                QTimer.singleShot(0, self._fit_detail_title)
        return super().eventFilter(watched, event)

    def _fit_detail_title(self):
        """Show up to four readable lines, scroll for unusually long filenames."""
        if not hasattr(self, "detail_title_scroll"):
            return

        full_text = str(self.detail_title.property("fullPackageName") or self.detail_title.text() or "")
        # Soft-wrap long underscore-heavy CC filenames without changing the
        # underlying filename used by Copy name and the Excel database.
        # Even a creator's filename with no separators should wrap rather than
        # extend outside the available title width.
        wrapped_name = re.sub(
            r"[^_\-.\s]{25,}",
            lambda match: "\u200b".join(match.group()[i:i + 18]
                                    for i in range(0, len(match.group()), 18)),
            full_text,
        )
        display_text = (wrapped_name
                        .replace("_", "_\u200b")
                        .replace("-", "-\u200b")
                        .replace(".", ".\u200b"))
        self.detail_title.setText(display_text)

        # Start with the full title width; if scrolling is needed, reserve room
        # for its scrollbar before measuring again to avoid clipped final lines.
        available = max(85, self.detail_title_scroll.viewport().width() - 5)
        limit = 92
        chosen_font = self.detail_title.font()
        measured = 0
        for size in (11.0, 10.5, 10.0, 9.5):
            test_font = self.detail_title.font()
            test_font.setPointSizeF(size)
            test_font.setBold(True)
            metrics = QFontMetrics(test_font)
            rect = metrics.boundingRect(QRect(0, 0, available, 5000), Qt.TextWordWrap, display_text)
            chosen_font = test_font
            measured = rect.height() + 8
            if measured <= limit - 5:
                break
        if measured > limit:
            # Re-measure with a visible scrollbar (rather than letting it
            # consume width and unexpectedly produce an extra clipped line).
            width = max(85, self.detail_title_scroll.width() - 23)
            metrics = QFontMetrics(chosen_font)
            measured = metrics.boundingRect(QRect(0, 0, width, 5000), Qt.TextWordWrap, display_text).height() + 8

        self.detail_title.setFont(chosen_font)
        content_height = max(40, measured)
        self.detail_title.setFixedHeight(content_height)
        self.detail_title_scroll.widget().setMinimumHeight(content_height)
        self.detail_title_scroll.setFixedHeight(min(limit, content_height + 2))
        # New selection starts at the filename's beginning, not wherever the
        # previous selected package's title was scrolled.
        if self.detail_title_scroll.verticalScrollBar().value():
            self.detail_title_scroll.verticalScrollBar().setValue(0)

    def _random_sim_usage_for_record(self, rec: Dict[str, str]):
        if not IS_DEV_BUILD:
            return [], []
        selected = selected_usages_for_record(rec, self.random_sim_usage_registry)
        dependencies = dependency_usages_for_record(rec, self.random_sim_usage_registry)
        return selected, dependencies

    def _random_sim_usage_html(self, rec: Dict[str, str]) -> str:
        if not IS_DEV_BUILD:
            return ""
        selected, dependencies = self._random_sim_usage_for_record(rec)
        rows = []
        if selected:
            links = []
            for usage in selected:
                sim_id = html.escape(str(usage.get("sim_id", "")), quote=True)
                name = html.escape(str(usage.get("name", "Random Sim") or "Random Sim"))
                links.append(f'<a href="randomsim:{sim_id}">{name}</a>')
            rows.append(
                "<div style='margin-bottom:10px'><b>Used in Random Sim:</b>&nbsp; "
                + ", ".join(links)
                + "</div>"
            )
        if dependencies:
            links = []
            for usage in dependencies:
                sim_id = html.escape(str(usage.get("sim_id", "")), quote=True)
                name = html.escape(str(usage.get("name", "Random Sim") or "Random Sim"))
                links.append(f'<a href="randomsim:{sim_id}">{name}</a>')
            rows.append(
                "<div style='margin-bottom:10px'><b>Required by Random Sim:</b>&nbsp; "
                + ", ".join(links)
                + "</div>"
            )
        return "".join(rows)

    def _update_random_sim_usage_badge(self, rec: Dict[str, str]):
        badge = getattr(self, "random_sim_usage_badge", None)
        if badge is None:
            return
        selected, dependencies = self._random_sim_usage_for_record(rec)
        if not selected and not dependencies:
            badge.setVisible(False)
            badge.setToolTip("Random Sim Studio collections that use this original CC package")
            return
        if selected and dependencies:
            text = f"Sim use: {len(selected)} (+{len(dependencies)} dep)"
        elif selected:
            text = f"Sim use: {len(selected)}"
        else:
            text = f"Sim dep: {len(dependencies)}"
        badge.setText(text)
        badge.setVisible(True)
        lines = []
        for usage in selected:
            lines.append(f"Used in: {usage.get('name', 'Random Sim')}")
        for usage in dependencies:
            lines.append(f"Required by: {usage.get('name', 'Random Sim')}")
        badge.setToolTip("\n".join(lines))

    def _open_random_sim_usage_link(self, url: QUrl):
        text = url.toString()
        if not text.startswith("randomsim:"):
            QDesktopServices.openUrl(url)
            return
        sim_id = text.split(":", 1)[1].strip().lstrip("/")
        sims = self.random_sim_usage_registry.get("sims", {})
        sim = sims.get(sim_id) if isinstance(sims, dict) else None
        if not isinstance(sim, dict):
            QMessageBox.information(self, "Random Sim link", "This Random Sim history entry is no longer available.")
            return
        folder = Path(str(sim.get("folder", "") or ""))
        if folder.is_dir():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))
            return
        QMessageBox.information(
            self,
            "Random Sim link",
            f"The Random Sim record still exists, but its exported folder was not found:\n\n{folder}",
        )

    def show_record(self, rec: Dict[str, str]):
        effective = dict(rec)
        row_num = int(rec.get("_ExcelRow", "0") or 0)
        if row_num and row_num in self.pending_changes:
            effective.update(self.pending_changes[row_num])

        package_name = rec.get("Package", "Package")
        self.detail_title.setProperty("fullPackageName", package_name)
        self.detail_title.setText(package_name)
        self.detail_title.setToolTip(package_name)
        self._fit_detail_title()
        QTimer.singleShot(0, self._fit_detail_title)
        self.status_badge.setText("Status: " + (effective.get("Status", "") or "—"))
        self.cc_type_badge.setText("CC Type: " + (rec.get("CC Type", "") or "—"))
        self.polys_badge.setText("Highest Polys: " + (rec.get("Highest Polys", "") or "—"))
        # Random Sim usage is an independent DEV relationship label. It never
        # replaces Status, so "Used in Game" and "Random Sims: N" can coexist.
        self._update_random_sim_usage_badge(rec)
        if row_num and "Image" in self.pending_changes.get(row_num, {}):
            effective["_ThumbnailTarget"] = effective.get("Image", "")
        self._show_thumbnail(effective)
        # Painting always fits the *current* visible rectangle, even after
        # title wrapping, window resizing or splitter changes.
        self.thumbnail_label.update()
        self._load_editor_for_record(rec)
        self._update_pin_button()
        self._refresh_pinned_bar()

        summary_html = []
        for field in _BASE_SUMMARY_FIELDS:
            if not (_public_field_visible(field) or (IS_PUBLIC_BUILD and field in PUBLIC_OVERVIEW_ONLY_FIELDS)):
                continue
            # CC Type and Highest Polys are already shown in the top summary
            # row. Keep Health/Risk and the organisation metadata below so the
            # Overview does not duplicate the same values twice.
            if IS_PUBLIC_BUILD and field in {"CC Type", "Highest Polys"}:
                continue
            value = effective.get(field, "")
            if not value:
                continue
            safe = (
                value.replace("&", "&amp;")
                     .replace("<", "&lt;")
                     .replace(">", "&gt;")
                     .replace("\n", "<br>")
            )
            if field == "Subtype" and (IS_DEV_BUILD or self.experimental_features):
                basis = _subtype_basis_label(effective)
                if basis:
                    safe += (
                        "<br><span style='font-size:11px; color:#666666'>"
                        f"Experimental · {html.escape(basis)}</span>"
                    )
            if field == "Skin Type" and (IS_DEV_BUILD or self.experimental_features):
                basis = _skin_type_basis_label(effective)
                if basis:
                    safe += (
                        "<br><span style='font-size:11px; color:#666666'>"
                        f"Experimental · {html.escape(basis)}</span>"
                    )
            if field == "Outfit Category" and (IS_DEV_BUILD or self.experimental_features):
                basis = _outfit_category_basis_label(effective)
                if basis:
                    safe += (
                        "<br><span style='font-size:11px; color:#666666'>"
                        f"Experimental · {html.escape(basis)}</span>"
                    )
            display_name = _display_field_name(field)
            if field == "CC Magic Profiles":
                display_name = "Profile" if len(_split_cc_magic_profiles(value)) == 1 else "Profiles"
            summary_html.append(f"<div style='margin-bottom:10px'><b>{display_name}:</b>&nbsp; {safe}</div>")
        usage_html = self._random_sim_usage_html(rec)
        if usage_html:
            summary_html.insert(0, usage_html)
        self.info_browser.setHtml("".join(summary_html) or "<span style='color:#171717'>No additional details.</span>")

        if hasattr(self, "casp_observations_browser"):
            casp_rows = []
            # Keep this tab limited to CASP-derived information. Experimental
            # Subtype is intentionally excluded because current Build/Buy subtype
            # values can be filename hints; Outfit Category is allowed here because
            # it is aggregated directly from CASP ClothingCategory observations.
            casp_fields = [("CC Type", "CC Type"), ("Age", "Age"), ("Gender", "Gender")]
            if IS_DEV_BUILD or self.experimental_features:
                casp_fields.append(("Outfit Category", "Outfit Category"))
            if IS_DEV_BUILD or self.advanced_features:
                casp_fields.insert(0, ("CASP Instance IDs", "CASP Instance IDs"))
            for label, field in casp_fields:
                value = str(effective.get(field, "") or "").strip()
                if not value:
                    continue
                safe_value = html.escape(value).replace("\n", "<br>")
                casp_rows.append(f"<div style='margin-bottom:10px'><b>{label}:</b>&nbsp; {safe_value}</div>")
            has_casp_identity = bool(str(effective.get("CASP Instance IDs", "") or "").strip())
            if has_casp_identity:
                if IS_PUBLIC_BUILD and not self.advanced_features:
                    casp_rows.append("<div style='margin-top:8px'><i>Enable Advanced features in Options to view the CASP Instance IDs.</i></div>")
                self.casp_observations_browser.setHtml("".join(casp_rows) or "Direct CASP identity is linked, but no additional display fields are available.")
            else:
                self.casp_observations_browser.setHtml(
                    "<p><b>No direct CASP observations are linked to this package yet.</b></p>"
                    "<p>Run the scanner on the original package and reload the Master Database.</p>"
                )

        if self.mdb:
            public_headers = [field for field in self.mdb.headers if _public_field_visible(field) or (IS_PUBLIC_BUILD and field in PUBLIC_OVERVIEW_ONLY_FIELDS)]
            self.fields_table.setRowCount(len(public_headers))
            for row, field in enumerate(public_headers):
                fitem = QTableWidgetItem(_display_field_name(field))
                vitem = QTableWidgetItem(effective.get(field, ""))
                self.fields_table.setItem(row, 0, fitem)
                self.fields_table.setItem(row, 1, vitem)

    def _table_editor_options(self, field: str) -> List[str]:
        if not self.mdb:
            return []
        if field == "Status":
            return list(PUBLIC_STATUS_OPTIONS)
        if field == "File Location":
            return [name for name, _path in self.file_location_registry_current]
        if field == "CC Magic Profiles":
            return self._all_profile_names()
        return []

    def _table_fill_requested(self, source_proxy_index: QModelIndex, target_proxy_index: QModelIndex):
        if not self.mdb or not source_proxy_index.isValid() or not target_proxy_index.isValid():
            return
        if source_proxy_index.column() != target_proxy_index.column():
            return
        if target_proxy_index.row() <= source_proxy_index.row():
            return

        source_index = self.proxy.mapToSource(source_proxy_index)
        if not source_index.isValid() or not (0 <= source_index.column() < len(self.model.columns)):
            return
        field = self.model.columns[source_index.column()]
        if field not in EDITABLE_FIELDS:
            return

        value = self.model.effective_value(source_index.row(), field)
        if not str(value).strip():
            QMessageBox.information(
                self,
                "Fill down",
                "The source cell is blank. Blank fill-down is blocked so populated values cannot be cleared accidentally.",
            )
            return

        # Resolve the complete displayed drag range before staging edits. This is
        # essential when the edited field also participates in the active filter:
        # rows may disappear from the proxy as soon as their value changes.
        mapped_source_rows: List[int] = []
        overwritten_nonblank = 0
        already_same = 0
        skipped_deletions = 0
        for proxy_row in range(source_proxy_index.row() + 1, target_proxy_index.row() + 1):
            proxy_idx = self.proxy.index(proxy_row, source_proxy_index.column())
            src_idx = self.proxy.mapToSource(proxy_idx)
            if not src_idx.isValid():
                continue
            src_row = src_idx.row()
            rec = self.model.record(src_row)
            excel_row = self.model._excel_row(rec)
            if excel_row in self.pending_deletions:
                skipped_deletions += 1
                continue
            current = self.model.effective_value(src_row, field)
            mapped_source_rows.append(src_row)
            if current == value:
                already_same += 1
            elif str(current).strip():
                overwritten_nonblank += 1

        if not mapped_source_rows:
            self.statusBar().showMessage("No eligible displayed rows were crossed by the fill drag.", 5000)
            return

        eligible = len(mapped_source_rows)
        changing = eligible - already_same
        label = _display_field_name(field)
        detail_lines = [
            f"Copy the {label} value into {eligible:,} displayed row(s)?",
            f"{changing:,} row(s) will change; {already_same:,} already have the same value.",
        ]
        if overwritten_nonblank:
            detail_lines.append(f"{overwritten_nonblank:,} non-empty value(s) will be replaced.")
        if skipped_deletions:
            detail_lines.append(f"{skipped_deletions:,} row(s) staged for deletion will be skipped.")
        detail_lines.append("These edits remain staged until you use Save Changes.")
        if getattr(self, "confirm_fill_action", None) is None or self.confirm_fill_action.isChecked():
            answer = QMessageBox.question(
                self,
                "Fill down",
                "\n\n".join(detail_lines),
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return

        changed_rows = self.model.stage_fill_value(mapped_source_rows, field, value)
        if not changed_rows:
            self.statusBar().showMessage(f"No {label} values needed changing.", 5000)
            return

        for source_row in changed_rows:
            self.model.notify_row_changed(source_row)
        if self.current_source_row in changed_rows:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)
        self._update_dirty_state()
        self._update_selected_dirty_label()
        self._apply_filters()
        self.statusBar().showMessage(
            f"Filled {label} into {len(changed_rows):,} row(s) — unsaved until Save Changes",
            8000,
        )

    def _table_manual_edit_changed(self, source_row: int, field: str, value: str):
        # The model has already updated the shared pending_changes dictionary.
        # Refresh only the UI; the safe XLSX write-back still happens exclusively in save_changes().
        self.model.notify_row_changed(source_row)
        if self.current_source_row == source_row:
            rec = self.model.record(source_row)
            if rec:
                self.show_record(rec)
        self._update_dirty_state()
        self._update_selected_dirty_label()
        self._apply_filters()
        self.statusBar().showMessage(
            f"Unsaved {field} change — press Ctrl+S or use Save Changes to write it to the MDB",
            7000,
        )

    def _populate_edit_options(self):
        if not self.mdb:
            return
        self._loading_editor = True
        try:
            self.status_edit.clear()
            self.status_edit.addItem("")
            self.status_edit.addItems(PUBLIC_STATUS_OPTIONS)
            self.file_location_edit.clear()
            self.file_location_edit.addItem("")
            self.file_location_edit.addItems([name for name, _path in self.file_location_registry_current])
            self.file_location_manage_btn.setEnabled(True)
        finally:
            self._loading_editor = False

    def _file_location_path_for_name(self, name: str) -> str:
        key = str(name or "").strip().casefold()
        for saved_name, folder_path in self.file_location_registry_current:
            if saved_name.casefold() == key:
                return folder_path
        return ""

    def _open_saved_folder(self, folder_path: str, friendly_name: str = "File Location") -> bool:
        """Open exactly the saved folder; never guess a replacement for a missing path."""
        raw_path = str(folder_path or "").strip()
        if not raw_path:
            QMessageBox.warning(self, friendly_name, "No folder path is saved for this File Location.")
            return False
        folder = Path(raw_path)
        if not folder.is_dir():
            QMessageBox.warning(
                self,
                f"{friendly_name} unavailable",
                f"The saved folder is not available:\n\n{raw_path}\n\n"
                "No other folder will be opened automatically. If you moved the folder, "
                "use Manage Locations… to update this location explicitly.",
            )
            return False
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder.resolve()))):
            QMessageBox.warning(self, friendly_name, f"Windows could not open this folder:\n\n{raw_path}")
            return False
        return True

    def _open_file_location_for_source_row(self, source_row: int) -> bool:
        if not self.mdb or source_row is None or source_row < 0:
            return False
        rec = self.model.record(source_row)
        if not rec:
            return False
        name = self.model.effective_value(source_row, "File Location").strip()
        if not name:
            QMessageBox.information(self, "Open Folder", "This package has no File Location assigned.")
            return False
        path = self._file_location_path_for_name(name)
        if not path:
            QMessageBox.warning(
                self, "File Location unavailable",
                f"'{name}' is assigned to this package, but no saved folder path exists for that name.\n\n"
                "The old Library Location categories are not physical paths and will not be used as substitutes.",
            )
            return False
        return self._open_saved_folder(path, name)

    def open_selected_file_location(self):
        if not self.mdb or self.current_source_row is None:
            QMessageBox.information(self, "Open Folder", "Select a package first.")
            return
        self._open_file_location_for_source_row(self.current_source_row)

    def _stage_file_location_rename(self, old_name: str, new_name: str) -> int:
        """Stage assignment-name updates so renaming a location cannot orphan package links."""
        old_key = old_name.casefold()
        changed = 0
        for source_row, rec in enumerate(self.model.records):
            try:
                excel_row = int(rec.get("_ExcelRow", "0") or 0)
            except (TypeError, ValueError):
                continue
            if not excel_row or excel_row in self.pending_deletions:
                continue
            current = self.model.effective_value(source_row, "File Location").strip()
            if current.casefold() != old_key:
                continue
            row_pending = dict(self.pending_changes.get(excel_row, {}))
            original = rec.get("File Location", "")
            if new_name == original:
                row_pending.pop("File Location", None)
            else:
                row_pending["File Location"] = new_name
            if row_pending:
                self.pending_changes[excel_row] = row_pending
            else:
                self.pending_changes.pop(excel_row, None)
            self.model.notify_row_changed(source_row)
            changed += 1
        if self.current_source_row is not None:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)
        return changed

    def manage_file_locations(self):
        """Manage named physical folders; all changes remain staged until Save Changes."""
        if not self.mdb:
            QMessageBox.information(self, "Manage File Locations", "Open a Master Database first.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Manage File Locations")
        dialog.resize(900, 460)
        layout = QVBoxLayout(dialog)
        intro = QLabel(
            "Each name points to the folder containing an original, unmerged package. Assigning a location never moves or copies a package. "
            "Missing folders are shown as unavailable; paths are never substituted automatically. "
            "Changes remain staged until Save Changes."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        table = QTableWidget(0, 4, dialog)
        table.setHorizontalHeaderLabels(["Location name", "Saved folder path", "Availability", "Open"])
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table, 1)

        buttons_row = QHBoxLayout()
        add_btn = QPushButton("Add Folder…")
        rename_btn = QPushButton("Rename…")
        change_path_btn = QPushButton("Change Folder Path…")
        remove_btn = QPushButton("Remove Location…")
        close_btn = QPushButton("Close")
        for btn in (add_btn, rename_btn, change_path_btn, remove_btn):
            buttons_row.addWidget(btn)
        self_location_selection_label = QLabel("Selected: none")
        self_location_selection_label.setObjectName("LocationSelectionLabel")
        buttons_row.addSpacing(8)
        buttons_row.addWidget(self_location_selection_label)
        buttons_row.addStretch(1)
        buttons_row.addWidget(close_btn)
        layout.addLayout(buttons_row)

        def refresh_table(select_name: Optional[str] = None):
            locations = sorted(self.file_location_registry_current, key=lambda item: item[0].casefold())
            table.setRowCount(len(locations))
            selected_row = -1
            for row, (name, folder_path) in enumerate(locations):
                name_item = QTableWidgetItem(name)
                path_item = QTableWidgetItem(folder_path)
                available = Path(folder_path).is_dir()
                availability_item = QTableWidgetItem("Available" if available else "Unavailable")
                for item in (name_item, path_item, availability_item):
                    item.setToolTip(item.text())
                name_item.setData(Qt.UserRole, name)
                path_item.setData(Qt.UserRole, folder_path)
                table.setItem(row, 0, name_item)
                table.setItem(row, 1, path_item)
                table.setItem(row, 2, availability_item)
                row_open_btn = QPushButton("Open")
                row_open_btn.setEnabled(available)
                row_open_btn.setToolTip("Open this exact saved folder in Windows File Explorer." if available else "This saved folder is currently unavailable.")
                row_open_btn.clicked.connect(lambda checked=False, p=folder_path, n=name: self._open_saved_folder(p, n))
                table.setCellWidget(row, 3, row_open_btn)
                if select_name and name.casefold() == select_name.casefold():
                    selected_row = row
            table.resizeColumnsToContents()
            if selected_row >= 0:
                table.selectRow(selected_row)
            update_buttons()

        def selected_location() -> Optional[tuple[str, str]]:
            row = table.currentRow()
            if row < 0 or table.item(row, 0) is None:
                return None
            return table.item(row, 0).text(), table.item(row, 1).text()

        def update_buttons(*_):
            selected = selected_location() if table.rowCount() else None
            has_selection = selected is not None
            rename_btn.setEnabled(has_selection)
            change_path_btn.setEnabled(has_selection)
            remove_btn.setEnabled(has_selection)
            self_location_selection_label.setText(f"Selected: {selected[0]}" if selected else "Selected: none")

        def set_registry_dirty():
            self.file_location_registry_current.sort(key=lambda item: item[0].casefold())
            self.file_location_registry_dirty = self.file_location_registry_current != self.file_location_registry_saved
            self._populate_edit_options()
            if self.current_source_row is not None:
                rec = self.model.record(self.current_source_row)
                if rec:
                    self.show_record(rec)
            self._populate_filters()
            self._apply_filters()
            self._update_dirty_state()

        def add_folder():
            start = str(self.settings.value("last_file_location_dir", str(self.mdb.path.parent)) or self.mdb.path.parent)
            folder = QFileDialog.getExistingDirectory(dialog, "Choose folder for File Location", start)
            if not folder:
                return
            folder_path = str(Path(folder).resolve())
            folder_key = os.path.normcase(os.path.abspath(folder_path)).casefold()
            for existing_name, existing_path in self.file_location_registry_current:
                if os.path.normcase(os.path.abspath(existing_path)).casefold() == folder_key:
                    QMessageBox.information(dialog, "File Location already exists", f"This folder is already saved as '{existing_name}'.")
                    refresh_table(existing_name)
                    return
            name, ok = QInputDialog.getText(dialog, "Name File Location", "Friendly location name:")
            if not ok:
                return
            name = name.strip()
            if not name:
                QMessageBox.information(dialog, "File Location", "Enter a non-empty location name.")
                return
            if any(old_name.casefold() == name.casefold() for old_name, _ in self.file_location_registry_current):
                QMessageBox.warning(dialog, "File Location", f"A location named '{name}' already exists.")
                return
            self.file_location_registry_current.append((name, folder_path))
            self.settings.setValue("last_file_location_dir", str(Path(folder_path).parent))
            set_registry_dirty()
            refresh_table(name)

        def rename_location():
            selected = selected_location()
            if not selected:
                return
            old_name, old_path = selected
            new_name, ok = QInputDialog.getText(dialog, "Rename File Location", "Friendly location name:", text=old_name)
            if not ok:
                return
            new_name = new_name.strip()
            if not new_name:
                QMessageBox.information(dialog, "File Location", "Enter a non-empty location name.")
                return
            if new_name.casefold() != old_name.casefold() and any(
                name.casefold() == new_name.casefold() for name, _ in self.file_location_registry_current
            ):
                QMessageBox.warning(dialog, "File Location", f"A location named '{new_name}' already exists.")
                return
            self.file_location_registry_current = [
                (new_name, path) if name.casefold() == old_name.casefold() else (name, path)
                for name, path in self.file_location_registry_current
            ]
            changed = 0
            if new_name != old_name:
                changed = self._stage_file_location_rename(old_name, new_name)
            set_registry_dirty()
            refresh_table(new_name)
            if changed:
                self.statusBar().showMessage(f"Staged File Location rename across {changed:,} package assignment(s). Save Changes to keep it.", 9000)

        def change_folder_path():
            selected = selected_location()
            if not selected:
                return
            name, old_path = selected
            start = old_path if Path(old_path).is_dir() else str(self.mdb.path.parent)
            folder = QFileDialog.getExistingDirectory(dialog, f"Choose folder for '{name}'", start)
            if not folder:
                return
            new_path = str(Path(folder).resolve())
            new_key = os.path.normcase(os.path.abspath(new_path)).casefold()
            duplicate = next((
                saved_name for saved_name, saved_path in self.file_location_registry_current
                if saved_name.casefold() != name.casefold()
                and os.path.normcase(os.path.abspath(saved_path)).casefold() == new_key
            ), None)
            if duplicate:
                QMessageBox.warning(dialog, "File Location", f"That folder is already saved as '{duplicate}'.")
                return
            self.file_location_registry_current = [
                (saved_name, new_path if saved_name.casefold() == name.casefold() else saved_path)
                for saved_name, saved_path in self.file_location_registry_current
            ]
            self.settings.setValue("last_file_location_dir", str(Path(new_path).parent))
            set_registry_dirty()
            refresh_table(name)

        def remove_location():
            selected = selected_location()
            if not selected:
                return
            name, _folder_path = selected
            usage = 0
            for source_row, rec in enumerate(self.model.records):
                try:
                    excel_row = int(rec.get("_ExcelRow", "0") or 0)
                except (TypeError, ValueError):
                    continue
                if not excel_row or excel_row in self.pending_deletions:
                    continue
                if self.model.effective_value(source_row, "File Location").strip().casefold() == name.casefold():
                    usage += 1
            if usage:
                QMessageBox.warning(
                    dialog, "File Location is in use",
                    f"'{name}' is assigned to {usage:,} package(s). Reassign those packages first, then remove the location."
                )
                return
            answer = QMessageBox.question(
                dialog, "Remove File Location?",
                f"Remove the saved location '{name}' from the Master Database?\n\n"
                "This does not delete or move the folder or any package files.",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
            self.file_location_registry_current = [
                item for item in self.file_location_registry_current if item[0].casefold() != name.casefold()
            ]
            set_registry_dirty()
            refresh_table()

        def open_location():
            selected = selected_location()
            if selected:
                self._open_saved_folder(selected[1], selected[0])

        add_btn.clicked.connect(add_folder)
        rename_btn.clicked.connect(rename_location)
        change_path_btn.clicked.connect(change_folder_path)
        remove_btn.clicked.connect(remove_location)
        close_btn.clicked.connect(dialog.accept)
        table.currentCellChanged.connect(lambda *_: update_buttons())
        refresh_table()
        dialog.exec()

    def _set_editor_enabled(self, enabled: bool):
        for widget in (
            getattr(self, "status_edit", None),
            getattr(self, "creator_edit", None),
            getattr(self, "file_location_edit", None),
            getattr(self, "cc_magic_profiles_edit", None),
            getattr(self, "merge_profiles_edit", None),
            getattr(self, "merge_profiles_set_btn", None),
            getattr(self, "merge_profiles_manage_btn", None),
            getattr(self, "notes_edit", None),
            getattr(self, "preview_image_add_btn", None),
            getattr(self, "preview_image_clear_btn", None),
            getattr(self, "cc_magic_profile_quick", None),
            getattr(self, "cc_profile_add_btn", None),
            getattr(self, "cc_profile_remove_btn", None),
            getattr(self, "cc_profile_manage_btn", None),
        ):
            if widget is not None:
                widget.setEnabled(enabled)
        if hasattr(self, "file_location_open_btn"):
            assigned_location = ""
            if enabled and self.current_source_row is not None:
                assigned_location = self.model.effective_value(self.current_source_row, "File Location").strip()
            self.file_location_open_btn.setEnabled(bool(enabled and assigned_location))
        if hasattr(self, "save_btn"):
            self.save_btn.setEnabled(
                bool(self.pending_changes) or bool(self.pending_deletions)
                or self.profile_registry_dirty or self.status_registry_dirty or self.file_location_registry_dirty
                or self.merge_profile_registry_dirty
            )
        if hasattr(self, "revert_selected_btn"):
            row_num = self._current_excel_row()
            self.revert_selected_btn.setEnabled(
                (enabled and row_num in self.pending_changes) or row_num in self.pending_deletions
            )

    def _all_profile_names(self) -> List[str]:
        """Return the profile registry plus any assigned legacy profile names."""
        result: List[str] = []
        seen = set()
        for raw in list(self.profile_registry_current) + _unique_cc_magic_profiles(self.model.records):
            name = str(raw or "").strip()
            if not name:
                continue
            key = name.casefold()
            if key in seen:
                continue
            seen.add(key)
            result.append(name)
        return sorted(result, key=str.casefold)

    def _refresh_cc_magic_profile_choices(self):
        combo = getattr(self, "cc_magic_profile_quick", None)
        if combo is None:
            return
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("Choose Profile…", "")
        if self.mdb:
            for profile in self._all_profile_names():
                combo.addItem(profile, profile)
        combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def _valid_profile_name(self, name: str) -> bool:
        return bool(name and ";" not in name and "\n" not in name and "\r" not in name)

    def _ensure_profile_in_registry(self, profile: str) -> bool:
        """Add a profile to the independent registry if needed; return True if changed."""
        profile = str(profile or "").strip()
        if not self._valid_profile_name(profile):
            return False
        keys = {p.casefold() for p in self.profile_registry_current}
        if profile.casefold() in keys:
            return False
        self.profile_registry_current.append(profile)
        self.profile_registry_current.sort(key=str.casefold)
        self.profile_registry_dirty = (
            [p.casefold() for p in self.profile_registry_current]
            != [p.casefold() for p in self.profile_registry_saved]
            or self.profile_registry_current != self.profile_registry_saved
        )
        self._refresh_cc_magic_profile_choices()
        self._populate_filters()
        self._update_dirty_state()
        return True

    def _create_profile(self):
        if not self.mdb:
            QMessageBox.information(self, "New Profile", "Open a Master Database first.")
            return
        name, ok = QInputDialog.getText(
            self,
            "New Profile",
            "Profile name:\n\nProfiles can represent a playthrough, theme, save, testing setup, or other organisational group.",
        )
        if not ok:
            return
        name = str(name or "").strip()
        if not name:
            QMessageBox.information(self, "New Profile", "Profile name cannot be blank.")
            return
        if not self._valid_profile_name(name):
            QMessageBox.information(self, "New Profile", "Use one profile name only (no semicolons or line breaks).")
            return
        existing = next((p for p in self._all_profile_names() if p.casefold() == name.casefold()), None)
        if existing:
            QMessageBox.information(self, "New Profile", f"'{existing}' already exists.")
            if hasattr(self, "cc_magic_profile_quick"):
                self.cc_magic_profile_quick.setCurrentText(existing)
            return
        self._ensure_profile_in_registry(name)
        if hasattr(self, "cc_magic_profile_quick"):
            self.cc_magic_profile_quick.setCurrentText(name)
        self.statusBar().showMessage(
            f"Created profile '{name}' (unsaved) — assign packages now or use Save Changes to keep the empty profile",
            9000,
        )

    def _cc_profile_quick_selected(self, index: int):
        """Add the selected existing Profile to the current package only."""
        if index <= 0 or self.current_source_row is None:
            return
        combo = getattr(self, "cc_magic_profile_quick", None)
        if combo is None:
            return
        profile = str(combo.itemData(index) or combo.itemText(index) or "").strip()
        if not profile:
            combo.setCurrentIndex(0)
            return
        changed = self._stage_profile_on_rows(profile, [self.current_source_row], True)
        combo.setCurrentIndex(0)
        if changed:
            self.statusBar().showMessage(
                f"Added '{profile}' to this package — unsaved until Save Changes", 7000
            )
        else:
            self.statusBar().showMessage(f"This package already has Profile '{profile}'.", 5000)

    def _cc_profile_tool_name(self) -> str:
        combo = getattr(self, "cc_magic_profile_quick", None)
        if combo is None or combo.currentIndex() <= 0:
            return ""
        return str(combo.currentData() or combo.currentText() or "").strip()

    def _selected_source_rows(self) -> List[int]:
        selection = self.table.selectionModel() if hasattr(self, "table") else None
        if not selection:
            return []
        rows = []
        seen = set()
        for proxy_index in selection.selectedRows():
            if not proxy_index.isValid():
                continue
            source_row = self.proxy.mapToSource(proxy_index).row()
            if source_row not in seen:
                seen.add(source_row)
                rows.append(source_row)
        return rows

    def _stage_profile_on_rows(self, profile: str, source_rows: List[int], add: bool) -> int:
        if add:
            self._ensure_profile_in_registry(profile)
        wanted = profile.casefold()
        changed_source_rows: List[int] = []
        for source_row in source_rows:
            rec = self.model.record(source_row)
            if not rec:
                continue
            try:
                excel_row = int(rec.get("_ExcelRow", "0") or 0)
            except (TypeError, ValueError):
                excel_row = 0
            if not excel_row or excel_row in self.pending_deletions:
                continue
            current = self.model.effective_value(source_row, "CC Magic Profiles")
            profiles = _split_cc_magic_profiles(current)
            keys = {p.casefold() for p in profiles}
            if add:
                if wanted in keys:
                    continue
                profiles.append(profile)
            else:
                if wanted not in keys:
                    continue
                profiles = [p for p in profiles if p.casefold() != wanted]
            new_value = _normalise_cc_magic_profiles("; ".join(profiles))
            row_pending = dict(self.pending_changes.get(excel_row, {}))
            if new_value == rec.get("CC Magic Profiles", ""):
                row_pending.pop("CC Magic Profiles", None)
            else:
                row_pending["CC Magic Profiles"] = new_value
            if row_pending:
                self.pending_changes[excel_row] = row_pending
            else:
                self.pending_changes.pop(excel_row, None)
            changed_source_rows.append(source_row)

        for source_row in changed_source_rows:
            self.model.notify_row_changed(source_row)
        if self.current_source_row is not None:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)
        self._update_dirty_state()
        self._update_selected_dirty_label()
        self._apply_filters()
        return len(changed_source_rows)

    def _update_profile_assignment_label(self):
        """The assignment control is always plural because multiple Profiles are supported."""
        if hasattr(self, "profiles_assigned_label"):
            self.profiles_assigned_label.setText("Profiles assigned to this package:")

    def _cc_profile_add_to_current(self):
        profile = self._cc_profile_tool_name()
        if not profile or self.current_source_row is None:
            return
        self._stage_profile_on_rows(profile, [self.current_source_row], True)
        if hasattr(self, "cc_magic_profile_quick"):
            self.cc_magic_profile_quick.setCurrentIndex(0)

    def _cc_profile_remove_from_current(self):
        profile = self._cc_profile_tool_name()
        if not profile:
            QMessageBox.information(self, "Profile", "Choose the profile to remove first.")
            return
        profiles = [p for p in _split_cc_magic_profiles(self.cc_magic_profiles_edit.toPlainText()) if p.casefold() != profile.casefold()]
        self.cc_magic_profiles_edit.setPlainText("\n".join(profiles))
        self.cc_magic_profile_quick.setCurrentText(profile)

    def _cc_profile_apply_to_selected(self, add: bool):
        if not self.mdb:
            return
        profile = self._cc_profile_tool_name()
        if not profile:
            QMessageBox.information(self, "Profile", "Choose an existing profile or type a new profile name first.")
            return
        if not self._valid_profile_name(profile):
            QMessageBox.information(self, "Profile", "Use one profile name only (no semicolons or line breaks).")
            return
        source_rows = self._selected_source_rows()
        if not source_rows:
            QMessageBox.information(self, "Profile", "Select one or more packages first. Use Ctrl+click for individual rows or Shift+click for a range.")
            return
        verb = "Add" if add else "Remove"
        relation = "to" if add else "from"
        answer = QMessageBox.question(
            self,
            f"{verb} Profile",
            f"{verb} '{profile}' {relation} the {len(source_rows):,} selected package(s)?\n\n"
            "This only stages unsaved MDB changes. Nothing is written until Save Changes.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        changed = self._stage_profile_on_rows(profile, source_rows, add)
        self.statusBar().showMessage(
            f"Staged profile change for {changed:,} selected package(s) — use Save Changes to write the MDB",
            8000,
        )
        if changed == 0:
            QMessageBox.information(self, "Profile", f"No changes were needed for '{profile}'.")

    def _cc_profile_apply_to_shown(self, add: bool):
        if not self.mdb:
            return
        profile = self._cc_profile_tool_name()
        if not profile:
            QMessageBox.information(self, "Profile", "Choose an existing profile or type a new profile name first.")
            return
        if not self._valid_profile_name(profile):
            QMessageBox.information(self, "Profile", "Use one profile name only (no semicolons or line breaks).")
            return
        shown = self.proxy.rowCount()
        if shown <= 0:
            QMessageBox.information(self, "Profile", "There are no currently shown Library rows to change.")
            return
        verb = "Add" if add else "Remove"
        relation = "to" if add else "from"
        answer = QMessageBox.question(
            self,
            f"{verb} Profile",
            f"{verb} '{profile}' {relation} all {shown:,} package(s) currently shown by the Library filters/search?\n\n"
            "This only stages unsaved MDB changes. Nothing is written until Save Changes.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        source_rows = []
        for proxy_row in range(shown):
            proxy_index = self.proxy.index(proxy_row, 0)
            if proxy_index.isValid():
                source_rows.append(self.proxy.mapToSource(proxy_index).row())
        changed = self._stage_profile_on_rows(profile, source_rows, add)
        self.statusBar().showMessage(
            f"Staged profile change for {changed:,} shown package(s) — use Save Changes to write the MDB",
            8000,
        )
        if changed == 0:
            QMessageBox.information(
                self,
                "Profile",
                f"No changes were needed. The shown packages already {'had' if add else 'did not have'} '{profile}'.",
            )

    def _cc_magic_profile_usage(self) -> List[tuple]:
        """Return registry profiles and package counts, including empty profiles."""
        usage: Dict[str, Dict[str, object]] = {}
        for profile in self._all_profile_names():
            usage[profile.casefold()] = {"name": profile, "count": 0}
        for source_row, rec in enumerate(self.model.records):
            value = self.model.effective_value(source_row, "CC Magic Profiles")
            for profile in _split_cc_magic_profiles(value):
                key = profile.casefold()
                if key not in usage:
                    usage[key] = {"name": profile, "count": 0}
                usage[key]["count"] = int(usage[key]["count"]) + 1
        return sorted(
            [(str(info["name"]), int(info["count"])) for info in usage.values()],
            key=lambda item: item[0].casefold(),
        )

    def _stage_cc_magic_profile_global_change(self, old_profile: str, new_profile: Optional[str]) -> int:
        """Stage profile rename/delete across packages and update the independent registry."""
        old_key = str(old_profile or "").strip().casefold()
        new_clean = str(new_profile or "").strip() if new_profile is not None else ""
        if not old_key:
            return 0

        # Registry change first. Renaming to an existing profile merges cleanly.
        registry = [p for p in self.profile_registry_current if p.casefold() != old_key]
        if new_clean and new_clean.casefold() not in {p.casefold() for p in registry}:
            registry.append(new_clean)
        registry.sort(key=str.casefold)
        self.profile_registry_current = registry
        self.profile_registry_dirty = self.profile_registry_current != self.profile_registry_saved

        changed_rows: List[int] = []
        for source_row, rec in enumerate(self.model.records):
            try:
                excel_row = int(rec.get("_ExcelRow", "0") or 0)
            except (TypeError, ValueError):
                excel_row = 0
            if not excel_row or excel_row in self.pending_deletions:
                continue
            current_profiles = _split_cc_magic_profiles(self.model.effective_value(source_row, "CC Magic Profiles"))
            if old_key not in {p.casefold() for p in current_profiles}:
                continue
            result_profiles: List[str] = []
            for profile in current_profiles:
                if profile.casefold() == old_key:
                    if new_clean:
                        result_profiles.append(new_clean)
                else:
                    result_profiles.append(profile)
            new_value = _normalise_cc_magic_profiles("; ".join(result_profiles))
            try:
                excel_row = int(rec.get("_ExcelRow", "0") or 0)
            except (TypeError, ValueError):
                excel_row = 0
            if not excel_row:
                continue
            row_pending = dict(self.pending_changes.get(excel_row, {}))
            if new_value == rec.get("CC Magic Profiles", ""):
                row_pending.pop("CC Magic Profiles", None)
            else:
                row_pending["CC Magic Profiles"] = new_value
            if row_pending:
                self.pending_changes[excel_row] = row_pending
            else:
                self.pending_changes.pop(excel_row, None)
            changed_rows.append(source_row)

        for source_row in changed_rows:
            self.model.notify_row_changed(source_row)
        self._refresh_cc_magic_profile_choices()
        self._populate_filters()
        if self.current_source_row is not None:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)
        self._update_dirty_state()
        self._update_selected_dirty_label()
        self._apply_filters()
        return len(changed_rows)

    def _manage_cc_magic_profiles(self):
        if not self.mdb:
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Manage Profiles")
        dialog.resize(580, 450)
        layout = QVBoxLayout(dialog)

        help_label = QLabel(
            "Profiles are optional groups for playthroughs, themes, saves, testing setups, or other organisational purposes. "
            "They are independent of File Location. Empty profiles are allowed. Create/rename/delete changes are staged first; "
            "nothing is written until Save Changes in the main window."
        )
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        table = QTableWidget(0, 2)
        table.setHorizontalHeaderLabels(["Profile", "Packages"])
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setStretchLastSection(False)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        layout.addWidget(table, 1)

        status = QLabel("")
        status.setWordWrap(True)
        layout.addWidget(status)

        action_row = QHBoxLayout()
        create_btn = QPushButton("New Profile…")
        rename_btn = QPushButton("Rename selected…")
        delete_btn = QPushButton("Delete selected…")
        action_row.addWidget(create_btn)
        action_row.addWidget(rename_btn)
        action_row.addWidget(delete_btn)
        action_row.addStretch(1)
        layout.addLayout(action_row)

        close_box = QDialogButtonBox(QDialogButtonBox.Close)
        close_box.rejected.connect(dialog.reject)
        layout.addWidget(close_box)

        def refresh_table(select_name: str = ""):
            rows = self._cc_magic_profile_usage()
            table.setRowCount(len(rows))
            wanted = select_name.casefold()
            select_row = -1
            for row, (profile, count) in enumerate(rows):
                name_item = QTableWidgetItem(profile)
                count_item = QTableWidgetItem(f"{count:,}")
                count_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(row, 0, name_item)
                table.setItem(row, 1, count_item)
                if wanted and profile.casefold() == wanted:
                    select_row = row
            if select_row >= 0:
                table.selectRow(select_row)
            elif rows:
                table.selectRow(0)
            rename_btn.setEnabled(bool(rows))
            delete_btn.setEnabled(bool(rows))
            registry_note = "registry changed" if self.profile_registry_dirty else "registry saved"
            status.setText(
                f"{len(rows):,} profile(s), including empty profiles. "
                f"Unsaved package fields: {sum(len(v) for v in self.pending_changes.values()):,}; {registry_note}."
            )

        def selected_profile() -> str:
            row = table.currentRow()
            item = table.item(row, 0) if row >= 0 else None
            return item.text().strip() if item else ""

        def create_profile_here():
            name, ok = QInputDialog.getText(dialog, "New Profile", "Profile name:")
            if not ok:
                return
            name = str(name or "").strip()
            if not name or not self._valid_profile_name(name):
                QMessageBox.information(dialog, "New Profile", "Enter one non-blank profile name with no semicolons or line breaks.")
                return
            existing = next((p for p in self._all_profile_names() if p.casefold() == name.casefold()), None)
            if existing:
                QMessageBox.information(dialog, "New Profile", f"'{existing}' already exists.")
                refresh_table(existing)
                return
            self._ensure_profile_in_registry(name)
            refresh_table(name)

        def rename_selected():
            old_name = selected_profile()
            if not old_name:
                return
            new_name, ok = QInputDialog.getText(dialog, "Rename Profile", f"Rename '{old_name}' to:", QLineEdit.Normal, old_name)
            if not ok:
                return
            new_name = str(new_name or "").strip()
            if not new_name:
                QMessageBox.information(dialog, "Rename Profile", "Profile name cannot be blank.")
                return
            if not self._valid_profile_name(new_name):
                QMessageBox.information(dialog, "Rename Profile", "Use one profile name only (no semicolons or line breaks).")
                return
            if new_name == old_name:
                return
            usage = dict((name.casefold(), count) for name, count in self._cc_magic_profile_usage())
            count = usage.get(old_name.casefold(), 0)
            merge_note = ""
            if new_name.casefold() != old_name.casefold() and new_name.casefold() in usage:
                merge_note = f"\n\n'{new_name}' already exists. Memberships will be merged without duplicates."
            answer = QMessageBox.question(
                dialog,
                "Rename Profile",
                f"Rename '{old_name}' to '{new_name}' across {count:,} package(s)?{merge_note}\n\n"
                "This stages the changes only. Use Save Changes in the main window to write them to the MDB.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
            changed = self._stage_cc_magic_profile_global_change(old_name, new_name)
            refresh_table(new_name)
            status.setText(f"Staged rename on {changed:,} package(s). Save Changes when ready.")

        def delete_selected():
            old_name = selected_profile()
            if not old_name:
                return
            usage = dict((name.casefold(), count) for name, count in self._cc_magic_profile_usage())
            count = usage.get(old_name.casefold(), 0)
            answer = QMessageBox.question(
                dialog,
                "Delete Profile",
                f"Delete profile '{old_name}' and remove it from {count:,} package(s)?\n\n"
                "This does NOT move or delete package files. The change is staged until Save Changes.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
            changed = self._stage_cc_magic_profile_global_change(old_name, None)
            refresh_table()
            status.setText(f"Staged profile deletion from {changed:,} package(s). Save Changes when ready.")

        create_btn.clicked.connect(create_profile_here)
        rename_btn.clicked.connect(rename_selected)
        delete_btn.clicked.connect(delete_selected)
        refresh_table()
        dialog.exec()

    def _all_merge_profile_names(self) -> List[str]:
        """Return the managed Merge Profile registry plus any assigned legacy names."""
        result: List[str] = []
        seen = set()
        for raw in self.merge_profile_registry_current:
            name = str(raw or "").strip()
            if name and name.casefold() not in seen:
                seen.add(name.casefold())
                result.append(name)
        for source_row, _rec in enumerate(self.model.records):
            for name in _split_cc_magic_profiles(self.model.effective_value(source_row, "Merge Profiles")):
                if name.casefold() in seen:
                    continue
                seen.add(name.casefold())
                result.append(name)
        return sorted(result, key=str.casefold)

    def _merge_profile_usage(self) -> List[tuple[str, int]]:
        usage = {name.casefold(): [name, 0] for name in self._all_merge_profile_names()}
        for source_row, _rec in enumerate(self.model.records):
            for name in _split_cc_magic_profiles(self.model.effective_value(source_row, "Merge Profiles")):
                key = name.casefold()
                usage.setdefault(key, [name, 0])
                usage[key][1] += 1
        return sorted([(str(v[0]), int(v[1])) for v in usage.values()], key=lambda item: item[0].casefold())

    def _valid_merge_profile_name(self, name: str) -> bool:
        clean = str(name or "").strip()
        return bool(clean and ";" not in clean and "\n" not in clean and "\r" not in clean)

    def _ensure_merge_profile_in_registry(self, name: str) -> bool:
        name = str(name or "").strip()
        if not self._valid_merge_profile_name(name):
            return False
        if name.casefold() in {p.casefold() for p in self.merge_profile_registry_current}:
            return False
        self.merge_profile_registry_current.append(name)
        self.merge_profile_registry_current.sort(key=str.casefold)
        self.merge_profile_registry_dirty = self.merge_profile_registry_current != self.merge_profile_registry_saved
        self._update_dirty_state()
        return True

    def _stage_merge_profiles_for_current(self, profiles: List[str]) -> bool:
        if not self.mdb or self.current_source_row is None:
            return False
        rec = self.model.record(self.current_source_row)
        if not rec:
            return False
        excel_row = self.model._excel_row(rec)
        if not excel_row or excel_row in self.pending_deletions:
            return False
        clean_profiles = []
        seen = set()
        for raw in profiles:
            name = str(raw or "").strip()
            if not self._valid_merge_profile_name(name) or name.casefold() in seen:
                continue
            seen.add(name.casefold())
            clean_profiles.append(name)
            if name.casefold() not in {p.casefold() for p in self.merge_profile_registry_current}:
                self.merge_profile_registry_current.append(name)
        self.merge_profile_registry_current.sort(key=str.casefold)
        self.merge_profile_registry_dirty = self.merge_profile_registry_current != self.merge_profile_registry_saved
        new_value = _normalise_cc_magic_profiles("; ".join(clean_profiles))
        row_pending = dict(self.pending_changes.get(excel_row, {}))
        if new_value == rec.get("Merge Profiles", ""):
            row_pending.pop("Merge Profiles", None)
        else:
            row_pending["Merge Profiles"] = new_value
        if row_pending:
            self.pending_changes[excel_row] = row_pending
        else:
            self.pending_changes.pop(excel_row, None)
        self.model.notify_row_changed(self.current_source_row)
        self.merge_profiles_edit.setPlainText(new_value.replace("; ", "\n"))
        self._update_dirty_state()
        self._update_selected_dirty_label()
        self._apply_filters()
        return True

    def _set_merge_profiles_row_visible(self, status_text: Optional[str] = None):
        if not hasattr(self, "merge_profiles_tools"):
            return
        # Merge membership is represented directly by Merge Profile assignments.
        # The row is therefore available for every selected package, regardless of Status.
        visible = self.current_source_row is not None
        try:
            self.edit_form.setRowVisible(self.merge_profiles_tools, visible)
        except AttributeError:
            self.merge_profiles_label.setVisible(visible)
            self.merge_profiles_tools.setVisible(visible)
        enabled = visible and self._current_excel_row() not in self.pending_deletions
        self.merge_profiles_set_btn.setEnabled(enabled)
        self.merge_profiles_manage_btn.setEnabled(bool(self.mdb))
        self._update_dynamic_editor_heights()

    def _status_editor_changed(self, *_):
        self._editor_changed()

    def _choose_merge_profiles_for_current(self):
        if not self.mdb or self.current_source_row is None:
            return
        names = self._all_merge_profile_names()
        if not names:
            QMessageBox.information(
                self,
                "Merge Profiles",
                "No Merge Profiles have been created yet. Use Manage Merge Profiles… to create the list first.",
            )
            return
        current = {
            p.casefold() for p in _split_cc_magic_profiles(
                self.model.effective_value(self.current_source_row, "Merge Profiles")
            )
        }
        dialog = QDialog(self)
        dialog.setWindowTitle("Set Merge Profiles")
        dialog.resize(460, 520)
        layout = QVBoxLayout(dialog)
        note = QLabel(
            "Tick every merged collection that contains this package. A package can belong to several Merge Profiles at once. "
            "These are separate from general Profiles."
        )
        note.setWordWrap(True)
        layout.addWidget(note)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        holder = QWidget()
        holder_layout = QVBoxLayout(holder)
        checks: Dict[str, QCheckBox] = {}
        for name in names:
            cb = QCheckBox(name)
            cb.setChecked(name.casefold() in current)
            checks[name] = cb
            holder_layout.addWidget(cb)
        holder_layout.addStretch(1)
        scroll.setWidget(holder)
        layout.addWidget(scroll, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.Accepted:
            return
        selected = [name for name, cb in checks.items() if cb.isChecked()]
        self._stage_merge_profiles_for_current(selected)
        self.statusBar().showMessage(
            f"Staged {len(selected):,} Merge Profile assignment(s) for this package — Save Changes to write the MDB",
            8000,
        )

    def _stage_merge_profile_global_change(self, old_name: str, new_name: Optional[str]) -> int:
        old_key = str(old_name or "").strip().casefold()
        new_clean = str(new_name or "").strip() if new_name is not None else ""
        if not old_key:
            return 0
        registry = [p for p in self.merge_profile_registry_current if p.casefold() != old_key]
        if new_clean and new_clean.casefold() not in {p.casefold() for p in registry}:
            registry.append(new_clean)
        registry.sort(key=str.casefold)
        self.merge_profile_registry_current = registry
        self.merge_profile_registry_dirty = registry != self.merge_profile_registry_saved
        changed_rows = []
        for source_row, rec in enumerate(self.model.records):
            excel_row = self.model._excel_row(rec)
            if not excel_row or excel_row in self.pending_deletions:
                continue
            current = _split_cc_magic_profiles(self.model.effective_value(source_row, "Merge Profiles"))
            if old_key not in {p.casefold() for p in current}:
                continue
            updated = []
            for profile in current:
                if profile.casefold() == old_key:
                    if new_clean:
                        updated.append(new_clean)
                else:
                    updated.append(profile)
            new_value = _normalise_cc_magic_profiles("; ".join(updated))
            row_pending = dict(self.pending_changes.get(excel_row, {}))
            if new_value == rec.get("Merge Profiles", ""):
                row_pending.pop("Merge Profiles", None)
            else:
                row_pending["Merge Profiles"] = new_value
            if row_pending:
                self.pending_changes[excel_row] = row_pending
            else:
                self.pending_changes.pop(excel_row, None)
            changed_rows.append(source_row)
        for source_row in changed_rows:
            self.model.notify_row_changed(source_row)
        if self.current_source_row is not None:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)
        self._update_dirty_state()
        self._update_selected_dirty_label()
        self._apply_filters()
        return len(changed_rows)

    def _manage_merge_profiles(self):
        if not self.mdb:
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Manage Merge Profiles")
        dialog.resize(580, 450)
        layout = QVBoxLayout(dialog)
        help_label = QLabel(
            "Merge Profiles are names for merged collections. They are separate from general Profiles, and one package can belong to several. "
            "Create, rename and remove names here. Changes remain staged until Save Changes."
        )
        help_label.setWordWrap(True)
        layout.addWidget(help_label)
        table = QTableWidget(0, 2)
        table.setHorizontalHeaderLabels(["Merge Profile", "Packages"])
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        layout.addWidget(table, 1)
        status = QLabel("")
        status.setWordWrap(True)
        layout.addWidget(status)
        row = QHBoxLayout()
        create_btn = QPushButton("New Merge Profile…")
        rename_btn = QPushButton("Rename selected…")
        delete_btn = QPushButton("Delete selected…")
        row.addWidget(create_btn)
        row.addWidget(rename_btn)
        row.addWidget(delete_btn)
        row.addStretch(1)
        layout.addLayout(row)
        close_box = QDialogButtonBox(QDialogButtonBox.Close)
        close_box.rejected.connect(dialog.reject)
        layout.addWidget(close_box)

        def selected_name() -> str:
            r = table.currentRow()
            item = table.item(r, 0) if r >= 0 else None
            return item.text().strip() if item else ""

        def refresh(select: str = ""):
            rows = self._merge_profile_usage()
            table.setRowCount(len(rows))
            select_row = -1
            for i, (name, count) in enumerate(rows):
                table.setItem(i, 0, QTableWidgetItem(name))
                count_item = QTableWidgetItem(f"{count:,}")
                count_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(i, 1, count_item)
                if select and name.casefold() == select.casefold():
                    select_row = i
            if select_row >= 0:
                table.selectRow(select_row)
            elif rows:
                table.selectRow(0)
            rename_btn.setEnabled(bool(rows))
            delete_btn.setEnabled(bool(rows))
            status.setText(
                f"{len(rows):,} Merge Profile(s). Registry {'changed' if self.merge_profile_registry_dirty else 'saved'}."
            )

        def create_profile():
            name, ok = QInputDialog.getText(dialog, "New Merge Profile", "Merge Profile name:")
            if not ok:
                return
            name = str(name or "").strip()
            if not self._valid_merge_profile_name(name):
                QMessageBox.information(dialog, "Merge Profile", "Enter one non-blank name with no semicolons or line breaks.")
                return
            existing = next((p for p in self._all_merge_profile_names() if p.casefold() == name.casefold()), None)
            if existing:
                QMessageBox.information(dialog, "Merge Profile", f"'{existing}' already exists.")
                refresh(existing)
                return
            self._ensure_merge_profile_in_registry(name)
            refresh(name)

        def rename_profile():
            old = selected_name()
            if not old:
                return
            new, ok = QInputDialog.getText(dialog, "Rename Merge Profile", f"Rename '{old}' to:", QLineEdit.Normal, old)
            if not ok:
                return
            new = str(new or "").strip()
            if not self._valid_merge_profile_name(new):
                QMessageBox.information(dialog, "Merge Profile", "Enter one non-blank name with no semicolons or line breaks.")
                return
            if new == old:
                return
            usage = dict((name.casefold(), count) for name, count in self._merge_profile_usage())
            count = usage.get(old.casefold(), 0)
            merge_note = ""
            if new.casefold() in usage and new.casefold() != old.casefold():
                merge_note = f"\n\n'{new}' already exists. Memberships will be merged without duplicates."
            answer = QMessageBox.question(
                dialog,
                "Rename Merge Profile",
                f"Rename '{old}' to '{new}' across {count:,} package(s)?{merge_note}\n\nThe change stays staged until Save Changes.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
            changed = self._stage_merge_profile_global_change(old, new)
            refresh(new)
            status.setText(f"Staged rename on {changed:,} package(s). Save Changes when ready.")

        def delete_profile():
            old = selected_name()
            if not old:
                return
            usage = dict((name.casefold(), count) for name, count in self._merge_profile_usage())
            count = usage.get(old.casefold(), 0)
            answer = QMessageBox.question(
                dialog,
                "Delete Merge Profile",
                f"Delete '{old}' and remove it from {count:,} package(s)?\n\nThis does not move or delete package files. The change stays staged until Save Changes.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
            changed = self._stage_merge_profile_global_change(old, None)
            refresh()
            status.setText(f"Staged removal from {changed:,} package(s). Save Changes when ready.")

        create_btn.clicked.connect(create_profile)
        rename_btn.clicked.connect(rename_profile)
        delete_btn.clicked.connect(delete_profile)
        refresh()
        dialog.exec()

    def _current_excel_row(self) -> Optional[int]:
        if self.current_source_row is None:
            return None
        rec = self.model.record(self.current_source_row)
        try:
            return int(rec.get("_ExcelRow", "0") or 0) or None
        except ValueError:
            return None

    def _ensure_combo_value(self, combo: QComboBox, value: str):
        if value and combo.findText(value) < 0:
            combo.addItem(value)
        combo.setCurrentText(value)

    def _load_editor_for_record(self, rec: Dict[str, str]):
        row_num = int(rec.get("_ExcelRow", "0") or 0)
        pending = self.pending_changes.get(row_num, {})
        self._loading_editor = True
        try:
            self._ensure_combo_value(self.status_edit, pending.get("Status", rec.get("Status", "")))
            self.creator_edit.setText(pending.get("Creator", rec.get("Creator", "")))
            self._ensure_combo_value(self.file_location_edit, pending.get("File Location", rec.get("File Location", "")))
            profile_value = pending.get("CC Magic Profiles", rec.get("CC Magic Profiles", ""))
            self.cc_magic_profiles_edit.setPlainText(
                _normalise_cc_magic_profiles(profile_value).replace("; ", "\n")
            )
            merge_profile_value = pending.get("Merge Profiles", rec.get("Merge Profiles", ""))
            self.merge_profiles_edit.setPlainText(
                _normalise_cc_magic_profiles(merge_profile_value).replace("; ", "\n")
            )
            self.notes_edit.setPlainText(pending.get("Notes", rec.get("Notes", "")))
            self.file_location_open_btn.setEnabled(bool(pending.get("File Location", rec.get("File Location", ""))))
        finally:
            self._loading_editor = False
        self._set_editor_enabled(row_num not in self.pending_deletions)
        self._set_merge_profiles_row_visible(pending.get("Status", rec.get("Status", "")))
        self._update_selected_dirty_label()

    def _editor_changed(self, *_):
        if self._loading_editor or not self.mdb or self.current_source_row is None:
            return
        rec = self.model.record(self.current_source_row)
        if not rec:
            return
        row_num = int(rec.get("_ExcelRow", "0") or 0)
        if not row_num or row_num in self.pending_deletions:
            return
        values = {
            "Status": self.status_edit.currentText(),
            "Creator": self.creator_edit.text().strip(),
            "File Location": self.file_location_edit.currentText(),
            "Notes": self.notes_edit.toPlainText(),
        }
        row_pending = dict(self.pending_changes.get(row_num, {}))
        for field, value in values.items():
            if value == rec.get(field, ""):
                row_pending.pop(field, None)
            else:
                row_pending[field] = value
        if row_pending:
            self.pending_changes[row_num] = row_pending
        else:
            self.pending_changes.pop(row_num, None)
        self.status_badge.setText("Status: " + (values["Status"] or "—"))
        self.file_location_open_btn.setEnabled(bool(values["File Location"].strip()))
        self.model.notify_row_changed(self.current_source_row)
        self._update_dirty_state()
        self._update_selected_dirty_label()
        self._apply_filters()

    def _set_pending_image_value(self, value: str):
        if not self.mdb or self.current_source_row is None:
            return
        rec = self.model.record(self.current_source_row)
        if not rec:
            return
        row_num = int(rec.get("_ExcelRow", "0") or 0)
        if not row_num or row_num in self.pending_deletions:
            return
        row_pending = dict(self.pending_changes.get(row_num, {}))
        original = rec.get("Image", "")
        if value == original:
            row_pending.pop("Image", None)
        else:
            row_pending["Image"] = value
        if row_pending:
            self.pending_changes[row_num] = row_pending
        else:
            self.pending_changes.pop(row_num, None)
        self.model.notify_row_changed(self.current_source_row)
        self.show_record(rec)
        self._update_dirty_state()
        self._update_selected_dirty_label()

    def _choose_manual_preview_image(self):
        if not self.mdb or self.current_source_row is None:
            return
        start = str(self.settings.value("last_image_dir", str(self.mdb.path.parent)) or self.mdb.path.parent)
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose Preview Image",
            start,
            "Images (*.png *.jpg *.jpeg *.bmp *.gif *.webp);;All files (*.*)",
        )
        if not path:
            return
        pix = QPixmap(path)
        if pix.isNull():
            QMessageBox.warning(self, "Preview image", "That file could not be opened as an image.")
            return
        chosen = str(Path(path).resolve())
        self.settings.setValue("last_image_dir", str(Path(chosen).parent))
        self._set_pending_image_value(chosen)
        self.statusBar().showMessage("Preview image changed — Save Changes to write it to the MDB", 7000)

    def _clear_manual_preview_image(self):
        if not self.mdb or self.current_source_row is None:
            return
        self._set_pending_image_value("")
        self.statusBar().showMessage("Preview image cleared — Save Changes to write it to the MDB; Scan can restore an extractable package preview", 9000)

    def _update_selected_dirty_label(self):
        if not hasattr(self, "edit_dirty_label"):
            return
        row_num = self._current_excel_row()
        fields = list(self.pending_changes.get(row_num or -1, {}).keys())
        if row_num in self.pending_deletions:
            self.edit_dirty_label.setText("Record staged for deletion — use Revert this to cancel")
        elif fields:
            self.edit_dirty_label.setText("Unsaved for this package: " + ", ".join(fields))
        else:
            self.edit_dirty_label.setText("No unsaved changes for this package")
        if hasattr(self, "revert_selected_btn"):
            self.revert_selected_btn.setEnabled(bool(fields) or row_num in self.pending_deletions)

    def _update_dirty_state(self):
        field_count = sum(len(v) for v in self.pending_changes.values())
        dirty = (
            field_count > 0 or bool(self.pending_deletions) or self.profile_registry_dirty
            or self.status_registry_dirty or self.file_location_registry_dirty
            or self.merge_profile_registry_dirty
        )
        registry_extra = (
            int(self.profile_registry_dirty) + int(self.status_registry_dirty)
            + int(self.file_location_registry_dirty) + int(self.merge_profile_registry_dirty)
        )
        if hasattr(self, "unsaved_label"):
            self.unsaved_label.setText(f"Unsaved: {field_count + len(self.pending_deletions) + registry_extra}")
        if hasattr(self, "save_action"):
            self.save_action.setEnabled(dirty)
        if hasattr(self, "revert_all_action"):
            self.revert_all_action.setEnabled(dirty)
        if hasattr(self, "save_btn"):
            self.save_btn.setEnabled(dirty)
        if self.mdb:
            marker = " *" if dirty else ""
            self.setWindowTitle(f"{APP_NAME} — {self.mdb.path.name}{marker} — {VERSION}")
        else:
            self.setWindowTitle(f"{APP_NAME} — {VERSION}")

    def revert_selected_changes(self):
        row_num = self._current_excel_row()
        if not row_num:
            return
        had_change = row_num in self.pending_changes or row_num in self.pending_deletions
        if not had_change:
            return
        self.pending_changes.pop(row_num, None)
        self.pending_deletions.discard(row_num)
        rec = self.model.record(self.current_source_row) if self.current_source_row is not None else {}
        if rec:
            self.show_record(rec)
        if self.current_source_row is not None:
            self.model.notify_row_changed(self.current_source_row)
        self._update_dirty_state()
        self._apply_filters()

    def revert_all_changes(self):
        if (not self.pending_changes and not self.pending_deletions and not self.profile_registry_dirty
                and not self.status_registry_dirty and not self.file_location_registry_dirty
                and not self.merge_profile_registry_dirty):
            return
        answer = QMessageBox.question(
            self,
            "Revert unsaved changes",
            "Discard all unsaved Library/Profile/Merge Profile/Status/File Location changes, including staged record deletions?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        self.pending_changes.clear()
        self.pending_deletions.clear()
        self.profile_registry_current = list(self.profile_registry_saved)
        self.profile_registry_dirty = False
        self.status_registry_dirty = False
        self.file_location_registry_current = list(self.file_location_registry_saved)
        self.file_location_registry_dirty = False
        self.merge_profile_registry_current = list(self.merge_profile_registry_saved)
        self.merge_profile_registry_dirty = False
        self._populate_edit_options()
        self._refresh_cc_magic_profile_choices()
        self.model.notify_all_changed()
        if self.current_source_row is not None:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)
        self._update_dirty_state()
        self._apply_filters()

    def review_legacy_status_values(self):
        """Preview and stage safe conversions from legacy Status labels.

        Nothing is written by this action. The user must still confirm the normal
        Save Changes workflow, which creates the workbook backup before writing.
        Legacy Delete is deliberately excluded because scanner behaviour may depend
        on it. Existing pending Status edits are included in the preview.
        """
        if not self.mdb:
            QMessageBox.information(
                self, "Review legacy Status values",
                "Open a Master Database first."
            )
            return

        conversions = {
            "keep": ("To Keep", "Keep → To Keep"),
            "review": ("To Review", "Review → To Review"),
            "favourite": ("", "Favourite → blank"),
            "unreviewed": ("", "Unreviewed → blank"),
            # In Merge existed only in unreleased RC builds. Merge membership is
            # now represented directly by Merge Profiles, so the redundant Status
            # can be safely cleared while keeping all Merge Profile assignments.
            "in merge": ("", "In Merge → blank"),
        }
        counts = {key: 0 for key in conversions}
        rows_to_stage = []

        for rec in self.model.records:
            try:
                excel_row = int(rec.get("_ExcelRow", "0") or 0)
            except (TypeError, ValueError):
                continue
            if excel_row <= 0 or excel_row in self.pending_deletions:
                continue
            pending = self.pending_changes.get(excel_row, {})
            current = str(pending.get("Status", rec.get("Status", "")) or "").strip()
            key = current.casefold()
            if key in conversions:
                counts[key] += 1
                rows_to_stage.append((excel_row, conversions[key][0], key))

        registry_differs = [
            str(value).strip().casefold() for value in (self.mdb.status_options or [])
        ] != [value.casefold() for value in PUBLIC_STATUS_OPTIONS]
        if not any(counts.values()) and not registry_differs:
            QMessageBox.information(
                self, "Review legacy Status values",
                "No legacy Keep, Review, Favourite, Unreviewed or In Merge Status values were found, "
                "and the Settings Status list already matches the current choices. "
                "Delete/Deleted values are intentionally left untouched."
            )
            return

        preview_lines = [
            "PREVIEW ONLY — nothing is written to the workbook yet.",
            "",
            f"Keep → To Keep: {counts['keep']} record(s)",
            f"Review → To Review: {counts['review']} record(s)",
            f"Favourite → blank: {counts['favourite']} record(s)",
            f"Unreviewed → blank: {counts['unreviewed']} record(s)",
            f"In Merge → blank: {counts['in merge']} record(s) (Merge Profile assignments are kept)",
            "",
            "Settings Status list: " + (
                "will be updated to the three current choices."
                if registry_differs else "already matches the current choices."
            ),
            "",
            "Legacy Delete/Deleted values will not be changed because they may have "
            "special scanner meaning.",
            "",
            "If you continue, these record changes will be staged in the app and the "
            "Settings Status list will be updated on Save Changes. The normal full-workbook "
            "backup will be created before writing.",
            "",
            "Stage these conversions?",
        ]
        answer = QMessageBox.question(
            self,
            "Preview legacy Status migration",
            "\\n".join(preview_lines),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            self.statusBar().showMessage("Status migration preview cancelled — nothing changed")
            return

        staged = 0
        if registry_differs:
            self.status_registry_dirty = True
        for excel_row, new_value, _key in rows_to_stage:
            fields = self.pending_changes.setdefault(excel_row, {})
            if fields.get("Status") != new_value:
                fields["Status"] = new_value
                staged += 1

        # Refresh records and any open editor from the staged values.
        for source_row, rec in enumerate(self.model.records):
            try:
                excel_row = int(rec.get("_ExcelRow", "0") or 0)
            except (TypeError, ValueError):
                continue
            fields = self.pending_changes.get(excel_row, {})
            if "Status" in fields:
                self.model.notify_row_changed(source_row)
        if self.current_source_row is not None:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)
        self._populate_filters()
        self._apply_filters()
        self._update_dirty_state()
        self._update_selected_dirty_label()
        self.statusBar().showMessage(
            f"Staged {staged} legacy Status conversion(s). Review them and use Save Changes to write.",
            10000,
        )

    def save_changes(self) -> bool:
        if not self.mdb or (
            not self.pending_changes
            and not self.pending_deletions
            and not self.profile_registry_dirty
            and not self.status_registry_dirty
            and not self.file_location_registry_dirty
            and not self.merge_profile_registry_dirty
        ):
            return True

        target_path = str(self.mdb.path.resolve())
        # No repetitive Save confirmation in PUBLIC RC4. The writer still creates
        # the same full MDB backup before every successful save. Delete Record has
        # its own explicit staging confirmation before it reaches this point.

        profile_registry_needs_save = self.profile_registry_dirty or any(
            "CC Magic Profiles" in fields for fields in self.pending_changes.values()
        )
        merge_profile_registry_needs_save = self.merge_profile_registry_dirty or any(
            "Merge Profiles" in fields for fields in self.pending_changes.values()
        )
        status_registry_was_dirty = self.status_registry_dirty
        profile_registry_was_dirty = profile_registry_needs_save
        merge_profile_registry_was_dirty = merge_profile_registry_needs_save
        file_location_registry_was_dirty = self.file_location_registry_dirty
        self.statusBar().showMessage("Saving Library/Profile changes and creating MDB backup…")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            result = MDBWorkbook().save_manual_changes(
                self.mdb,
                self.pending_changes,
                profile_options=self.profile_registry_current if profile_registry_needs_save else None,
                status_options=PUBLIC_STATUS_OPTIONS if status_registry_was_dirty else None,
                delete_rows=self.pending_deletions,
                file_locations=self.file_location_registry_current if file_location_registry_was_dirty else None,
                merge_profile_options=self.merge_profile_registry_current if merge_profile_registry_needs_save else None,
            )
        except WorkbookChangedError as exc:
            QMessageBox.warning(self, "Master Database changed", str(exc))
            self.statusBar().showMessage("Save cancelled — reload the Master Database first")
            return False
        except Exception as exc:
            QMessageBox.critical(self, "Could not save changes", str(exc))
            self.statusBar().showMessage("Save failed — the original Master Database was left unchanged")
            return False
        finally:
            QApplication.restoreOverrideCursor()

        deleted_rows = set(self.pending_deletions)
        changes_copy = {row: dict(fields) for row, fields in self.pending_changes.items()}
        for source_row, rec in enumerate(self.model.records):
            try:
                excel_row = int(rec.get("_ExcelRow", "0") or 0)
            except (TypeError, ValueError):
                continue
            row_changes = changes_copy.get(excel_row)
            if row_changes and excel_row not in deleted_rows:
                rec.update(row_changes)
                self.model.notify_row_changed(source_row)

        if deleted_rows:
            deleted_records = [
                rec for rec in self.mdb.records
                if int(rec.get("_ExcelRow", "0") or 0) in deleted_rows
            ]
            deleted_thumbnail_count = sum(1 for rec in deleted_records if rec.get("Thumbnail") == "Yes")
            deleted_pin_ids = {self._record_pin_id(rec) for rec in deleted_records}
            self.pinned_packages = [pin for pin in self.pinned_packages if pin.get("id") not in deleted_pin_ids]
            self._save_pins_for_current_mdb()
            self.mdb.records[:] = [
                rec for rec in self.mdb.records
                if int(rec.get("_ExcelRow", "0") or 0) not in deleted_rows
            ]
            self.mdb.thumbnail_count = max(0, self.mdb.thumbnail_count - deleted_thumbnail_count)
            self.model.set_records(self.mdb.records)
            self.current_source_row = None
            self.table.clearSelection()
            self.packages_label.setText(f"Packages: {len(self.mdb.records):,}")
            self.thumbs_label.setText(f"Thumbnails: {self.mdb.thumbnail_count:,}")

        self.pending_changes.clear()
        self.pending_deletions.clear()
        self.profile_registry_saved = list(self.profile_registry_current)
        self.profile_registry_dirty = False
        self.status_registry_dirty = False
        self.file_location_registry_saved = list(self.file_location_registry_current)
        self.file_location_registry_dirty = False
        self.merge_profile_registry_saved = list(self.merge_profile_registry_current)
        self.merge_profile_registry_dirty = False
        self._populate_edit_options()
        self._refresh_cc_magic_profile_choices()
        self._populate_filters()
        self._apply_filters()
        self._update_dirty_state()
        self._update_selected_dirty_label()
        if self.current_source_row is not None:
            rec = self.model.record(self.current_source_row)
            if rec:
                self.show_record(rec)

        registry_note = " Profile registry saved." if profile_registry_was_dirty else ""
        if status_registry_was_dirty:
            registry_note += " Status options updated."
        if file_location_registry_was_dirty:
            registry_note += " File Location registry saved."
        if merge_profile_registry_was_dirty:
            registry_note += " Merge Profile registry saved."
        delete_summary = f" Deleted {result.deleted_rows} library record(s)." if result.deleted_rows else ""
        self.statusBar().showMessage(
            f"Saved {result.changed_fields} package field change(s) across {result.changed_rows} package(s)."
            f"{delete_summary}{registry_note} Backup: {result.backup_path.name}",
            12000,
        )
        return True

    def _confirm_pending_changes(self) -> bool:
        if (not self.pending_changes and not self.pending_deletions and not self.profile_registry_dirty
                and not self.status_registry_dirty and not self.file_location_registry_dirty
                and not self.merge_profile_registry_dirty):
            return True
        box = QMessageBox(self)
        box.setWindowTitle("Unsaved changes")
        box.setText("You have unsaved Library/Profile/Merge Profile/Status/File Location changes, including any staged record deletions.")
        box.setInformativeText("Save them before continuing?")
        box.setStandardButtons(QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
        box.setDefaultButton(QMessageBox.Save)
        choice = box.exec()
        if choice == QMessageBox.Save:
            return self.save_changes()
        if choice == QMessageBox.Discard:
            self.pending_changes.clear()
            self.pending_deletions.clear()
            self.profile_registry_current = list(self.profile_registry_saved)
            self.profile_registry_dirty = False
            self.status_registry_dirty = False
            self.file_location_registry_current = list(self.file_location_registry_saved)
            self.file_location_registry_dirty = False
            self.merge_profile_registry_current = list(self.merge_profile_registry_saved)
            self.merge_profile_registry_dirty = False
            self._populate_edit_options()
            self._refresh_cc_magic_profile_choices()
            self.model.notify_all_changed()
            self._populate_filters()
            if self.current_source_row is not None:
                rec = self.model.record(self.current_source_row)
                if rec:
                    self.show_record(rec)
            self._update_dirty_state()
            self._update_selected_dirty_label()
            self._apply_filters()
            return True
        return False

    def closeEvent(self, event):
        if self.scan_process is not None and self.scan_process.state() != QProcess.NotRunning:
            QMessageBox.warning(
                self,
                "Scanner is running",
                "Wait for the scanner to finish before closing the app. "
                "This protects the MDB from being interrupted during a save."
            )
            event.ignore()
            return
        if self._confirm_pending_changes():
            event.accept()
        else:
            event.ignore()

    def _show_thumbnail(self, rec: Dict[str, str]):
        target = rec.get("_ThumbnailTarget", "")
        path = target_to_local_path(target) if target else None
        if path is None or not path.exists():
            self.thumbnail_label.setToolTip("")
            self.thumbnail_label.clear_thumbnail("No thumbnail found")
            return
        pix = QPixmap(str(path))
        if pix.isNull():
            self.thumbnail_label.setToolTip("")
            self.thumbnail_label.clear_thumbnail("Thumbnail could not be opened")
            return
        # Keep the full-resolution source and let ResponsiveThumbnailLabel refit
        # it whenever the window or splitter changes size. No fixed cap, no stale
        # oversized pixmap, and no viewer-side cropping.
        self.thumbnail_label.setToolTip(
            f"Stored preview: {pix.width()} × {pix.height()} px. CAS Curator fits the complete stored image without cropping."
        )
        self.thumbnail_label.set_thumbnail(pix)


def main():
    # Give Windows a stable application identity so the custom CAS Curator icon
    # is used consistently for the running window/taskbar instead of inheriting
    # a generic Python/Qt identity.
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("CASCurator.CASCurator")
        except Exception:
            pass
    app = QApplication(sys.argv)
    if "Fusion" in QStyleFactory.keys():
        app.setStyle("Fusion")
    app.setOrganizationName("CASCurator")
    app.setApplicationName("CAS Curator")
    icon_path = Path(__file__).resolve().parent / "assets" / "cas_curator.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
