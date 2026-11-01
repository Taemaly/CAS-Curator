"""Theme Colour Studio. Deliberately imports no scanner/MDB modules.

Edits are transient until the user explicitly chooses Save As or Update.
The rendered mock Library / Scan / Guide pages are live Qt widgets using
exactly the custom style role keys consumed by app.py, not a static image.
"""
from __future__ import annotations

import copy
import random
import shutil
from datetime import datetime

from PySide6.QtCore import Qt, QMimeData, QSize, QTimer
from PySide6.QtGui import QDrag, QFont, QIcon, QPixmap, QPainter, QColor, QPalette
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QFrame, QGridLayout,
    QHBoxLayout, QInputDialog, QLabel, QMessageBox, QPushButton,
    QScrollArea, QSizePolicy, QStyle, QStyleOptionButton, QTabWidget, QVBoxLayout, QWidget,
)

from theme_studio_core import (
    MAX_IMPORT_COLOURS, ROLE_GROUPS, ROLES, add_custom_colour,
    automatic_text_for, edited_theme, mix, normalise_palette, readability_report, resolve_preview_roles, shuffle_roles, text_for,
)
from colour_wheel import ColourWheelDialog
from custom_colour_storage import (
    MAX_SAVED_COLOURS, add_saved_colour, default_saved_colour_file,
    load_saved_colours, remove_saved_colour, save_saved_colours,
)




def _studio_standard_icon(widget: QWidget, *names: str) -> QIcon:
    """Return the first available native Qt icon from the CAS Curator-style set."""
    for name in names:
        enum = getattr(QStyle.StandardPixmap, name, None)
        if enum is not None:
            icon = widget.style().standardIcon(enum)
            if not icon.isNull():
                return icon
    return QIcon()


def _set_button_icon(button: QPushButton, *names: str) -> None:
    icon = _studio_standard_icon(button, *names)
    if not icon.isNull():
        button.setIcon(icon)
        button.setIconSize(QSize(16, 16))


def _colour_action_icon(colour: str = '#6B94B3') -> QIcon:
    """Small framed swatch used for colour-specific actions."""
    pix = QPixmap(18, 18)
    pix.fill(Qt.transparent)
    painter = QPainter(pix)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setPen(QColor('#59606B'))
    painter.setBrush(QColor(colour))
    painter.drawRoundedRect(2, 2, 13, 13, 3, 3)
    painter.end()
    return QIcon(pix)


def _set_colour_button_icon(button: QPushButton, colour: str = '#6B94B3') -> None:
    button.setIcon(_colour_action_icon(colour))
    button.setIconSize(QSize(16, 16))


class OutlinedCheckBox(QCheckBox):
    """Native checkbox with a theme-aware outline that stays visible on pale themes."""
    def paintEvent(self, event):
        super().paintEvent(event)
        option = QStyleOptionButton()
        self.initStyleOption(option)
        rect = self.style().subElementRect(QStyle.SubElement.SE_CheckBoxIndicator, option, self)
        if not rect.isValid():
            return
        outline = self.palette().color(QPalette.ColorRole.Text)
        outline.setAlpha(190 if self.isEnabled() else 110)
        painter = QPainter(self)
        painter.setPen(outline)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRect(rect.adjusted(0, 0, -1, -1))

class PaletteSwatch(QFrame):
    """Drag a saved palette colour onto any unlocked colour assignment."""
    def __init__(self, colour: str, parent=None):
        super().__init__(parent)
        self.colour = colour
        self.setFixedSize(42, 34)
        self.setToolTip(f"{colour} — drag onto a colour section below")
        self.setStyleSheet(f"background: {colour}; border: 2px solid #59606B; border-radius: 5px;")

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.LeftButton:
            drag = QDrag(self)
            data = QMimeData()
            data.setText(self.colour)
            drag.setMimeData(data)
            drag.exec(Qt.CopyAction)
        else:
            super().mouseMoveEvent(event)


class SavedColourSwatch(QFrame):
    """Persistent personal swatch: left-click adds; right-click removes."""
    def __init__(self, colour: str, on_add, on_remove, parent=None):
        super().__init__(parent)
        self.colour = colour
        self._on_add = on_add
        self._on_remove = on_remove
        self.setFixedSize(36, 29)
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip(f"{colour} — click to add to this theme palette; right-click to remove from My Colours")
        self.setStyleSheet(f"background: {colour}; border: 2px solid #59606B; border-radius: 5px;")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._on_add(self.colour)
            event.accept()
            return
        if event.button() == Qt.RightButton:
            self._on_remove(self.colour)
            event.accept()
            return
        super().mousePressEvent(event)


def colour_option_icon(hex_colour: str) -> QIcon:
    """Draw a framed colour chip beside a HEX code in a dropdown.

    The dark outline remains visible even for near-white or pastel colours.
    This is presentation only: the combo's userData is still its original HEX.
    """
    picture = QPixmap(24, 18)
    picture.fill(Qt.transparent)
    painter = QPainter(picture)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setPen(QColor('#535C68'))
    painter.setBrush(QColor(hex_colour))
    painter.drawRoundedRect(2, 2, 19, 13, 2, 2)
    painter.end()
    return QIcon(picture)


class DropColorCombo(QComboBox):
    def __init__(self, on_drop, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self._on_drop = on_drop
        self.setIconSize(QSize(24, 18))
        # Large imported palettes remain usable without a popup taller than
        # the screen. Qt supplies scrolling for the remaining colour choices.
        self.setMaxVisibleItems(16)
        self.setToolTip('Pick a palette colour (shown beside its HEX code), or drag a swatch here.')

    def dragEnterEvent(self, event):
        if event.mimeData().hasText() and normalise_palette([event.mimeData().text()]):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if self.isEnabled() and event.mimeData().hasText():
            col = normalise_palette([event.mimeData().text()])
            if col:
                self._on_drop(col[0])
                event.acceptProposedAction()
                return
        event.ignore()


class PreviewRoleFrame(QFrame):
    """Clickable preview surface with separate fill and border hit targets."""
    def __init__(self, role: str, on_click=None, parent=None, border_role: str | None = None):
        super().__init__(parent)
        self.role = role
        self.border_role = border_role
        self._on_click = on_click
        self.setCursor(Qt.PointingHandCursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self._on_click:
            role = self.role
            if self.border_role:
                pos = event.position()
                edge = 6
                if (pos.x() <= edge or pos.y() <= edge or
                        pos.x() >= self.width() - edge - 1 or pos.y() >= self.height() - edge - 1):
                    role = self.border_role
            self._on_click(role)
            # PreviewRoleFrames are deliberately nested (for example a filter
            # button inside the filter panel, or a thumbnail inside the side
            # panel).  QFrame's default mouse handler ignores the event, which
            # lets Qt propagate the SAME click to the parent PreviewRoleFrame.
            # That produced a correct jump for a split second followed by the
            # parent role immediately winning.  Once this frame handles the
            # click, consume it here so exactly one role is emitted.
            event.accept()
            return
        super().mousePressEvent(event)


class PreviewCanvas(QWidget):
    """Small live Qt facsimile of the actual editable colour-role surfaces."""
    def __init__(self, section: str, on_role_click=None, parent=None):
        super().__init__(parent)
        self.section = section
        self.on_role_click = on_role_click
        self.painted = []
        self.setMinimumWidth(310)
        root = QVBoxLayout(self)
        root.setContentsMargins(5, 5, 5, 5)
        root.setSpacing(5)
        if section == 'Library':
            self._build_library(root)
        elif section == 'Scan':
            self._build_scan(root)
        else:
            self._build_guide(root)
        root.addStretch(1)

    def block(self, parent_layout, role, title=None, height=0, border_role=None, margins=5, fixed_border=None):
        frame = PreviewRoleFrame(role, self.on_role_click, self, border_role=border_role)
        frame.setObjectName('StudioBlock' + str(len(self.painted)))
        if border_role:
            frame.setToolTip(
                f'Click centre: {ROLES.get(role, role)}\n'
                f'Click edge: {ROLES.get(border_role, border_role)}'
            )
        else:
            frame.setToolTip(f'Click to jump to Theme Studio section: {ROLES.get(role, role)}')
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(margins, margins, margins, margins)
        layout.setSpacing(4)
        if height:
            frame.setMinimumHeight(height)
        parent_layout.addWidget(frame)
        self.painted.append((frame, role, border_role, None, fixed_border))
        if title:
            self.write(layout, title, role, bold=True)
        return layout

    def fixed_block(self, parent_layout, title, height=0, fill='#FFFFFF', border='#BEBEBE', note=''):
        """Render a real app surface which is intentionally not Theme-Studio controlled."""
        frame = QFrame(self)
        frame.setObjectName('StudioFixed' + str(len(self.painted)))
        frame.setCursor(Qt.ArrowCursor)
        frame.setAttribute(Qt.WA_NoMousePropagation, True)
        frame.setToolTip(note or 'This surface uses a fixed app colour and has no Theme Studio colour section.')
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(0)
        if height:
            frame.setMinimumHeight(height)
        lab = QLabel(title, frame)
        lab.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        lab.setStyleSheet('background: transparent; color: #202020; border: none; font-weight: 600;')
        layout.addWidget(lab)
        frame.setStyleSheet(
            f'QFrame#{frame.objectName()} {{ background-color: {fill}; border: 1px solid {border}; border-radius: 4px; }}'
        )
        parent_layout.addWidget(frame)
        return layout

    def write(self, layout, caption, role, bold=False, muted=False):
        lab = QLabel(caption, self)
        lab.setWordWrap(True)
        # Let clicks land on the containing preview block so nested text cannot
        # steal the fill/border hit target.
        lab.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        if bold:
            font = lab.font()
            font.setBold(True)
            lab.setFont(font)
        layout.addWidget(lab)
        self.painted.append((lab, role, None, 'muted_label' if muted else 'label', None))
        return lab

    def _top(self, root):
        top = self.block(root, 'ui_top_fill', height=42, fixed_border='#B7B7B7')
        nav = QHBoxLayout()
        nav.setSpacing(4)
        self.write(nav, 'CAS CURATOR', 'ui_top_fill', bold=True)
        nav.addStretch(1)
        # The toolbar panel shares ui_top_fill with the Search/filter panel,
        # but the actual Library / Scan / Guide buttons are a separate
        # ui_filter_button role.  Keep that distinction visible in the live
        # preview so editing/shuffling the button colour never looks locked to
        # the panel colour.
        for caption in ('Library', 'Scan', 'Guide'):
            self.block(
                nav, 'ui_filter_button', caption, height=26, margins=3,
                fixed_border='#BEBEBE'
            )
        top.addLayout(nav)

    def _build_library(self, root):
        self._top(root)
        # The MDB area and the path interior share ui_top_fill.  Only the
        # nested path outline uses ui_mdb_path_border, matching the real app:
        # the path has no independent fill colour of its own.
        mdb = self.block(root, 'ui_top_fill', height=38)
        mdb_row = QHBoxLayout()
        mdb_row.setSpacing(5)
        self.write(mdb_row, 'Opened Master Database:', 'ui_top_fill', bold=True)
        path_frame = PreviewRoleFrame(
            'ui_top_fill', self.on_role_click, self, border_role='ui_mdb_path_border'
        )
        path_frame.setObjectName('StudioMDBPath' + str(len(self.painted)))
        path_frame.setToolTip(
            'Click centre: Main toolbar / MDB panel background\n'
            'Click edge: Master Database path border'
        )
        path_layout = QHBoxLayout(path_frame)
        path_layout.setContentsMargins(6, 3, 6, 3)
        path_layout.setSpacing(0)
        self.write(path_layout, 'D:\\Sims3\\Library.xlsx', 'ui_top_fill')
        path_frame.setMinimumHeight(27)
        mdb_row.addWidget(path_frame, 1)
        mdb.addLayout(mdb_row)
        self.painted.append((path_frame, 'ui_top_fill', 'ui_mdb_path_border', None, None))
        # Toolbar and Search/filter areas intentionally share ui_top_fill.
        # The actual buttons remain a separate ui_filter_button role.
        filters = self.block(root, 'ui_top_fill', height=88, border_role='ui_scrollbar_color')
        self.write(filters, 'Packages: 10,369    Thumbnails: 9,191', 'ui_top_fill', bold=True, muted=True)
        filterrow = QHBoxLayout()
        # Search/input backgrounds are fixed white in app.py; do not pretend
        # ui_header_fill controls them. This was the main v0.4.71 mis-map.
        self.fixed_block(
            filterrow, 'Search packages…', height=29,
            note='Search field fill is fixed white in CAS Curator. Click the surrounding filter panel or coloured filter button to edit those roles.'
        )
        self.block(filterrow, 'ui_filter_button', 'Category ▾', height=29, fixed_border='#BEBEBE')
        filters.addLayout(filterrow)
        # ui_header_fill is used by real header/pinned-tab surfaces, not by the
        # white Search field. Give it its own honest preview target.
        # Keep the Library header strip as an explicit clickable preview
        # target.  Clicking this strip jumps straight to the Library header
        # strip colour role rather than leaving the user to find the row.
        self.block(
            filters, 'ui_header_fill', 'Library header strip / pinned tab',
            height=24, fixed_border='#B7B7B7'
        )

        middle = QHBoxLayout()
        middle.setSpacing(5)
        table = self.block(middle, 'ui_main_fill', height=228, fixed_border='#D6D6D6')
        self.block(table, 'ui_table_header_fill', 'Package                 Type         Status', height=28, fixed_border='#B7B7B7')
        self.block(table, 'ui_main_fill', '001   Cloudy Hair        Hair         Safe', height=33, fixed_border='#D6D6D6')
        self.block(table, 'ui_row_alt', '002   Rose Dress          Outfit      Safe', height=33, fixed_border='#D6D6D6')
        self.block(table, 'ui_main_fill', '003   Berry Liner        Makeup    Review', height=33, fixed_border='#D6D6D6')
        self.block(table, 'ui_row_alt', '004   Cottage Shoes    Shoes       Safe', height=33, fixed_border='#D6D6D6')

        side = self.block(middle, 'ui_side_fill', height=228, fixed_border='#CFCFCF')
        self.write(side, 'Selected CC', 'ui_side_fill', bold=True)
        self.block(side, 'ui_thumb_fill', 'PREVIEW', height=56, border_role='ui_info_card_border')
        self.block(side, 'ui_tab_fill', 'Inactive detail tab', height=28, fixed_border='#C7C7C7')
        self.block(side, 'ui_tab_selected_fill', 'Selected detail tab', height=28, fixed_border='#C7C7C7')
        self.block(side, 'ui_info_card_fill', 'Hair collection\nProfile: Testing', height=52, border_role='ui_info_card_border')
        root.addLayout(middle)

        scrollbar = QHBoxLayout()
        scrollbar.addStretch(1)
        self.block(scrollbar, 'ui_scrollbar_color', 'scrollbar', height=10, fixed_border='transparent', margins=1)
        root.addLayout(scrollbar)

    def _build_scan(self, root):
        self._top(root)
        page = self.block(root, 'ui_scan_page_fill', 'CAS Health Scanner', height=288, fixed_border='transparent')
        # For custom themes ui_info_card_border is the shared Scan structural
        # border role; centre clicks select the relevant fill, edge clicks select
        # Structural borders.
        card = self.block(page, 'ui_scan_card_fill', '1. Scan setup', height=165, border_role='ui_info_card_border')
        for label in ('Source folder:', 'Master database:'):
            joined = self.block(card, 'ui_scan_row_fill', height=45, border_role='ui_info_card_border')
            self.write(joined, label, 'ui_scan_row_fill', muted=True)
            row = QHBoxLayout()
            self.block(row, 'ui_scan_input_fill', 'C:\\Sims3\\Packages\\…', height=22, fixed_border='#AFAFAF')
            self.block(row, 'ui_scan_button_fill', 'Browse', height=22, fixed_border='#AFAFAF')
            joined.addLayout(row)
        self.block(page, 'ui_scan_button_fill', 'Preview scan', height=31, fixed_border='#AFAFAF')
        self.block(page, 'ui_scan_log_fill', 'Scanner output\nPreparing scan…\nNo package files have been changed.', height=83, fixed_border='#AFAFAF')
        self.block(page, 'ui_scan_progress_fill', '  0 / 100 scanned', height=24, fixed_border='#AFAFAF')

    def _build_guide(self, root):
        self._top(root)
        page = self.block(root, 'ui_guide_page_fill', 'CAS Curator Guide', height=315, fixed_border='transparent')
        nav = self.block(page, 'ui_guide_card_fill', 'Quick navigation', height=66, border_role='ui_guide_border')
        self.write(nav, 'Library  •  Scan  •  Routing  •  Themes', 'ui_guide_card_fill', muted=True)
        browser = self.block(page, 'ui_guide_browser_fill', 'Getting started', height=205, border_role='ui_guide_border')
        self.block(browser, 'ui_guide_section_fill', 'What CAS Curator Can Do', height=30, border_role='ui_guide_border')
        self.block(browser, 'ui_guide_highlight_1', 'Where to find it: Library / Scan / Tools', height=28, border_role='ui_guide_border')
        self.block(browser, 'ui_guide_highlight_2', 'Tip / explanation highlight', height=28, border_role='ui_guide_border')
        self.block(browser, 'ui_guide_highlight_3', 'Important / caution highlight', height=28, border_role='ui_guide_border')

    def apply_colours(self, roles):
        for widget, role, border_role, typ, fixed_border in self.painted:
            colour = roles[role]
            if typ in {'label', 'muted_label'}:
                ink = automatic_text_for(colour, muted=(typ == 'muted_label'))
                widget.setStyleSheet(f'background: transparent; color: {ink}; border: none;')
            else:
                if border_role:
                    border = roles[border_role]
                elif fixed_border is not None:
                    border = fixed_border
                else:
                    border = mix(colour, '#777777', .28)
                widget.setStyleSheet(
                    f'QFrame#{widget.objectName()} {{ background-color: {colour}; '
                    f'border: 1px solid {border}; border-radius: 4px; }}'
                )



class ThemeStudio(QDialog):
    def __init__(self, main_window, initial_name='', parent=None):
        super().__init__(parent or main_window)
        self.main = main_window
        self.setObjectName('ThemeStudio')
        self.setWindowTitle('Theme Colour Studio')
        # Keep the modal usable on laptops and Windows high-DPI displays.
        # The editor body can scroll; the save/footer bar cannot.
        self.setMinimumSize(780, 460)
        self._updating = False
        self.undo_stack = []
        self._session_baseline = None
        self.assignments = {}
        self.locks = {}  # Section Locks: exact role/slot stays fixed.
        self.shuffle_locks = set()  # Shuffle Locks: colours survive but may move.
        self.palette = []
        self._original_palette = []
        # Colours already persisted in My Colours can be added to the working
        # palette without making the theme itself dirty. They remain available
        # from My Colours even if the user closes without saving the theme.
        self._auto_saved_palette_colours = set()
        storage_edition = getattr(main_window, 'theme_storage_edition', 'DEV')
        self.saved_colour_file = default_saved_colour_file(storage_edition)
        # First PUBLIC use inherits the existing DEV My Colours shelf once,
        # then the two editions remain independent.
        if str(storage_edition).upper() == 'PUBLIC' and not self.saved_colour_file.exists():
            dev_colours = load_saved_colours(default_saved_colour_file('DEV'))
            if dev_colours:
                save_saved_colours(dev_colours, self.saved_colour_file)
        self.saved_colours = load_saved_colours(self.saved_colour_file)
        self._my_colours_dialog = None
        self.source_name = ''
        self.base_theme = {}
        self.base_payload = {}
        self.role_rows = {}
        self._highlighted_role = None
        self._highlight_generation = 0
        self.colour_library_rows = True
        self._original_colour_library_rows = True

        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 9, 10, 9)
        outer.setSpacing(7)
        caption = QLabel('Import a palette or create your own swatches using the colour wheel, then assign colours to sections. '
                         'Preview changes before saving. Built-in themes are never editable.', self)
        caption.setWordWrap(True)
        outer.addWidget(caption)
        first = QHBoxLayout()
        first.addWidget(QLabel('Custom theme:'))
        self.theme_picker = QComboBox(self)
        self.theme_picker.setMinimumWidth(235)
        for name in sorted(self.main.custom_theme_files, key=str.casefold):
            self.theme_picker.addItem(name)
        self.theme_picker.currentTextChanged.connect(self._source_changed)
        first.addWidget(self.theme_picker)
        self.manage_themes_button = QPushButton('Manage Themes…', self)
        _set_button_icon(self.manage_themes_button, 'SP_FileDialogDetailedView', 'SP_DialogOpenButton')
        self.manage_themes_button.setToolTip('Import themes or palettes, manage custom themes, export a backup, open the theme folder, or reload themes.')
        self.manage_themes_button.clicked.connect(self.open_manage_themes_menu)
        first.addWidget(self.manage_themes_button)
        first.addStretch(1)
        self.source_label = QLabel('')
        first.addWidget(self.source_label)
        outer.addLayout(first)

        row_style_line = QHBoxLayout()
        self.colour_library_rows_check = QCheckBox('Colour Library rows with this custom theme', self)
        self.colour_library_rows_check.setToolTip(
            'On: Library rows use this custom theme’s Main table background and Alternating table rows colours. '
            'Off: Library rows use the normal neutral white/grey treatment used by built-in themes.'
        )
        self.colour_library_rows_check.toggled.connect(self._colour_library_rows_changed)
        row_style_line.addWidget(self.colour_library_rows_check)
        row_style_hint = QLabel('Turn off for default neutral rows.', self)
        row_style_hint.setStyleSheet('color: #5A6871;')
        row_style_line.addWidget(row_style_hint)
        row_style_line.addStretch(1)
        outer.addLayout(row_style_line)

        body_widget = QWidget(self)
        content = QHBoxLayout(body_widget)
        content.setContentsMargins(0, 0, 0, 0)
        content.setSpacing(10)
        left = QFrame(self)
        left.setObjectName('EditorSidebar')
        # Give long colour-role names enough room at common Windows
        # scaling levels.  In particular, the linked Toolbar + search/filter
        # panel role should be readable without truncation.
        left.setFixedWidth(535)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(8, 7, 8, 7)
        left_layout.setSpacing(6)
        # Keep the palette title on its own compact row so it remains fully
        # readable at Windows 125–150% scaling. The three palette actions sit
        # directly beneath it rather than compressing the title horizontally.
        palette_title = QLabel('YOUR PALETTE — drag colours onto any unlocked section', self)
        left_layout.addWidget(palette_title)
        palette_header = QHBoxLayout()
        palette_header.setContentsMargins(0, 0, 0, 0)
        palette_header.setSpacing(6)
        self.my_colours_button = QPushButton('My Colours…', self)
        _set_colour_button_icon(self.my_colours_button, '#8B6FA8')
        self.my_colours_button.setToolTip('Open your persistent reusable colour shelf without changing the colour-role list layout.')
        self.my_colours_button.clicked.connect(self.open_my_colours_dialog)
        palette_header.addWidget(self.my_colours_button, 1)
        self.paste_palette_button = QPushButton('Paste Link / HEX…', self)
        _set_colour_button_icon(self.paste_palette_button, '#4D8F7A')
        self.paste_palette_button.setToolTip('Open the same Paste Link / HEX Colours importer available from the Theme menu. Paste a supported palette website link or HEX colour list.')
        self.paste_palette_button.clicked.connect(self.open_existing_palette_importer)
        palette_header.addWidget(self.paste_palette_button, 1)
        self.add_colour_button = QPushButton('+ Colour Wheel…', self)
        _set_colour_button_icon(self.add_colour_button, '#5F9EC8')
        self.add_colour_button.setToolTip('Create a colour on the wheel or paste a HEX value, then add it to this palette.')
        self.add_colour_button.clicked.connect(self.open_colour_wheel)
        palette_header.addWidget(self.add_colour_button, 1)
        self.my_colours_button.setText(f'My Colours ({len(self.saved_colours)})…')
        left_layout.addLayout(palette_header)
        self.palette_scroll = QScrollArea(self)
        self.palette_scroll.setWidgetResizable(True)
        self.palette_scroll.setFixedHeight(132)
        self.palette_widget = QWidget()
        self.palette_grid = QGridLayout(self.palette_widget)
        self.palette_grid.setContentsMargins(4, 3, 4, 3)
        self.palette_grid.setSpacing(3)
        self.palette_scroll.setWidget(self.palette_widget)
        left_layout.addWidget(self.palette_scroll)
        self.palette_status = QLabel('')
        self.palette_status.setWordWrap(True)
        left_layout.addWidget(self.palette_status)

        lock_help = QLabel('COLOURABLE SECTIONS — Text colours are chosen automatically for readability. Section Lock keeps this row + colour fixed. Shuffle Lock keeps this colour in every Shuffle, but the colour may move to another unlocked row.')
        lock_help.setWordWrap(True)
        left_layout.addWidget(lock_help)
        self.roles_scroll = QScrollArea(self)
        self.roles_scroll.setWidgetResizable(True)
        self.roles_scroll.setMinimumHeight(190)
        scroll_inner = QWidget(self)
        self.role_layout = QVBoxLayout(scroll_inner)
        self.role_layout.setContentsMargins(2, 2, 3, 3)
        self.role_layout.setSpacing(5)
        self.controls = {}
        for category, group in ROLE_GROUPS:
            section = QLabel(category.upper(), scroll_inner)
            section.setStyleSheet('font-weight: bold; color: #2B6683; margin-top: 7px;')
            self.role_layout.addWidget(section)
            for role, label in group:
                role_frame = QFrame(scroll_inner)
                role_frame.setObjectName('ThemeStudioRoleRow')
                role_frame.setStyleSheet('QFrame#ThemeStudioRoleRow { border: 2px solid transparent; border-radius: 6px; background: transparent; }')
                role_block = QVBoxLayout(role_frame)
                role_block.setContentsMargins(4, 3, 4, 3)
                role_block.setSpacing(1)
                row = QHBoxLayout()
                row.setSpacing(3)
                title = QLabel(label, scroll_inner)
                title.setFixedWidth(220)
                title.setToolTip(label)
                row.addWidget(title)
                colour_marker = QLabel(scroll_inner)
                colour_marker.setFixedSize(23, 23)
                colour_marker.setToolTip('Effective colour in your live preview')
                row.addWidget(colour_marker)
                combo = DropColorCombo(lambda colour, r=role: self.set_colour(r, colour), scroll_inner)
                combo.setMinimumWidth(118)
                combo.setSizePolicy(combo.sizePolicy().horizontalPolicy(), combo.sizePolicy().verticalPolicy())
                combo.currentIndexChanged.connect(lambda _idx, r=role: self._combo_changed(r))
                row.addWidget(combo, 1)
                role_block.addLayout(row)

                lock_row = QHBoxLayout()
                lock_row.setSpacing(9)
                # Keep the lock controls aligned beneath the colour control
                # after widening the role-name column above.
                lock_row.addSpacing(246)
                section_lock = OutlinedCheckBox('Section Lock', scroll_inner)
                section_lock.setObjectName('ThemeStudioLockCheck')
                section_lock.setMinimumHeight(26)
                section_lock.setStyleSheet(
                    'QCheckBox#ThemeStudioLockCheck { spacing: 7px; font-weight: 600; } '
                    'QCheckBox#ThemeStudioLockCheck::indicator { width: 18px; height: 18px; } '
                    'QCheckBox#ThemeStudioLockCheck:checked { font-weight: 700; }'
                )
                section_lock.setToolTip('Fix this section to its current colour. Section-locked sections are not edited, swapped, auto-filled or shuffled.')
                section_lock.toggled.connect(lambda yes, r=role: self._section_lock_changed(r, yes))
                lock_row.addWidget(section_lock)
                shuffle_lock = OutlinedCheckBox('Shuffle Lock', scroll_inner)
                shuffle_lock.setObjectName('ThemeStudioLockCheck')
                shuffle_lock.setMinimumHeight(26)
                shuffle_lock.setStyleSheet(
                    'QCheckBox#ThemeStudioLockCheck { spacing: 7px; font-weight: 600; } '
                    'QCheckBox#ThemeStudioLockCheck::indicator { width: 18px; height: 18px; } '
                    'QCheckBox#ThemeStudioLockCheck:checked { font-weight: 700; }'
                )
                shuffle_lock.setToolTip('Protect this colour across every Shuffle Unlocked action. The colour may move to a different unlocked section.')
                shuffle_lock.toggled.connect(lambda yes, r=role: self._shuffle_lock_changed(r, yes))
                lock_row.addWidget(shuffle_lock)
                lock_row.addStretch(1)
                role_block.addLayout(lock_row)
                self.role_layout.addWidget(role_frame)
                self.role_rows[role] = role_frame
                self.controls[role] = (combo, section_lock, shuffle_lock, colour_marker)
        self.role_layout.addStretch()
        self.roles_scroll.setWidget(scroll_inner)
        left_layout.addWidget(self.roles_scroll, 1)

        swap_title = QLabel('SWAP ANY TWO UNLOCKED SECTIONS')
        left_layout.addWidget(swap_title)
        swap_row = QHBoxLayout()
        self.swap_a, self.swap_b = QComboBox(self), QComboBox(self)
        for role, label in ROLES.items():
            for cb in (self.swap_a, self.swap_b):
                cb.addItem(label, role)
        self.swap_b.setCurrentIndex(1)
        swap_row.addWidget(self.swap_a)
        swap_row.addWidget(self.swap_b)
        swap_btn = QPushButton('Swap')
        _set_button_icon(swap_btn, 'SP_BrowserReload', 'SP_ArrowForward')
        swap_btn.clicked.connect(self.swap_two)
        swap_row.addWidget(swap_btn)
        left_layout.addLayout(swap_row)
        # Sticky action strip: these controls must remain visible even when the
        # editor body scrolls on smaller/high-DPI displays. Keep this outside
        # StudioBodyScroll so Shuffle is always reachable.
        self.sticky_actions = QFrame(self)
        self.sticky_actions.setObjectName('StudioStickyActions')
        actions = QHBoxLayout(self.sticky_actions)
        actions.setContentsMargins(8, 6, 8, 6)
        actions.setSpacing(8)
        fill = QPushButton('Auto-fill')
        _set_button_icon(fill, 'SP_DialogApplyButton', 'SP_ArrowForward')
        fill.setToolTip('Spread distinct imported colours across as many unlocked sections as possible.')
        fill.clicked.connect(self.autofill_unlocked)
        actions.addWidget(fill)
        shuffle = QPushButton('Shuffle Unlocked')
        _set_button_icon(shuffle, 'SP_BrowserReload', 'SP_ArrowForward')
        shuffle.setToolTip('Shuffle only the unlocked parts of the theme. Section Locks stay in place and Shuffle Lock colours are guaranteed to remain, but may move.')
        shuffle.clicked.connect(self.shuffle_unlocked)
        actions.addWidget(shuffle)
        self.undo_btn = QPushButton('Undo')
        _set_button_icon(self.undo_btn, 'SP_ArrowBack', 'SP_ArrowLeft')
        self.undo_btn.clicked.connect(self.undo)
        actions.addWidget(self.undo_btn)
        self.revert_btn = QPushButton('Revert Session')
        _set_button_icon(self.revert_btn, 'SP_DialogResetButton', 'SP_BrowserReload')
        self.revert_btn.setToolTip('Restore Theme Colour Studio to exactly how this custom theme looked when you opened it. Nothing is saved until Update or Save.')
        self.revert_btn.clicked.connect(self.revert)
        actions.addWidget(self.revert_btn)

        # Size the sticky action strip from the real button text + icons instead of
        # forcing it into the historical 438px sidebar width. This prevents longer
        # labels such as 'Shuffle Unlocked' and 'Revert Session' from being clipped
        # at 125% display scaling. Keep one row and widen the editor sidebar to match.
        sticky_buttons = (fill, shuffle, self.undo_btn, self.revert_btn)
        for button in sticky_buttons:
            button.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
            button.setMinimumWidth(button.sizeHint().width() + 8)
        actions.activate()
        sticky_width = max(535, actions.sizeHint().width())
        self.sticky_actions.setFixedWidth(sticky_width)
        left.setFixedWidth(sticky_width)

        content.addWidget(left)

        right = QWidget(self)
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(5)
        title = QLabel('LIVE PREVIEW — current saved theme versus your unsaved edits', self)
        right_layout.addWidget(title)
        self.preview_tabs = QTabWidget(self)
        self.canvases = {}
        for tab in ('Library', 'Scan', 'Guide'):
            page = QWidget(self)
            h = QHBoxLayout(page)
            h.setContentsMargins(3, 5, 3, 3)
            for side, heading in (('before', 'CURRENT SAVED'), ('after', 'YOUR EDITS')):
                wrapper = QFrame(page)
                wrapper_layout = QVBoxLayout(wrapper)
                wrapper_layout.setContentsMargins(3, 3, 3, 3)
                label = QLabel(heading)
                label.setStyleSheet('font-weight: bold; color: #334654;')
                wrapper_layout.addWidget(label)
                scroll = QScrollArea(wrapper)
                scroll.setWidgetResizable(True)
                scroll.setMinimumWidth(260)
                canvas = PreviewCanvas(tab, self._preview_role_clicked)
                scroll.setWidget(canvas)
                wrapper_layout.addWidget(scroll, 1)
                self.canvases[(tab, side)] = canvas
                h.addWidget(wrapper, 1)
            self.preview_tabs.addTab(page, tab)
        right_layout.addWidget(self.preview_tabs, 1)
        reminder = QLabel('Preview uses live Qt colour roles (sample content, not real CC). '
                          'Your main application stays unchanged until you save.')
        reminder.setWordWrap(True)
        right_layout.addWidget(reminder)
        readability_row = QHBoxLayout()
        self.readability_status = QLabel('Readability: checking…', self)
        self.readability_status.setWordWrap(True)
        readability_row.addWidget(self.readability_status, 1)
        self.readability_button = QPushButton('Check Readability…', self)
        _set_button_icon(self.readability_button, 'SP_MessageBoxInformation', 'SP_DialogHelpButton')
        self.readability_button.setToolTip('Checks CAS Curator’s automatic text colours against the live theme surfaces. Nothing is changed automatically.')
        self.readability_button.clicked.connect(self.show_readability_report)
        readability_row.addWidget(self.readability_button)
        right_layout.addLayout(readability_row)
        content.addWidget(right, 1)

        # Make the whole editing surface scroll vertically if OS scaling or
        # a small monitor leaves insufficient room. The sticky action strip and
        # save/footer bar below this scroll area remain visible regardless of scrolling.
        self.body_scroll = QScrollArea(self)
        self.body_scroll.setObjectName('StudioBodyScroll')
        self.body_scroll.setFrameShape(QFrame.NoFrame)
        self.body_scroll.setWidgetResizable(True)
        self.body_scroll.setMinimumHeight(220)
        self.body_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.body_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.body_scroll.setWidget(body_widget)
        outer.addWidget(self.body_scroll, 1)

        sticky_action_row = QHBoxLayout()
        sticky_action_row.setContentsMargins(0, 0, 0, 0)
        sticky_action_row.addWidget(self.sticky_actions, 0, Qt.AlignLeft)
        sticky_action_row.addStretch(1)
        outer.addLayout(sticky_action_row)

        footer = QFrame(self)
        footer.setObjectName('StudioFooter')
        bottom = QHBoxLayout(footer)
        bottom.setContentsMargins(8, 7, 8, 7)
        bottom.addWidget(QLabel('Only the chosen custom theme / new variation can be changed. Built-in themes stay protected.'))
        bottom.addStretch(1)
        apply_button = QPushButton('Apply This Theme')
        _set_button_icon(apply_button, 'SP_DialogApplyButton', 'SP_ArrowForward')
        apply_button.setToolTip('Use this custom theme in CAS Curator. If you have unsaved edits, CAS Curator will ask to save them first.')
        apply_button.clicked.connect(self.apply_this_theme)
        bottom.addWidget(apply_button)
        update = QPushButton('Update Theme…')
        _set_button_icon(update, 'SP_DialogApplyButton', 'SP_DialogSaveButton')
        update.clicked.connect(self.update_existing)
        bottom.addWidget(update)
        save = QPushButton('Save as New Theme…')
        _set_button_icon(save, 'SP_DialogSaveButton', 'SP_FileIcon')
        save.clicked.connect(self.save_as_new)
        bottom.addWidget(save)
        cancel = QPushButton('Close')
        _set_button_icon(cancel, 'SP_DialogCloseButton', 'SP_DialogCancelButton')
        cancel.clicked.connect(self.reject)
        bottom.addWidget(cancel)
        outer.addWidget(footer, 0)

        self.setStyleSheet('''
            QDialog#ThemeStudio { background: #EDF0F2; color: #202D35; }
            QFrame#EditorSidebar { background: #F6F7F9; border: 1px solid #C5CED4; border-radius: 8px; }
            QFrame#StudioStickyActions { background: #F6F7F9; border: 1px solid #C5CED4; border-radius: 8px; }
            QFrame#StudioFooter { background: #E0E8ED; border-top: 1px solid #ADBCC6; border-radius: 4px; }
            QScrollArea#StudioBodyScroll { background: transparent; border: none; }
            QFrame#EditorSidebar QLabel { color: #24333C; }
            QScrollArea { border: 1px solid #C6CDD2; background: #F6F7F9; }
            QScrollBar:vertical { width: 17px; }
            QScrollBar::handle:vertical { background: #8A9CAB; min-height: 34px; border-radius: 5px; margin: 2px; }
            QComboBox { min-height: 23px; }
        ''')
        # An 810-pixel logical-height dialog can extend beneath the taskbar at
        # 125–150% display scaling. Bound the opening size to the available
        # desktop area while retaining resizability/maximise support.
        screen = self.screen() or QApplication.primaryScreen()
        if screen is not None:
            available = screen.availableGeometry()
            self.resize(min(1380, max(780, available.width() - 36)),
                        min(810, max(460, available.height() - 76)))
        else:
            self.resize(1180, 690)

        chosen = initial_name if initial_name in self.main.custom_theme_files else self.theme_picker.currentText()
        if chosen:
            self._updating = True
            self.theme_picker.setCurrentText(chosen)
            self._updating = False
            self.load_theme(chosen)

    def _preview_role_clicked(self, role):
        """Jump from a clicked preview surface to its exact editable role.

        Do not move keyboard focus to the target combo. With nested QScrollAreas,
        Qt's automatic focus visibility handling can override our scroll position
        a moment later and snap back to the previously focused editor row.
        """
        row = self.role_rows.get(role)
        if row is None:
            return
        self._highlight_generation += 1
        generation = self._highlight_generation
        if self._highlighted_role and self._highlighted_role in self.role_rows:
            self._set_role_row_highlight(self._highlighted_role, False)
        self._highlighted_role = role
        self._set_role_row_highlight(role, True)

        # A colour combo can retain keyboard focus while the user clicks a
        # preview. Clear that stale focus before positioning the role list so
        # Qt cannot immediately auto-scroll back to the old focused row.
        focused = QApplication.focusWidget()
        if focused is not None and self.roles_scroll.isAncestorOf(focused):
            focused.clearFocus()

        # Scroll only after the highlight stylesheet/layout has settled, then
        # repeat once after the mouse event completes. This makes the clicked
        # role, rather than focus machinery, authoritative.
        QTimer.singleShot(0, lambda r=role, g=generation: self._centre_preview_role(r, g))
        QTimer.singleShot(80, lambda r=role, g=generation: self._centre_preview_role(r, g))
        QTimer.singleShot(6500, lambda r=role, g=generation: self._clear_preview_role_highlight(r, g))

    def _centre_preview_role(self, role, generation):
        if generation != self._highlight_generation or role != self._highlighted_role:
            return
        row = self.role_rows.get(role)
        inner = self.roles_scroll.widget()
        if row is None or inner is None:
            return
        bar = self.roles_scroll.verticalScrollBar()
        row_y = row.mapTo(inner, row.rect().topLeft()).y()
        viewport_height = self.roles_scroll.viewport().height()
        target = row_y - max(0, (viewport_height - row.height()) // 2)
        target = max(bar.minimum(), min(target, bar.maximum()))
        bar.setValue(target)

    def _set_role_row_highlight(self, role, active):
        row = self.role_rows.get(role)
        if row is None:
            return
        if active:
            row.setStyleSheet(
                'QFrame#ThemeStudioRoleRow { border: 3px solid #2B6683; '
                'border-radius: 6px; background-color: rgba(43, 102, 131, 28); }'
            )
            row.setToolTip(f'Preview selection: {ROLES.get(role, role)}')
        else:
            row.setStyleSheet(
                'QFrame#ThemeStudioRoleRow { border: 2px solid transparent; '
                'border-radius: 6px; background: transparent; }'
            )
            row.setToolTip('')

    def _clear_preview_role_highlight(self, role, generation):
        # A timer from an older click must never clear a newer selection.
        if generation != self._highlight_generation or role != self._highlighted_role:
            return
        self._set_role_row_highlight(role, False)
        self._highlighted_role = None

    def _refresh_theme_picker_after_manage(self):
        """Refresh custom-theme choices after a shared Manage Themes action."""
        current = self.source_name or self.theme_picker.currentText()
        names = sorted(self.main.custom_theme_files, key=str.casefold)
        if not names:
            QMessageBox.information(self, 'Theme Colour Studio',
                'There are no custom themes left to edit. You can create or import one from Theme > Manage Themes.')
            self.reject()
            return

        target = current if current in self.main.custom_theme_files else (
            self.main.theme_name if self.main.theme_name in self.main.custom_theme_files else names[0]
        )
        self._updating = True
        try:
            self.theme_picker.clear()
            self.theme_picker.addItems(names)
            self.theme_picker.setCurrentText(target)
        finally:
            self._updating = False

        # If the source was renamed/deleted while the manager was open, load the
        # replacement selection. Otherwise keep any unsaved editor work intact.
        if target != self.source_name:
            self.load_theme(target)

    def open_manage_themes_menu(self):
        """Open the same Manage Themes actions that appear in the main Theme menu."""
        self.main.show_manage_themes_menu(self.manage_themes_button, from_studio=True)
        self._refresh_theme_picker_after_manage()

    def open_existing_palette_importer(self):
        """Open MainWindow's existing website/HEX palette importer.

        Theme Studio deliberately reuses the exact same dialog and import path
        as Theme > Manage Themes > Paste Link / HEX Colours…, rather
        than maintaining a second parser or website importer.  Refresh only the
        picker afterwards; do not discard or switch away from unsaved Studio
        edits automatically.
        """
        current = self.source_name or self.theme_picker.currentText()
        self.main.paste_palette()
        self._updating = True
        try:
            self.theme_picker.clear()
            for name in sorted(self.main.custom_theme_files, key=str.casefold):
                self.theme_picker.addItem(name)
            if current in self.main.custom_theme_files:
                self.theme_picker.setCurrentText(current)
            elif self.main.theme_name in self.main.custom_theme_files:
                self.theme_picker.setCurrentText(self.main.theme_name)
        finally:
            self._updating = False

    def _source_changed(self, name):
        if self._updating or not name or name == self.source_name:
            return
        if self.is_dirty():
            choice = QMessageBox.question(self, 'Discard Unsaved Colours?',
                'Switch themes and discard your unsaved edits?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if choice != QMessageBox.Yes:
                self._updating = True
                self.theme_picker.setCurrentText(self.source_name)
                self._updating = False
                return
        self.load_theme(name)

    def load_theme(self, name):
        if name not in self.main.custom_theme_files:
            raise ValueError('Theme Colour Studio can edit custom themes only.')
        self.source_name = name
        self.base_payload = copy.deepcopy(self.main.custom_theme_info[name])
        self.base_theme = copy.deepcopy(self.base_payload['theme'])
        self.palette = normalise_palette(self.base_payload.get('palette') or self.base_theme.get('_palette', []))
        if len(self.palette) < 2:
            self.palette = normalise_palette([self.base_theme.get('accent', '#537988'),
                                              self.base_theme.get('accent_soft', '#CBDDD9')])
        self._original_palette = list(self.palette)
        self._auto_saved_palette_colours.clear()
        self.assignments = {role: None for role in ROLES}
        # studio_locks is the legacy v3 field. Treat it as Section Lock so older
        # themes keep exactly the behaviour they had before this split.
        persisted_locks = self.base_payload.get('studio_section_locks') or self.base_payload.get('studio_locks') or {}
        self.locks = {role: bool(persisted_locks.get(role)) for role in ROLES}
        self.shuffle_locks = set(normalise_palette(self.base_payload.get('studio_shuffle_locks') or []))
        # Older custom themes pre-date this option and keep their historical
        # coloured-row appearance. Users can switch any theme to neutral rows.
        self.colour_library_rows = bool(self.base_theme.get('_colour_library_rows', True))
        self._original_colour_library_rows = self.colour_library_rows
        # A section cannot meaningfully use both lock types at once. Section Lock
        # wins for legacy themes that persisted both, because it is the stronger
        # instruction: keep this exact section and colour fixed.
        current_for_locks = self._effective()
        for role, is_locked in self.locks.items():
            if is_locked:
                self.shuffle_locks.discard(current_for_locks[role])
        self.undo_stack = []
        # Session baseline is the exact in-memory state at open/switch time.
        # Revert restores this snapshot rather than trying to reconstruct it later.
        self._session_baseline = self._capture_state()
        self.source_label.setText(f'{len(self.palette)} available colours • no files altered')
        self._rebuild_swatches()
        self._sync_editors()
        self._refresh_preview()

    def _rebuild_swatches(self):
        while self.palette_grid.count():
            item = self.palette_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for index, colour in enumerate(self.palette):
            self.palette_grid.addWidget(PaletteSwatch(colour, self.palette_widget), index // 8, index % 8)
        self.palette_status.setText(
            f'All {len(self.palette)} colours shown ({MAX_IMPORT_COLOURS} maximum). '
            'Changes are only saved when you choose Save or Update.')
        # The wheel remains useful even when this theme palette is full because
        # a colour can still be saved into the cross-theme My Colours shelf.
        self.add_colour_button.setEnabled(True)
        self.source_label.setText(f'{len(self.palette)} available colours • '+
                                  ('unsaved edits' if self.is_dirty() else 'no files altered'))

    def _rebuild_saved_colour_swatches(self):
        # My Colours deliberately lives in its own dialog. Keeping its QScrollArea
        # out of the editor sidebar preserves the accepted v0.4.75 colour-role
        # scroll geometry, especially at 125% Windows display scaling.
        self.my_colours_button.setText(f'My Colours ({len(self.saved_colours)})…')
        dialog = self._my_colours_dialog
        # Populate the shelf even before dialog.exec() makes the dialog visible.
        # v0.4.78 incorrectly gated this on isVisible(), so saved colours appeared
        # immediately after adding them but vanished from the reopened window.
        if dialog is None:
            return
        grid = dialog.saved_colour_grid
        while grid.count():
            item = grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        if not self.saved_colours:
            empty = QLabel('No saved colours yet. Click + Add Colour… below to save your first reusable colour.', dialog.saved_colour_widget)
            empty.setWordWrap(True)
            empty.setAlignment(Qt.AlignCenter)
            grid.addWidget(empty, 0, 0, 1, 8)
        else:
            for index, colour in enumerate(self.saved_colours):
                grid.addWidget(
                    SavedColourSwatch(colour, self._add_my_colour_to_palette,
                                      self._remove_my_colour, dialog.saved_colour_widget),
                    index // 9, index % 9)
        dialog.saved_colour_count.setText(f'{len(self.saved_colours)}/{MAX_SAVED_COLOURS}')

    def open_my_colours_dialog(self):
        # My Colours is global across themes. Reload from the persistent store
        # each time this window opens so a theme switch can never leave a stale
        # or apparently empty in-memory shelf.
        self.saved_colours = load_saved_colours(self.saved_colour_file)
        dialog = QDialog(self)
        dialog.setWindowTitle('My Colours — Theme Colour Studio')
        dialog.setModal(True)
        dialog.resize(450, 270)
        layout = QVBoxLayout(dialog)
        note = QLabel(
            'Reusable colours saved from the Colour Wheel. Click a swatch to add it to the current theme palette; '
            'right-click a swatch to remove it from My Colours. Existing themes are never changed by removing a saved colour.',
            dialog)
        note.setWordWrap(True)
        layout.addWidget(note)
        autosave_note = QLabel('✓ My Colours saves automatically — Add, Remove and Clear are saved immediately. No Save button is needed.', dialog)
        autosave_note.setWordWrap(True)
        layout.addWidget(autosave_note)
        header = QHBoxLayout()
        header.addWidget(QLabel('MY COLOURS', dialog))
        header.addStretch(1)
        dialog.saved_colour_count = QLabel('', dialog)
        header.addWidget(dialog.saved_colour_count)
        layout.addLayout(header)
        scroll = QScrollArea(dialog)
        scroll.setWidgetResizable(True)
        dialog.saved_colour_widget = QWidget(scroll)
        dialog.saved_colour_grid = QGridLayout(dialog.saved_colour_widget)
        dialog.saved_colour_grid.setContentsMargins(6, 6, 6, 6)
        dialog.saved_colour_grid.setSpacing(5)
        scroll.setWidget(dialog.saved_colour_widget)
        layout.addWidget(scroll, 1)
        buttons = QHBoxLayout()
        add_button = QPushButton('+ Add Colour…', dialog)
        _set_colour_button_icon(add_button, '#5F9EC8')
        add_button.setToolTip('Choose a reusable colour. It is saved to My Colours immediately.')
        add_button.clicked.connect(self._add_colour_to_my_colours_only)
        buttons.addWidget(add_button)
        remove_button = QPushButton('Remove Colour…', dialog)
        _set_button_icon(remove_button, 'SP_TrashIcon', 'SP_DialogDiscardButton')
        remove_button.setToolTip('Choose one saved colour to remove from My Colours. Existing themes are not changed.')
        remove_button.clicked.connect(self._remove_saved_colour_from_list)
        buttons.addWidget(remove_button)
        clear_button = QPushButton('Clear All…', dialog)
        _set_button_icon(clear_button, 'SP_TrashIcon', 'SP_DialogDiscardButton')
        clear_button.setToolTip('Remove every saved colour from My Colours. Existing themes are not changed.')
        clear_button.clicked.connect(self._clear_all_saved_colours)
        buttons.addWidget(clear_button)
        buttons.addStretch(1)
        close_button = QPushButton('Close', dialog)
        _set_button_icon(close_button, 'SP_DialogCloseButton', 'SP_DialogCancelButton')
        close_button.clicked.connect(dialog.accept)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)
        self._my_colours_dialog = dialog
        self._rebuild_saved_colour_swatches()
        dialog.exec()
        self._my_colours_dialog = None


    def _add_colour_to_my_colours_only(self):
        default_colour = self.saved_colours[-1] if self.saved_colours else (self.palette[-1] if self.palette else '#ABC4AB')
        picker = ColourWheelDialog(default_colour, self._my_colours_dialog or self)
        if picker.exec() != QDialog.Accepted:
            return
        selected = picker.selected_hex()
        try:
            updated_saved, added = add_saved_colour(self.saved_colours, selected)
        except ValueError as exc:
            QMessageBox.warning(self, 'My Colours', str(exc))
            return
        if not added:
            QMessageBox.information(self, 'My Colours', f'{selected} is already saved in My Colours.')
            return
        old_saved = list(self.saved_colours)
        self.saved_colours = updated_saved
        if not self._persist_saved_colours():
            self.saved_colours = old_saved
            self._rebuild_saved_colour_swatches()

    def _persist_saved_colours(self):
        try:
            self.saved_colours = save_saved_colours(self.saved_colours, self.saved_colour_file)
        except OSError as exc:
            QMessageBox.warning(self, 'My Colours', f'Could not save My Colours.\n\n{exc}')
            return False
        self._rebuild_saved_colour_swatches()
        return True

    def _remove_saved_colour_from_list(self):
        if not self.saved_colours:
            QMessageBox.information(self, 'My Colours', 'There are no saved colours to remove.')
            return

        # Removal deliberately uses its own modal selector. The normal My Colours
        # swatches are left-click-to-add palette controls, so reusing them here
        # makes "select for deletion" ambiguous and can accidentally add colours
        # to the working palette. These tiles ONLY select a saved colour.
        chooser = QDialog(self._my_colours_dialog or self)
        chooser.setWindowTitle('Remove Colour from My Colours')
        chooser.setModal(True)
        chooser.resize(470, 340)
        chooser_layout = QVBoxLayout(chooser)

        instructions = QLabel(
            'Select the saved colour you want to remove. Selecting a colour here does not add it to the theme palette.',
            chooser)
        instructions.setWordWrap(True)
        chooser_layout.addWidget(instructions)

        selected_label = QLabel('Selected: none', chooser)
        selected_label.setStyleSheet('font-weight: 600;')
        chooser_layout.addWidget(selected_label)

        scroll = QScrollArea(chooser)
        scroll.setWidgetResizable(True)
        tile_widget = QWidget(scroll)
        tile_grid = QGridLayout(tile_widget)
        tile_grid.setContentsMargins(8, 8, 8, 8)
        tile_grid.setSpacing(8)
        scroll.setWidget(tile_widget)
        chooser_layout.addWidget(scroll, 1)

        selected = {'colour': None}
        tile_buttons = []

        action_row = QHBoxLayout()
        action_row.addStretch(1)
        cancel_button = QPushButton('Cancel', chooser)
        _set_button_icon(cancel_button, 'SP_DialogCancelButton', 'SP_DialogCloseButton')
        remove_selected_button = QPushButton('Remove Selected', chooser)
        _set_button_icon(remove_selected_button, 'SP_TrashIcon', 'SP_DialogDiscardButton')
        remove_selected_button.setEnabled(False)
        action_row.addWidget(cancel_button)
        action_row.addWidget(remove_selected_button)
        chooser_layout.addLayout(action_row)

        def refresh_selection(colour):
            selected['colour'] = colour
            selected_label.setText(f'Selected: {colour.upper()}')
            remove_selected_button.setEnabled(True)
            for button, button_colour in tile_buttons:
                is_selected = button_colour == colour
                button.setChecked(is_selected)
                marker = '✓ ' if is_selected else ''
                border = '4px solid #1F6FEB' if is_selected else '2px solid #59606B'
                button.setText(f'{marker}{button_colour.upper()}')
                button.setStyleSheet(
                    f'QPushButton {{ background: {button_colour}; color: {text_for(button_colour)}; '
                    f'border: {border}; border-radius: 6px; font-weight: 600; padding: 6px; }}')

        for index, saved_colour in enumerate(self.saved_colours):
            tile = QPushButton(saved_colour.upper(), tile_widget)
            tile.setCheckable(True)
            tile.setAutoExclusive(False)
            tile.setMinimumSize(125, 46)
            tile.setCursor(Qt.PointingHandCursor)
            tile.setToolTip(f'Select {saved_colour.upper()} for removal')
            tile.setStyleSheet(
                f'QPushButton {{ background: {saved_colour}; color: {text_for(saved_colour)}; '
                'border: 2px solid #59606B; border-radius: 6px; font-weight: 600; padding: 6px; }}')
            tile.clicked.connect(lambda checked=False, colour=saved_colour: refresh_selection(colour))
            tile_buttons.append((tile, saved_colour))
            tile_grid.addWidget(tile, index // 3, index % 3)

        cancel_button.clicked.connect(chooser.reject)
        remove_selected_button.clicked.connect(chooser.accept)

        if chooser.exec() != QDialog.Accepted:
            return
        colour = selected['colour']
        if not colour:
            return
        self._remove_my_colour(colour)

    def _clear_all_saved_colours(self):
        if not self.saved_colours:
            QMessageBox.information(self, 'My Colours', 'There are no saved colours to clear.')
            return
        answer = QMessageBox.question(
            self._my_colours_dialog or self, 'Clear all My Colours?',
            f'Remove all {len(self.saved_colours)} saved colours from My Colours?\n\n'
            'Existing themes that already use these colours will not be changed.',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            return
        old = list(self.saved_colours)
        self.saved_colours = []
        if not self._persist_saved_colours():
            self.saved_colours = old
            self._rebuild_saved_colour_swatches()

    def _add_my_colour_to_palette(self, colour):
        if colour in self.palette:
            self.palette_status.setText(f'{colour} is already in this theme palette.')
            return
        if len(self.palette) >= MAX_IMPORT_COLOURS:
            QMessageBox.information(self, 'Palette Full',
                f'This theme palette already has {MAX_IMPORT_COLOURS} unique colours. My Colours is unchanged.')
            return
        try:
            updated, added = add_custom_colour(self.palette, colour)
        except ValueError as exc:
            QMessageBox.warning(self, 'Add Colour', str(exc))
            return
        if not added:
            return
        self.snapshot()
        self.palette = updated
        self._auto_saved_palette_colours.add(colour)
        self._rebuild_swatches()
        self._sync_editors()
        self._refresh_preview()

    def _remove_my_colour(self, colour):
        # Show the actual saved colour, not only its HEX code, so removal is easy to verify visually.
        preview = QPixmap(52, 52)
        preview.fill(Qt.transparent)
        painter = QPainter(preview)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(QColor('#59606B'))
        painter.setBrush(QColor(colour))
        painter.drawRoundedRect(3, 3, 45, 45, 6, 6)
        painter.end()

        confirm = QMessageBox(self._my_colours_dialog or self)
        confirm.setWindowTitle('Remove from My Colours?')
        confirm.setIconPixmap(preview)
        confirm.setText('Remove this saved colour from My Colours?')
        confirm.setInformativeText(
            f'Colour: {colour}\n\n'
            'Themes that already contain this colour will not be changed.')
        confirm.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        confirm.setDefaultButton(QMessageBox.No)
        answer = confirm.exec()
        if answer != QMessageBox.Yes:
            return
        updated, removed = remove_saved_colour(self.saved_colours, colour)
        if not removed:
            return
        old = list(self.saved_colours)
        self.saved_colours = updated
        if not self._persist_saved_colours():
            self.saved_colours = old
            self._rebuild_saved_colour_swatches()

    def open_colour_wheel(self):
        default_colour = self.palette[-1] if self.palette else (self.saved_colours[-1] if self.saved_colours else '#ABC4AB')
        picker = ColourWheelDialog(default_colour, self)
        if picker.exec() != QDialog.Accepted:
            return
        selected = picker.selected_hex()
        if picker.save_in_my_colours():
            try:
                updated_saved, saved_added = add_saved_colour(self.saved_colours, selected)
            except ValueError as exc:
                QMessageBox.warning(self, 'My Colours', str(exc))
                updated_saved, saved_added = self.saved_colours, False
            if saved_added:
                old_saved = list(self.saved_colours)
                self.saved_colours = updated_saved
                if not self._persist_saved_colours():
                    self.saved_colours = old_saved
                    self._rebuild_saved_colour_swatches()
        colour_is_auto_saved = picker.save_in_my_colours() and selected in self.saved_colours
        if len(self.palette) >= MAX_IMPORT_COLOURS:
            QMessageBox.information(self, 'Palette Full',
                f'This theme palette already has {MAX_IMPORT_COLOURS} unique colours. '
                'The colour was still kept in My Colours if that option was selected.')
            return
        try:
            updated, added = add_custom_colour(self.palette, selected)
        except ValueError as exc:
            QMessageBox.warning(self, 'Add Colour', str(exc))
            return
        if not added:
            self.palette_status.setText(f'{selected} is already in this theme palette.')
            return
        self.snapshot()
        self.palette = updated
        if colour_is_auto_saved:
            self._auto_saved_palette_colours.add(selected)
        self._rebuild_swatches()
        self._sync_editors()
        self._refresh_preview()

    def _capture_state(self):
        return (dict(self.assignments), dict(self.locks), set(self.shuffle_locks),
                list(self.palette), bool(self.colour_library_rows))

    def _restore_state(self, state):
        assignments, locks, shuffle_locks, palette, colour_library_rows = state
        self.assignments = dict(assignments)
        self.locks = dict(locks)
        self.shuffle_locks = set(shuffle_locks)
        self.palette = list(palette)
        self.colour_library_rows = bool(colour_library_rows)
        self._auto_saved_palette_colours.intersection_update(self.palette)

    def _sync_history_controls(self):
        self.undo_btn.setEnabled(bool(self.undo_stack))
        self.revert_btn.setEnabled(self.is_dirty())
        self.source_label.setText(
            f'{len(self.palette)} available colours • ' +
            ('unsaved edits' if self.is_dirty() else 'no files altered')
        )

    def snapshot(self):
        self.undo_stack.append(self._capture_state())
        self.undo_stack = self.undo_stack[-40:]
        self._sync_history_controls()

    def _initial_section_locks(self):
        return self.base_payload.get('studio_section_locks') or self.base_payload.get('studio_locks') or {}

    def _initial_shuffle_locks(self):
        return set(normalise_palette(self.base_payload.get('studio_shuffle_locks') or []))

    def is_dirty(self):
        initial_locks = self._initial_section_locks()
        original_colours = set(self._original_palette)
        auto_saved_extras = self._auto_saved_palette_colours - original_colours
        palette_for_dirty_check = [
            colour for colour in self.palette if colour not in auto_saved_extras
        ]
        return (palette_for_dirty_check != self._original_palette or
                any(value is not None for value in self.assignments.values()) or
                any(self.locks[r] != bool(initial_locks.get(r)) for r in ROLES) or
                self.shuffle_locks != self._initial_shuffle_locks() or
                self.colour_library_rows != self._original_colour_library_rows)

    def _effective(self):
        return resolve_preview_roles(edited_theme(self.base_theme, self.assignments))

    def _shuffle_lock_capacity_ok(self, section_locks=None, protected=None):
        section_locks = dict(self.locks if section_locks is None else section_locks)
        protected = set(self.shuffle_locks if protected is None else protected)
        current = self._effective()
        fixed_colours = {current[role] for role in ROLES if section_locks.get(role)}
        movable = [colour for colour in protected if colour not in fixed_colours]
        available = sum(1 for role in ROLES if not section_locks.get(role))
        return len(movable) <= available

    def _sync_editors(self):
        self._updating = True
        current = self._effective()
        # 'Keep current' means the saved colour, not the possibly unsaved edit.
        saved = resolve_preview_roles(self.base_theme)
        try:
            self.colour_library_rows_check.setChecked(bool(self.colour_library_rows))
            for role, (combo, section_lock, shuffle_lock, marker) in self.controls.items():
                combo.clear()
                base_colour = saved[role]
                combo.addItem(colour_option_icon(base_colour), 'Keep current', None)
                combo.setItemData(0, f'Keep saved colour: {base_colour}', Qt.ToolTipRole)
                for col in self.palette:
                    combo.addItem(colour_option_icon(col), col, col)
                chosen = self.assignments[role]
                if chosen is not None:
                    idx = combo.findData(chosen)
                    if idx < 0:
                        combo.addItem(colour_option_icon(chosen), f'Swapped {chosen}', chosen)
                        idx = combo.findData(chosen)
                    combo.setCurrentIndex(idx)
                else:
                    combo.setCurrentIndex(0)
                section_lock.setChecked(self.locks[role])
                shuffle_lock.setChecked((not self.locks[role]) and current[role] in self.shuffle_locks)
                combo.setEnabled(not self.locks[role])
                marker.setStyleSheet(f'background: {current[role]}; border: 1px solid #777; border-radius: 4px;')
                marker.setToolTip(f'{ROLES[role]}: {current[role]}')
        finally:
            self._updating = False
        self._sync_history_controls()

    def _refresh_preview(self):
        before = resolve_preview_roles(self.base_theme)
        before['_adaptive_text'] = bool(self.base_theme.get('_adaptive_text'))
        if not bool(self.base_theme.get('_colour_library_rows', True)):
            before['ui_main_fill'] = '#FAFAFA'
            before['ui_row_alt'] = '#EEEEEE'
        after_theme = edited_theme(self.base_theme, self.assignments)
        after_theme['_colour_library_rows'] = bool(self.colour_library_rows)
        after = resolve_preview_roles(after_theme)
        after['_adaptive_text'] = True
        if not self.colour_library_rows:
            after['ui_main_fill'] = '#FAFAFA'
            after['ui_row_alt'] = '#EEEEEE'
        for tab in ('Library', 'Scan', 'Guide'):
            self.canvases[(tab, 'before')].apply_colours(before)
            self.canvases[(tab, 'after')].apply_colours(after)
        for role, (_combo, _section_lock, _shuffle_lock, marker) in self.controls.items():
            marker.setStyleSheet(f'background: {after[role]}; border: 1px solid #777; border-radius: 4px;')
            marker.setToolTip(f'{ROLES[role]}: {after[role]}')
        self._update_readability_status()

    def _current_readability_report(self):
        theme = edited_theme(self.base_theme, self.assignments)
        theme['_colour_library_rows'] = bool(self.colour_library_rows)
        if not self.colour_library_rows:
            theme['ui_main_fill'] = '#FAFAFA'
            theme['ui_row_alt'] = '#EEEEEE'
        return readability_report(theme)

    def _update_readability_status(self):
        if not hasattr(self, 'readability_status') or not self.base_theme:
            return
        report = self._current_readability_report()
        poor = sum(item['status'] == 'poor' for item in report)
        borderline = sum(item['status'] == 'borderline' for item in report)
        warnings = poor + borderline
        if not warnings:
            text = f'Readability: ✓ {len(report)} text checks passed'
            style = 'font-weight: 650; color: #27613A;'
        else:
            parts = []
            if poor:
                parts.append(f'{poor} poor')
            if borderline:
                parts.append(f'{borderline} borderline')
            text = 'Readability: ⚠ ' + ', '.join(parts)
            style = 'font-weight: 650; color: #8A4B14;'
        self.readability_status.setText(text)
        self.readability_status.setStyleSheet(style)

    def show_readability_report(self):
        report = self._current_readability_report()
        warnings = [item for item in report if item['status'] != 'good']
        poor = sum(item['status'] == 'poor' for item in warnings)
        borderline = sum(item['status'] == 'borderline' for item in warnings)

        dialog = QDialog(self)
        dialog.setWindowTitle('Theme Readability Check')
        dialog.resize(650, 520)
        layout = QVBoxLayout(dialog)
        intro = QLabel(
            'Advisory only — CAS Curator will not block or change your colours. '
            'The check mirrors the automatic text colours CAS Curator uses for this custom theme. '
            'Normal-text target: 4.5:1; 3.0–4.49:1 is borderline; below 3.0:1 is poor.', dialog)
        intro.setWordWrap(True)
        layout.addWidget(intro)

        if not warnings:
            result = QLabel(f'✓ All {len(report)} checked text/background combinations meet 4.5:1.', dialog)
            result.setStyleSheet('font-weight: 700; color: #27613A; padding: 8px;')
            result.setWordWrap(True)
            layout.addWidget(result)
            layout.addStretch(1)
        else:
            summary = QLabel(
                f'⚠ {len(warnings)} warning(s): {poor} poor, {borderline} borderline. '
                'Use Jump to go straight to the Theme Studio section that controls that background.', dialog)
            summary.setWordWrap(True)
            summary.setStyleSheet('font-weight: 650; color: #7A4317; padding: 4px;')
            layout.addWidget(summary)

            scroll = QScrollArea(dialog)
            scroll.setWidgetResizable(True)
            host = QWidget(scroll)
            rows = QVBoxLayout(host)
            rows.setContentsMargins(5, 5, 5, 5)
            rows.setSpacing(6)
            for item in warnings:
                card = QFrame(host)
                card.setObjectName('ReadabilityWarningCard')
                row = QHBoxLayout(card)
                row.setContentsMargins(8, 7, 8, 7)
                severity = 'POOR' if item['status'] == 'poor' else 'BORDERLINE'
                details = QLabel(
                    f"<b>{item['group']} › {item['label']}</b><br>"
                    f"{severity} — {item['ratio']:.2f}:1 &nbsp; "
                    f"{item.get('text_role', 'Automatic text')} {item['foreground']} on {item['background']}", card)
                details.setWordWrap(True)
                row.addWidget(details, 1)
                jump = QPushButton('Jump', card)
                _set_button_icon(jump, 'SP_ArrowForward', 'SP_ArrowRight')
                role = item['role']
                jump.clicked.connect(lambda _checked=False, r=role, d=dialog: (d.accept(), self._preview_role_clicked(r)))
                row.addWidget(jump)
                rows.addWidget(card)
            rows.addStretch(1)
            scroll.setWidget(host)
            layout.addWidget(scroll, 1)

        close = QPushButton('Close', dialog)
        _set_button_icon(close, 'SP_DialogCloseButton', 'SP_DialogCancelButton')
        close.clicked.connect(dialog.accept)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        buttons.addWidget(close)
        layout.addLayout(buttons)
        dialog.setStyleSheet(
            'QFrame#ReadabilityWarningCard { background: #F7F8F9; border: 1px solid #C7CED3; border-radius: 6px; }')
        dialog.exec()

    def _colour_library_rows_changed(self, checked):
        if self._updating or bool(checked) == self.colour_library_rows:
            return
        self.snapshot()
        self.colour_library_rows = bool(checked)
        self._sync_editors()
        self._refresh_preview()

    def _combo_changed(self, role):
        if self._updating or self.locks[role]:
            return
        value = self.controls[role][0].currentData()
        if value == self.assignments[role]:
            return
        self.snapshot()
        self.assignments[role] = value
        # Shuffle Lock belongs to a colour rather than a row. If the user changes
        # this row, any previously protected colour stays protected and its tick
        # follows that colour wherever it is currently used.
        self._sync_editors()
        self._refresh_preview()

    def _section_lock_changed(self, role, locked):
        if self._updating or self.locks[role] == locked:
            return
        current_colour = self._effective()[role]
        proposed = dict(self.locks)
        proposed[role] = locked
        proposed_shuffle = set(self.shuffle_locks)
        if locked:
            # Section Lock is stronger and makes a Shuffle Lock on the same
            # current colour redundant. Choosing it automatically switches
            # this row to Section Lock.
            proposed_shuffle.discard(current_colour)
        if locked and not self._shuffle_lock_capacity_ok(protected=proposed_shuffle, section_locks=proposed):
            QMessageBox.information(self, 'Section Lock',
                'That Section Lock would leave too few movable sections for the current Shuffle Lock colours.')
            self._sync_editors()
            return
        self.snapshot()
        self.locks[role] = locked
        self.shuffle_locks = proposed_shuffle
        self._sync_editors()

    def _shuffle_lock_changed(self, role, locked):
        if self._updating:
            return
        current_colour = self._effective()[role]
        proposed = set(self.shuffle_locks)
        proposed_sections = dict(self.locks)
        if locked:
            proposed.add(current_colour)
            # Choosing Shuffle Lock automatically releases Section Lock on
            # this row, because Shuffle Lock explicitly allows the colour to move.
            proposed_sections[role] = False
        else:
            proposed.discard(current_colour)
        if proposed == self.shuffle_locks and proposed_sections == self.locks:
            return
        if not self._shuffle_lock_capacity_ok(protected=proposed, section_locks=proposed_sections):
            QMessageBox.information(self, 'Shuffle Lock',
                'There are more Shuffle Lock colours than unlocked sections available. Unlock a section or remove another Shuffle Lock first.')
            self._sync_editors()
            return
        self.snapshot()
        self.locks = proposed_sections
        self.shuffle_locks = proposed
        self._sync_editors()

    def set_colour(self, role, value):
        value = normalise_palette([value])
        if self.locks.get(role) or not value or value[0] not in self.palette:
            return
        if self.assignments[role] == value[0]:
            return
        self.snapshot()
        self.assignments[role] = value[0]
        self._sync_editors()
        self._refresh_preview()

    def swap_two(self):
        one, two = self.swap_a.currentData(), self.swap_b.currentData()
        if not one or not two or one == two:
            QMessageBox.information(self, 'Swap Colours', 'Choose two different sections.')
            return
        if self.locks[one] or self.locks[two]:
            QMessageBox.information(self, 'Section Lock', 'Remove the Section Lock from both sections before swapping.')
            return
        current = self._effective()
        self.snapshot()
        self.assignments[one], self.assignments[two] = current[two], current[one]
        self._sync_editors()
        self._refresh_preview()

    def autofill_unlocked(self):
        if all(self.locks.values()):
            QMessageBox.information(self, 'All Sections Locked', 'Remove a Section Lock to automatically colour a section.')
            return
        self.snapshot()
        # Auto-fill respects Section Locks. Shuffle Locks are deliberately only a
        # Shuffle promise; the protected colour set remains saved for the next Shuffle.
        current = self._effective()
        chosen = [current[role] for role in ROLES if self.locks[role]]
        available = [colour for colour in self.palette if colour not in chosen]
        if not available:
            available = list(self.palette)
        cursor = 0
        for role in ROLES:
            if self.locks[role]:
                continue
            self.assignments[role] = available[cursor % len(available)]
            cursor += 1
        self._sync_editors()
        self._refresh_preview()

    def shuffle_unlocked(self):
        if all(self.locks.values()):
            QMessageBox.information(self, 'All Sections Locked', 'Remove at least one Section Lock to use Shuffle Unlocked.')
            return
        current = self._effective()
        self.snapshot()
        try:
            self.assignments = shuffle_roles(
                self.assignments,
                self.locks,
                self.palette,
                random.Random(),
                protected_colours=self.shuffle_locks,
                current_colours=current,
            )
        except ValueError as exc:
            self.undo_stack.pop()
            QMessageBox.information(self, 'Shuffle Lock', str(exc))
            return
        self._sync_editors()
        self._refresh_preview()

    def undo(self):
        if not self.undo_stack:
            return
        self._restore_state(self.undo_stack.pop())
        self._rebuild_swatches()
        self._sync_editors()
        self._refresh_preview()

    def revert(self):
        if not self.is_dirty() or self._session_baseline is None:
            return
        # Revert itself is undoable: one Undo returns to the exact pre-Revert state.
        self.snapshot()
        self._restore_state(self._session_baseline)
        self._rebuild_swatches()
        self._sync_editors()
        self._refresh_preview()

    def _make_payload(self, name):
        data = copy.deepcopy(self.base_payload)
        data['schema'] = max(int(data.get('schema', 2)), 5)
        data['name'] = name
        data['palette'] = list(self.palette)
        data['theme'] = edited_theme(self.base_theme, self.assignments)
        data['theme']['_palette'] = list(self.palette)
        data['theme']['_colour_library_rows'] = bool(self.colour_library_rows)
        # Keep the legacy key for older DEV builds while writing the explicit v4
        # names used by the split-lock Theme Studio.
        data['studio_locks'] = dict(self.locks)
        data['studio_section_locks'] = dict(self.locks)
        palette_order = {colour: index for index, colour in enumerate(self.palette)}
        data['studio_shuffle_locks'] = sorted(self.shuffle_locks, key=lambda c: (palette_order.get(c, 9999), c))
        edition = str(getattr(self.main, 'theme_storage_edition', 'DEV')).upper()
        data['private_dev_only'] = edition != 'PUBLIC'
        data['edition_origin'] = edition
        # Never call an editor-generated variation the website's original name.
        original = data.get('original_name') or ''
        if name != self.source_name:
            data['name_source'] = 'custom'
        elif original and name.casefold() != str(original).casefold():
            data['name_source'] = 'custom'
        return data

    def _finish_save(self, name):
        self.main._load_custom_themes()
        self.main.apply_theme(name)
        self.main._populate_theme_menu()
        self.accept()

    def _save_existing_theme(self, name):
        old_path = self.main.custom_theme_files.get(name)
        if old_path is None:
            return False
        backup_dir = self.main.custom_theme_dir / 'backups'
        try:
            backup_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            shutil.copy2(old_path, backup_dir / f'{old_path.stem}_{stamp}.s3theme')
            self.main._save_theme_payload(old_path, self._make_payload(name))
        except OSError as exc:
            QMessageBox.warning(self, 'Theme Save Failed', str(exc))
            return False
        self._finish_save(name)
        return True

    def apply_this_theme(self):
        """Apply the selected custom theme without returning to the main Theme menu.

        Unsaved Studio edits must be persisted first so the active theme and the
        saved theme cannot silently disagree after restart.
        """
        name = self.source_name
        if name not in self.main.custom_theme_files:
            return
        if self.is_dirty():
            answer = QMessageBox.question(
                self,
                'Save and Apply Theme?',
                f"'{name}' has unsaved changes. Save those changes and apply the updated theme?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if answer != QMessageBox.Yes:
                return
            self._save_existing_theme(name)
            return
        self.main.apply_theme(name)
        self.main._populate_theme_menu()
        self.accept()

    def update_existing(self):
        name = self.source_name
        if name not in self.main.custom_theme_files:
            return
        if not self.is_dirty():
            QMessageBox.information(self, 'No Unsaved Changes', 'There are no changes to save.')
            return
        answer = QMessageBox.question(self, 'Update Theme?',
            f"Update '{name}'? A backup of the previous custom theme will be saved first.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            return
        self._save_existing_theme(name)

    def save_as_new(self):
        initial = f'{self.source_name} Variation'
        while initial.casefold() in {name.casefold() for name in self.main.custom_theme_files}:
            initial += ' II'
        name, accepted = QInputDialog.getText(self, 'Save as a New Theme',
                                              'New theme name:', text=initial)
        if not accepted:
            return
        name = name.strip()
        if not name or len(name) > 100:
            QMessageBox.warning(self, 'Theme Name', 'Choose a name between 1 and 100 characters.')
            return
        if name.casefold() in {built.casefold() for built in self.main.themes}:
            QMessageBox.warning(self, 'Protected or Duplicate Theme',
                                'A theme already uses that name. Choose another.')
            return
        try:
            path = self.main._next_custom_theme_path(name)
            self.main._save_theme_payload(path, self._make_payload(name))
        except OSError as exc:
            QMessageBox.warning(self, 'Theme Save Failed', str(exc))
            return
        self._finish_save(name)

    def reject(self):
        if self.is_dirty():
            box = QMessageBox(self)
            box.setWindowTitle('Save Theme Changes?')
            box.setIcon(QMessageBox.Question)
            box.setText('Do you want to save your theme changes before exiting Theme Colour Studio?')
            save_button = box.addButton('Save', QMessageBox.ButtonRole.AcceptRole)
            dont_save_button = box.addButton("Don't Save", QMessageBox.ButtonRole.DestructiveRole)
            cancel_button = box.addButton('Cancel', QMessageBox.ButtonRole.RejectRole)
            box.setDefaultButton(save_button)
            box.exec()
            clicked = box.clickedButton()
            if clicked is cancel_button or clicked is None:
                return
            if clicked is save_button:
                if self.source_name in self.main.custom_theme_files:
                    self._save_existing_theme(self.source_name)
                else:
                    self.save_as_new()
                return
            if clicked is not dont_save_button:
                return
        super().reject()
