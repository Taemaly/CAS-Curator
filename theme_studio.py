"""DEV-only Visual Theme Studio. Deliberately imports no scanner/MDB modules.

Edits are transient until the user explicitly chooses Save As or Update.
The rendered mock Library / Scan / Guide pages are live Qt widgets using
exactly the custom style role keys consumed by app.py, not a static image.
"""
from __future__ import annotations

import copy
import random
import shutil
from datetime import datetime

from PySide6.QtCore import Qt, QMimeData, QSize
from PySide6.QtGui import QDrag, QFont, QIcon, QPixmap, QPainter, QColor
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QFrame, QGridLayout,
    QHBoxLayout, QInputDialog, QLabel, QMessageBox, QPushButton,
    QScrollArea, QTabWidget, QVBoxLayout, QWidget,
)

from theme_studio_core import (
    MAX_IMPORT_COLOURS, ROLE_GROUPS, ROLES, add_custom_colour,
    edited_theme, mix, normalise_palette, resolve_preview_roles, shuffle_roles, text_for,
)
from colour_wheel import ColourWheelDialog


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


class PreviewCanvas(QWidget):
    """Small live Qt facsimile of existing role-based Library/Scan/Guide areas."""
    def __init__(self, section: str, parent=None):
        super().__init__(parent)
        self.section = section
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

    def block(self, parent_layout, role, title=None, height=0, border_role=None, margins=5):
        frame = QFrame(self)
        frame.setObjectName('StudioBlock' + str(len(self.painted)))
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(margins, margins, margins, margins)
        layout.setSpacing(4)
        if height:
            frame.setMinimumHeight(height)
        parent_layout.addWidget(frame)
        self.painted.append((frame, role, border_role, None))
        if title:
            self.write(layout, title, role, bold=True)
        return layout

    def write(self, layout, caption, role, bold=False):
        lab = QLabel(caption, self)
        lab.setWordWrap(True)
        if bold:
            font = lab.font()
            font.setBold(True)
            lab.setFont(font)
        layout.addWidget(lab)
        self.painted.append((lab, role, None, 'label'))
        return lab

    def _top(self, root):
        top = self.block(root, 'ui_top_fill', height=42)
        nav = QHBoxLayout()
        nav.setSpacing(4)
        self.write(nav, 'CAS CURATOR', 'ui_top_fill', bold=True)
        nav.addStretch(1)
        self.write(nav, 'Library  •  Scan  •  Guide', 'ui_top_fill')
        top.addLayout(nav)

    def _build_library(self, root):
        self._top(root)
        filters = self.block(root, 'ui_filter_panel_fill', height=56, border_role='ui_info_card_border')
        self.write(filters, 'Packages: 10,369    Thumbnails: 9,191', 'ui_filter_panel_fill', bold=True)
        filterrow = QHBoxLayout()
        inp = self.block(filterrow, 'ui_header_fill', 'Search packages…', height=29)
        inp.setSpacing(0)
        self.block(filterrow, 'ui_filter_button', 'Category ▾', height=29)
        filters.addLayout(filterrow)
        middle = QHBoxLayout()
        middle.setSpacing(5)
        table = self.block(middle, 'ui_main_fill', height=228)
        self.block(table, 'ui_table_header_fill', 'Package                 Type         Status', height=28)
        self.block(table, 'ui_main_fill', '001   Cloudy Hair        Hair         Safe', height=33, border_role='ui_info_card_border')
        self.block(table, 'ui_row_alt', '002   Rose Dress          Outfit      Safe', height=33, border_role='ui_info_card_border')
        self.block(table, 'ui_main_fill', '003   Berry Liner        Makeup    Review', height=33, border_role='ui_info_card_border')
        self.block(table, 'ui_row_alt', '004   Cottage Shoes    Shoes       Safe', height=33, border_role='ui_info_card_border')
        side = self.block(middle, 'ui_side_fill', height=228, border_role='ui_info_card_border')
        self.write(side, 'Selected CC', 'ui_side_fill', bold=True)
        self.block(side, 'ui_thumb_fill', 'PREVIEW', height=56)
        self.block(side, 'ui_tab_fill', 'Overview  •  All fields', height=30)
        self.block(side, 'ui_tab_selected_fill', 'Overview', height=26)
        self.block(side, 'ui_info_card_fill', 'Hair collection\nProfile: Testing', height=52, border_role='ui_info_card_border')
        root.addLayout(middle)
        scrollbar = QHBoxLayout()
        scrollbar.addStretch(1)
        self.block(scrollbar, 'ui_scrollbar_color', height=5)
        root.addLayout(scrollbar)

    def _build_scan(self, root):
        self._top(root)
        page = self.block(root, 'ui_scan_page_fill', 'CAS Health Scanner', height=288)
        card = self.block(page, 'ui_scan_card_fill', '1. Scan setup', height=165, border_role='ui_info_card_border')
        for label in ('Source folder:', 'Master database:'):
            joined = self.block(card, 'ui_scan_row_fill', height=45, border_role='ui_info_card_border')
            self.write(joined, label, 'ui_scan_row_fill')
            row = QHBoxLayout()
            self.block(row, 'ui_scan_input_fill', 'C:\\Sims3\\Packages\\…', height=22)
            self.block(row, 'ui_scan_button_fill', 'Browse', height=22)
            joined.addLayout(row)
        self.block(page, 'ui_scan_button_fill', 'Preview scan', height=31)
        log = self.block(page, 'ui_scan_log_fill', 'Scanner output\nPreparing scan…\nNo package files have been changed.', height=83)
        self.block(page, 'ui_scan_progress_fill', '  0 / 100 scanned', height=24)

    def _build_guide(self, root):
        self._top(root)
        page = self.block(root, 'ui_guide_page_fill', 'CAS Curator Guide', height=270)
        nav = self.block(page, 'ui_guide_card_fill', 'Quick navigation', height=66, border_role='ui_guide_border')
        self.write(nav, 'Library  •  Scan  •  Routing  •  Themes', 'ui_guide_card_fill')
        browser = self.block(page, 'ui_guide_browser_fill', 'Getting started', height=165, border_role='ui_guide_border')
        for s in ('1. Open your Master Database.', '2. Review categorised CAS packages.',
                  '3. Preview routing before moving files.', '4. Save a backup before any changes.'):
            self.write(browser, s, 'ui_guide_browser_fill')

    def apply_colours(self, roles):
        for widget, role, border_role, typ in self.painted:
            colour = roles[role]
            ink = text_for(colour)
            if typ == 'label':
                widget.setStyleSheet(f'background: transparent; color: {ink}; border: none;')
            else:
                border = roles.get(border_role) if border_role else mix(colour, '#777777', .28)
                widget.setStyleSheet(
                    f'QFrame#{widget.objectName()} {{ background-color: {colour}; '
                    f'border: 1px solid {border}; border-radius: 4px; }}'
                )


class ThemeStudio(QDialog):
    def __init__(self, main_window, initial_name='', parent=None):
        super().__init__(parent or main_window)
        self.main = main_window
        self.setObjectName('ThemeStudio')
        self.setWindowTitle('Visual Theme Studio — Private DEV')
        # Keep the modal usable on laptops and Windows high-DPI displays.
        # The editor body can scroll; the save/footer bar cannot.
        self.setMinimumSize(780, 460)
        self._updating = False
        self.undo_stack = []
        self.assignments = {}
        self.locks = {}
        self.palette = []
        self._original_palette = []
        self.source_name = ''
        self.base_theme = {}
        self.base_payload = {}

        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 9, 10, 9)
        outer.setSpacing(7)
        caption = QLabel('Import a palette or create your own swatches using the colour wheel, then assign colours to sections. '
                         'Preview changes before saving. Built-in themes are never editable.', self)
        caption.setWordWrap(True)
        outer.addWidget(caption)
        first = QHBoxLayout()
        first.addWidget(QLabel('Imported theme:'))
        self.theme_picker = QComboBox(self)
        self.theme_picker.setMinimumWidth(235)
        for name in sorted(self.main.custom_theme_files, key=str.casefold):
            self.theme_picker.addItem(name)
        self.theme_picker.currentTextChanged.connect(self._source_changed)
        first.addWidget(self.theme_picker)
        first.addStretch(1)
        self.source_label = QLabel('')
        first.addWidget(self.source_label)
        outer.addLayout(first)

        body_widget = QWidget(self)
        content = QHBoxLayout(body_widget)
        content.setContentsMargins(0, 0, 0, 0)
        content.setSpacing(10)
        left = QFrame(self)
        left.setObjectName('EditorSidebar')
        left.setFixedWidth(438)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(8, 7, 8, 7)
        left_layout.setSpacing(6)
        palette_header = QHBoxLayout()
        palette_header.addWidget(QLabel('YOUR PALETTE — drag colours onto any unlocked section'))
        palette_header.addStretch(1)
        self.add_colour_button = QPushButton('+ Colour Wheel…', self)
        self.add_colour_button.setToolTip('Create a colour on the wheel or paste a HEX value, then add it to this palette.')
        self.add_colour_button.clicked.connect(self.open_colour_wheel)
        palette_header.addWidget(self.add_colour_button)
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
        left_layout.addWidget(QLabel('COLOURABLE SECTIONS — choose, drop or lock a colour'))
        roles_scroll = QScrollArea(self)
        roles_scroll.setWidgetResizable(True)
        roles_scroll.setMinimumHeight(190)
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
                row = QHBoxLayout()
                row.setSpacing(3)
                title = QLabel(label, scroll_inner)
                title.setFixedWidth(157)
                title.setToolTip(label)
                row.addWidget(title)
                colour_marker = QLabel(scroll_inner)
                colour_marker.setFixedSize(23, 23)
                colour_marker.setToolTip('Effective colour in your live preview')
                row.addWidget(colour_marker)
                combo = DropColorCombo(lambda colour, r=role: self.set_colour(r, colour), scroll_inner)
                combo.setMinimumWidth(145)
                combo.setSizePolicy(combo.sizePolicy().horizontalPolicy(), combo.sizePolicy().verticalPolicy())
                combo.currentIndexChanged.connect(lambda _idx, r=role: self._combo_changed(r))
                row.addWidget(combo, 1)
                lock = QCheckBox('Lock', scroll_inner)
                lock.setToolTip('Keep this colour when shuffling; locked sections cannot be swapped or edited.')
                lock.toggled.connect(lambda yes, r=role: self._lock_changed(r, yes))
                row.addWidget(lock)
                self.role_layout.addLayout(row)
                self.controls[role] = (combo, lock, colour_marker)
        self.role_layout.addStretch()
        roles_scroll.setWidget(scroll_inner)
        left_layout.addWidget(roles_scroll, 1)

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
        swap_btn.clicked.connect(self.swap_two)
        swap_row.addWidget(swap_btn)
        left_layout.addLayout(swap_row)
        actions = QHBoxLayout()
        fill = QPushButton('Auto-fill')
        fill.setToolTip('Spread distinct imported colours across as many unlocked sections as possible.')
        fill.clicked.connect(self.autofill_unlocked)
        actions.addWidget(fill)
        shuffle = QPushButton('Shuffle')
        shuffle.setToolTip('Try a fresh colour combination, keeping locked sections unchanged.')
        shuffle.clicked.connect(self.shuffle_unlocked)
        actions.addWidget(shuffle)
        self.undo_btn = QPushButton('Undo')
        self.undo_btn.clicked.connect(self.undo)
        actions.addWidget(self.undo_btn)
        reset = QPushButton('Revert')
        reset.clicked.connect(self.revert)
        actions.addWidget(reset)
        left_layout.addLayout(actions)
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
                canvas = PreviewCanvas(tab)
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
        content.addWidget(right, 1)

        # Make the whole editing surface scroll vertically if OS scaling or
        # a small monitor leaves insufficient room. The action bar BELOW this
        # scroll area remains visible regardless of scrolling.
        self.body_scroll = QScrollArea(self)
        self.body_scroll.setObjectName('StudioBodyScroll')
        self.body_scroll.setFrameShape(QFrame.NoFrame)
        self.body_scroll.setWidgetResizable(True)
        self.body_scroll.setMinimumHeight(220)
        self.body_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.body_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.body_scroll.setWidget(body_widget)
        outer.addWidget(self.body_scroll, 1)

        footer = QFrame(self)
        footer.setObjectName('StudioFooter')
        bottom = QHBoxLayout(footer)
        bottom.setContentsMargins(8, 7, 8, 7)
        bottom.addWidget(QLabel('Only the chosen imported theme / new variation can be changed.'))
        bottom.addStretch(1)
        update = QPushButton('Update Imported Theme…')
        update.clicked.connect(self.update_existing)
        bottom.addWidget(update)
        save = QPushButton('Save as New Theme…')
        save.clicked.connect(self.save_as_new)
        bottom.addWidget(save)
        cancel = QPushButton('Close')
        cancel.clicked.connect(self.reject)
        bottom.addWidget(cancel)
        outer.addWidget(footer, 0)

        self.setStyleSheet('''
            QDialog#ThemeStudio { background: #EDF0F2; color: #202D35; }
            QFrame#EditorSidebar { background: #F6F7F9; border: 1px solid #C5CED4; border-radius: 8px; }
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
            raise ValueError('Visual Theme Studio can edit imported themes only.')
        self.source_name = name
        self.base_payload = copy.deepcopy(self.main.custom_theme_info[name])
        self.base_theme = copy.deepcopy(self.base_payload['theme'])
        self.palette = normalise_palette(self.base_payload.get('palette') or self.base_theme.get('_palette', []))
        if len(self.palette) < 2:
            self.palette = normalise_palette([self.base_theme.get('accent', '#537988'),
                                              self.base_theme.get('accent_soft', '#CBDDD9')])
        self._original_palette = list(self.palette)
        self.assignments = {role: None for role in ROLES}
        persisted_locks = self.base_payload.get('studio_locks') or {}
        self.locks = {role: bool(persisted_locks.get(role)) for role in ROLES}
        self.undo_stack = []
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
        self.add_colour_button.setEnabled(len(self.palette) < MAX_IMPORT_COLOURS)
        self.source_label.setText(f'{len(self.palette)} available colours • '+
                                  ('unsaved edits' if self.is_dirty() else 'no files altered'))

    def open_colour_wheel(self):
        if len(self.palette) >= MAX_IMPORT_COLOURS:
            QMessageBox.information(self, 'Palette Full',
                f'This palette already has {MAX_IMPORT_COLOURS} unique colours.')
            return
        default_colour = self.palette[-1] if self.palette else '#ABC4AB'
        picker = ColourWheelDialog(default_colour, self)
        if picker.exec() != QDialog.Accepted:
            return
        try:
            updated, added = add_custom_colour(self.palette, picker.selected_hex())
        except ValueError as exc:
            QMessageBox.warning(self, 'Add Colour', str(exc))
            return
        if not added:
            QMessageBox.information(self, 'Already in Palette',
                'That colour is already in your palette. You can select it from the swatches.')
            return
        self.snapshot()
        self.palette = updated
        self._rebuild_swatches()
        self._sync_editors()
        self._refresh_preview()

    def snapshot(self):
        self.undo_stack.append((dict(self.assignments), dict(self.locks), list(self.palette)))
        self.undo_stack = self.undo_stack[-40:]
        self.undo_btn.setEnabled(True)

    def is_dirty(self):
        initial_locks = self.base_payload.get('studio_locks') or {}
        return (self.palette != self._original_palette or
                any(value is not None for value in self.assignments.values()) or
                any(self.locks[r] != bool(initial_locks.get(r)) for r in ROLES))

    def _effective(self):
        return resolve_preview_roles(edited_theme(self.base_theme, self.assignments))

    def _sync_editors(self):
        self._updating = True
        current = self._effective()
        # 'Keep current' means the saved colour, not the possibly unsaved edit.
        saved = resolve_preview_roles(self.base_theme)
        try:
            for role, (combo, lock, marker) in self.controls.items():
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
                lock.setChecked(self.locks[role])
                combo.setEnabled(not self.locks[role])
                marker.setStyleSheet(f'background: {current[role]}; border: 1px solid #777; border-radius: 4px;')
                marker.setToolTip(f'{ROLES[role]}: {current[role]}')
        finally:
            self._updating = False
        self.undo_btn.setEnabled(bool(self.undo_stack))

    def _refresh_preview(self):
        before = resolve_preview_roles(self.base_theme)
        after = self._effective()
        for tab in ('Library', 'Scan', 'Guide'):
            self.canvases[(tab, 'before')].apply_colours(before)
            self.canvases[(tab, 'after')].apply_colours(after)
        for role, (_combo, _lock, marker) in self.controls.items():
            marker.setStyleSheet(f'background: {after[role]}; border: 1px solid #777; border-radius: 4px;')
            marker.setToolTip(f'{ROLES[role]}: {after[role]}')

    def _combo_changed(self, role):
        if self._updating or self.locks[role]:
            return
        value = self.controls[role][0].currentData()
        if value == self.assignments[role]:
            return
        self.snapshot()
        self.assignments[role] = value
        self._refresh_preview()

    def _lock_changed(self, role, locked):
        if self._updating or self.locks[role] == locked:
            return
        self.snapshot()
        self.locks[role] = locked
        self.controls[role][0].setEnabled(not locked)

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
            QMessageBox.information(self, 'Locked Colour', 'Unlock both sections before swapping.')
            return
        current = self._effective()
        self.snapshot()
        self.assignments[one], self.assignments[two] = current[two], current[one]
        self._sync_editors()
        self._refresh_preview()

    def autofill_unlocked(self):
        if all(self.locks.values()):
            QMessageBox.information(self, 'All Colours Locked', 'Unlock a section to automatically colour it.')
            return
        self.snapshot()
        # Prefer every imported swatch once before reusing one. Locked
        # sections retain their existing colours and their palette slots.
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
            QMessageBox.information(self, 'All Colours Locked', 'Unlock at least one section to shuffle.')
            return
        self.snapshot()
        self.assignments = shuffle_roles(self.assignments, self.locks, self.palette, random.Random())
        self._sync_editors()
        self._refresh_preview()

    def undo(self):
        if not self.undo_stack:
            return
        self.assignments, self.locks, self.palette = self.undo_stack.pop()
        self._rebuild_swatches()
        self._sync_editors()
        self._refresh_preview()

    def revert(self):
        if not self.is_dirty():
            return
        self.snapshot()
        self.assignments = {role: None for role in ROLES}
        self.palette = list(self._original_palette)
        self._rebuild_swatches()
        existing = self.base_payload.get('studio_locks') or {}
        self.locks = {role: bool(existing.get(role)) for role in ROLES}
        self._sync_editors()
        self._refresh_preview()

    def _make_payload(self, name):
        data = copy.deepcopy(self.base_payload)
        data['schema'] = max(int(data.get('schema', 2)), 3)
        data['name'] = name
        data['palette'] = list(self.palette)
        data['theme'] = edited_theme(self.base_theme, self.assignments)
        data['theme']['_palette'] = list(self.palette)
        data['studio_locks'] = dict(self.locks)
        data['private_dev_only'] = True
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

    def update_existing(self):
        name = self.source_name
        if name not in self.main.custom_theme_files:
            return
        if not self.is_dirty():
            QMessageBox.information(self, 'No Unsaved Changes', 'There are no changes to save.')
            return
        answer = QMessageBox.question(self, 'Update Imported Theme?',
            f"Update '{name}'? A backup of the previous imported theme will be saved first.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            return
        old_path = self.main.custom_theme_files[name]
        backup_dir = self.main.custom_theme_dir / 'backups'
        try:
            backup_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            shutil.copy2(old_path, backup_dir / f'{old_path.stem}_{stamp}.s3theme')
            self.main._save_theme_payload(old_path, self._make_payload(name))
        except OSError as exc:
            QMessageBox.warning(self, 'Theme Save Failed', str(exc))
            return
        self._finish_save(name)

    def save_as_new(self):
        initial = f'{self.source_name} Variation'
        while initial.casefold() in {name.casefold() for name in self.main.custom_theme_files}:
            initial += ' II'
        name, accepted = QInputDialog.getText(self, 'Save as a New Imported Theme',
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
            answer = QMessageBox.question(self, 'Discard Colour Edits?',
                'Close without saving your edited colours or new locks?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if answer != QMessageBox.Yes:
                return
        super().reject()
