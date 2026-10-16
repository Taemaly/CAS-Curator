"""Reusable Qt colour-wheel and exact HEX editor for user-defined theme swatches.

Everything happens in this dialog; the parent theme studio remains unchanged
until the user confirms Add Colour. The wheel is drawn locally, no web access.
"""
from __future__ import annotations

from PySide6.QtCore import Qt, QRectF, Signal, QRegularExpression
from PySide6.QtGui import QColor, QPainter, QPen, QImage, QRegularExpressionValidator
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QSlider, QDialogButtonBox, QMessageBox,
)

from colour_wheel_math import hue_from_point, saturation_value_from_point


class HSVColourWheel(QWidget):
    colourChanged = Signal(QColor)
    SIDE = 122
    SIZE = 264
    RADIUS = 109
    RING_WIDTH = 19

    def __init__(self, colour='#ABC4AB', parent=None):
        super().__init__(parent)
        self.setFixedSize(self.SIZE, self.SIZE)
        self.setCursor(Qt.CrossCursor)
        self.hue, self.saturation, self.value = 130., .3, .7
        self._active_part = None
        self._square_image = None
        self.set_colour(QColor(colour), notify=False)

    @property
    def square_left(self):
        return (self.SIZE - self.SIDE) / 2

    @property
    def square_top(self):
        return (self.SIZE - self.SIDE) / 2

    def current_colour(self):
        return QColor.fromHsvF(self.hue / 360., self.saturation, self.value)

    def set_colour(self, colour, notify=True):
        if not isinstance(colour, QColor):
            colour = QColor(colour)
        if not colour.isValid():
            return
        h, s, v, _alpha = colour.getHsvF()
        if h >= 0:
            new_hue = (h * 360.) % 360.
            if round(new_hue, 2) != round(self.hue, 2):
                self._square_image = None
            self.hue = new_hue
        self.saturation, self.value = max(0., s), max(0., v)
        self.update()
        if notify:
            self.colourChanged.emit(self.current_colour())

    def set_value(self, value):
        self.value = max(0., min(1., value))
        self.update()
        self.colourChanged.emit(self.current_colour())

    def _sv_image(self):
        if self._square_image is None:
            image = QImage(self.SIDE, self.SIDE, QImage.Format_RGB32)
            for y in range(self.SIDE):
                v = 1. - y / (self.SIDE - 1)
                for x in range(self.SIDE):
                    s = x / (self.SIDE - 1)
                    image.setPixelColor(x, y, QColor.fromHsvF(self.hue / 360., s, v))
            self._square_image = image
        return self._square_image

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        centre = self.SIZE / 2
        radius = self.RADIUS
        painter.setPen(QPen(QColor('#66717C'), self.RING_WIDTH + 2))
        painter.setBrush(Qt.NoBrush)
        painter.drawEllipse(QRectF(centre - radius, centre - radius, radius * 2, radius * 2))
        rect = QRectF(centre - radius, centre - radius, radius * 2, radius * 2)
        for angle in range(360):
            painter.setPen(QPen(QColor.fromHsv(angle, 255, 255), self.RING_WIDTH))
            painter.drawArc(rect, angle * 16, 18)
        # Select hue on the circular ring. The Qt arc orientation matches hue_from_point.
        import math
        rad = math.radians(self.hue)
        hx = centre + radius * math.cos(rad)
        hy = centre - radius * math.sin(rad)
        painter.setPen(QPen(QColor('#17222A'), 2))
        painter.setBrush(QColor('#FFFFFF'))
        painter.drawEllipse(QRectF(hx - 7, hy - 7, 14, 14))
        left, top = self.square_left, self.square_top
        painter.drawImage(int(left), int(top), self._sv_image())
        painter.setBrush(Qt.NoBrush)
        painter.setPen(QPen(QColor('#3F4D5A'), 1))
        painter.drawRect(QRectF(left, top, self.SIDE, self.SIDE))
        sx = left + self.saturation * (self.SIDE - 1)
        sy = top + (1. - self.value) * (self.SIDE - 1)
        painter.setPen(QPen(QColor('#20272B'), 2))
        painter.setBrush(QColor('#FFFFFF'))
        painter.drawEllipse(QRectF(sx - 5, sy - 5, 10, 10))
        painter.end()

    def _part_at(self, pos):
        import math
        cx = cy = self.SIZE / 2
        dx, dy = pos.x() - cx, pos.y() - cy
        dist = math.hypot(dx, dy)
        if self.RADIUS - self.RING_WIDTH / 2 - 7 <= dist <= self.RADIUS + self.RING_WIDTH / 2 + 7:
            return 'hue'
        if (self.square_left <= pos.x() <= self.square_left + self.SIDE and
                self.square_top <= pos.y() <= self.square_top + self.SIDE):
            return 'sv'
        return None

    def _apply_pos(self, pos):
        if self._active_part == 'hue':
            new_hue = hue_from_point(pos.x(), pos.y(), self.SIZE / 2, self.SIZE / 2)
            if round(new_hue, 2) != round(self.hue, 2):
                self.hue = new_hue
                self._square_image = None
        elif self._active_part == 'sv':
            self.saturation, self.value = saturation_value_from_point(
                pos.x(), pos.y(), self.square_left, self.square_top, self.SIDE - 1)
        else:
            return
        self.update()
        self.colourChanged.emit(self.current_colour())

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._active_part = self._part_at(event.position())
            self._apply_pos(event.position())
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._active_part is not None and event.buttons() & Qt.LeftButton:
            self._apply_pos(event.position())
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._active_part = None
        super().mouseReleaseEvent(event)


class ColourWheelDialog(QDialog):
    """A circular wheel, saturation/value square, brightness slider and exact HEX input."""
    def __init__(self, initial='#ABC4AB', parent=None):
        super().__init__(parent)
        self.setWindowTitle('Add Your Own Colour')
        self.setMinimumWidth(332)
        main = QVBoxLayout(self)
        intro = QLabel('Choose a hue on the wheel, adjust the square, or paste a HEX colour.')
        intro.setWordWrap(True)
        main.addWidget(intro)
        self.wheel = HSVColourWheel(initial, self)
        main.addWidget(self.wheel, alignment=Qt.AlignHCenter)
        slider_row = QHBoxLayout()
        slider_row.addWidget(QLabel('Brightness'))
        self.brightness = QSlider(Qt.Horizontal, self)
        self.brightness.setRange(0, 100)
        self.brightness.setValue(round(self.wheel.value * 100))
        self.brightness.valueChanged.connect(lambda v: self.wheel.set_value(v / 100.))
        slider_row.addWidget(self.brightness, 1)
        main.addLayout(slider_row)
        hex_row = QHBoxLayout()
        hex_row.addWidget(QLabel('HEX'))
        self.hex_input = QLineEdit(self)
        self.hex_input.setMaxLength(7)
        self.hex_input.setValidator(QRegularExpressionValidator(
            QRegularExpression(r'#[0-9a-fA-F]{0,6}'), self.hex_input))
        self.hex_input.setText(self.wheel.current_colour().name(QColor.HexRgb).upper())
        self.hex_input.setToolTip('Enter six HEX digits, e.g. #D5A6BD')
        self.hex_input.editingFinished.connect(self._read_hex)
        hex_row.addWidget(self.hex_input, 1)
        self.sample = QLabel(self)
        self.sample.setFixedSize(64, 26)
        hex_row.addWidget(self.sample)
        main.addLayout(hex_row)
        self.wheel.colourChanged.connect(self._wheel_changed)
        self._wheel_changed(self.wheel.current_colour())
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.button(QDialogButtonBox.Ok).setText('Add Colour')
        buttons.accepted.connect(self._accept_hex)
        buttons.rejected.connect(self.reject)
        main.addWidget(buttons)

    def _wheel_changed(self, colour):
        self.brightness.blockSignals(True)
        self.brightness.setValue(round(self.wheel.value * 100))
        self.brightness.blockSignals(False)
        self.hex_input.setText(colour.name(QColor.HexRgb).upper())
        self.sample.setStyleSheet(f'background: {colour.name()}; border: 1px solid #43515E; border-radius: 4px;')

    def _read_hex(self):
        text = self.hex_input.text().strip()
        if len(text) == 7 and text.startswith('#'):
            colour = QColor(text)
            if colour.isValid():
                self.wheel.set_colour(colour)
                return True
        return False

    def _accept_hex(self):
        if not self._read_hex():
            QMessageBox.warning(self, 'Invalid Colour', 'Enter a valid six-digit HEX colour such as #D5A6BD.')
            self.hex_input.setFocus()
            return
        self.accept()

    def selected_hex(self):
        return self.wheel.current_colour().name(QColor.HexRgb).upper()
