"""Small dependency-free helpers for the CAS Curator custom colour wheel."""
import math


def hue_from_point(x, y, center_x, center_y):
    """Hue in degrees (red at 3 o'clock), increasing counter-clockwise."""
    return math.degrees(math.atan2(center_y - y, x - center_x)) % 360.0


def saturation_value_from_point(x, y, left, top, side):
    """S increases rightward; V increases upward inside the SV square."""
    if side <= 0:
        raise ValueError('SV square must have a positive size')
    saturation = max(0., min(1., (x - left) / side))
    value = max(0., min(1., 1. - (y - top) / side))
    return saturation, value
