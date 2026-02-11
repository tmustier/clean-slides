"""Color helpers for chart rendering."""

from __future__ import annotations

from typing import Any

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR

THEME_COLOR_MAP: dict[str, MSO_THEME_COLOR] = {
    "accent1": MSO_THEME_COLOR.ACCENT_1,
    "accent2": MSO_THEME_COLOR.ACCENT_2,
    "accent3": MSO_THEME_COLOR.ACCENT_3,
    "accent4": MSO_THEME_COLOR.ACCENT_4,
    "accent5": MSO_THEME_COLOR.ACCENT_5,
    "accent6": MSO_THEME_COLOR.ACCENT_6,
    "bg1": MSO_THEME_COLOR.BACKGROUND_1,
    "bg2": MSO_THEME_COLOR.BACKGROUND_2,
    "tx1": MSO_THEME_COLOR.TEXT_1,
    "tx2": MSO_THEME_COLOR.TEXT_2,
    "hlink": MSO_THEME_COLOR.HYPERLINK,
    "folhlink": MSO_THEME_COLOR.FOLLOWED_HYPERLINK,
}


def hex_to_rgb(value: str) -> RGBColor:
    """Parse hex color value (e.g. ``#4472C4``)."""
    normalized = value.strip().lstrip("#")
    if len(normalized) != 6:
        raise ValueError(f"Invalid hex color: {value}")
    r = int(normalized[0:2], 16)
    g = int(normalized[2:4], 16)
    b = int(normalized[4:6], 16)
    return RGBColor(r, g, b)


def normalize_theme_color(value: str) -> MSO_THEME_COLOR | None:
    """Resolve a theme-color alias (e.g. ``accent1``)."""
    key = value.strip().lower().replace("_", "")
    return THEME_COLOR_MAP.get(key)


def resolve_color(value: str | None) -> tuple[RGBColor | None, MSO_THEME_COLOR | None]:
    """Resolve color input as either RGB or theme color."""
    if not value:
        return None, None
    theme = normalize_theme_color(value)
    if theme is not None:
        return None, theme
    try:
        return hex_to_rgb(value), None
    except ValueError:
        return None, None


def apply_color(target: Any, value: RGBColor | str | None) -> bool:
    """Apply a color to a python-pptx color target.

    Returns ``True`` when a color could be resolved and applied.
    """
    if value is None:
        return False
    if isinstance(value, RGBColor):
        target.rgb = value
        return True

    rgb, theme = resolve_color(value)
    if theme is not None:
        target.theme_color = theme
        return True
    if rgb is not None:
        target.rgb = rgb
        return True
    return False
