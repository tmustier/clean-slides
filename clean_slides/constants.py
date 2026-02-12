"""
Design system constants derived from the template configuration.

All values are loaded from :data:`template_config.TEMPLATE_CONFIG` and
can be refreshed at runtime via :func:`_reload`.
"""

from __future__ import annotations

from typing import Any

from pptx.dml.color import RGBColor
from pptx.util import Emu, Inches, Pt

from .template_config import TEMPLATE_CONFIG, TemplateConfig

# ---------------------------------------------------------------------------
# Unit conversions (not template-specific)
# ---------------------------------------------------------------------------

EMU_PER_INCH = 914_400  # 1 inch = 914,400 EMU
EMU_PER_PT = 12_700  # 1 point = 12,700 EMU
OOXML_FONT_SCALE = 100  # OOXML stores font sizes in hundredths of a point


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_DEFAULT_COLORS: dict[str, str] = {
    "body_text": "tx1",
    "col_header": "tx1",
    "col_superheader": "tx1",
    "row_header": "accent2",
    "row_superheader": "tx1",
    "divider": "tx1",
    "link": "accent2",
}


def _section(config: TemplateConfig, key: str) -> dict[str, Any]:
    return config.section(key)


def _rgb(value: str) -> RGBColor:
    rgb_hex_str = value.lstrip("#")
    if len(rgb_hex_str) != 6:
        raise ValueError(f"Invalid RGB hex string: {value!r}")
    r = int(rgb_hex_str[0:2], 16)
    g = int(rgb_hex_str[2:4], 16)
    b = int(rgb_hex_str[4:6], 16)
    return RGBColor(r, g, b)


def _rgb_optional(value: object) -> RGBColor | None:
    if not value:
        return None
    return _rgb(str(value))


# ============================================================================
# COLORS
# ============================================================================


class Colors:
    """Design system color palette."""

    MIDNIGHT: RGBColor
    LIGHT_1: RGBColor
    LIGHT_2: RGBColor
    SLATE: RGBColor
    ELECTRIC_BLUE: RGBColor
    CYAN: RGBColor
    RED: RGBColor
    ORANGE: RGBColor
    BEIGE: RGBColor
    LIGHT_GREEN: RGBColor
    GREEN: RGBColor


class Fonts:
    """Font families."""

    HEADLINE: str
    BODY: str


class FontSizes:
    """Standard font sizes in points."""

    TITLE: int
    SUBTITLE: int
    TABLE_HEADER: int
    TABLE_BODY: int
    FOOTNOTE: int
    TRACKER: int
    DEFAULT: int


class Layout:
    """Slide layout measurements."""

    SLIDE_WIDTH: Emu
    SLIDE_HEIGHT: Emu
    LEFT_MARGIN: Emu
    RIGHT_MARGIN: Emu
    TOP_MARGIN: Emu
    TRACKER_Y: Emu
    TITLE_Y: Emu
    SUBTITLE_Y: Emu
    HEADER_LINE_Y: Emu
    CONTENT_START_Y: Emu
    FOOTER_LINE_Y: Emu
    FOOTER_Y: Emu
    CONTENT_WIDTH: Emu
    CONTENT_HEIGHT: Emu


class Dividers:
    """Line weights for dividers."""

    HEADER: Pt
    ROW: Pt
    FOOTER: Pt


class DefaultColors:
    """Default theme color names used when YAML doesn't override."""

    BODY_TEXT: str
    COL_HEADER: str
    COL_SUPERHEADER: str
    ROW_HEADER: str
    ROW_SUPERHEADER: str
    DIVIDER: str
    LINK: str


class TableDefaults:
    """Default table styling."""

    MIN_ROW_HEIGHT: Inches
    CELL_PADDING: Emu
    HEADER_ROW_HEIGHT: Inches
    WIDTH_STEP: Emu
    DEFAULT_WIDTH: float
    ROW_HEADER_WIDTH: float
    MOON_WIDTH: float
    BULLETS_WIDTH: float
    SUPERHEADER_PT_BOOST: int
    LINE_SPACING: float
    LEGEND_LABEL_PT: int
    LEGEND_Y_OFFSET_INCHES: float
    LEGEND_ROW_HEIGHT_RATIO: float
    LEGEND_GAP_RATIO: float
    LEGEND_CHAR_WIDTH_RATIO: float
    SHAPE_ID_START: int
    MAX_HEADER_LINES: int
    ROW_HEADER_TARGET_LINES: int


# Module-level mutable state (rewritten by _reload).
MOON_COLORS: dict[str, RGBColor | None] = {}
MOON_COLORS_HEX: dict[str, str | None] = {}
MOON_FILLS: dict[str, int] = {}
MOON_ARC_ADJUSTMENTS: dict[int, tuple[str, ...] | None] = {}
MOON_SIZE_EMU: int = 0
MOON_GROUP: dict[str, Any] = {}
ICON_DEFAULT_SIZE_EMU: int = 0

BULLET_LEVELS: list[dict[str, Any]] = []
BULLET_CHARS: dict[int, str] = {}
BULLET_MARGINS: dict[int, tuple[int, int, float]] = {}
BULLET_DEF_TAB_SZ_EMU: int = 0
BULLET_BU_FONT: dict[str, Any] = {}
BULLET_DEF_RPR: dict[str, Any] = {}


# ============================================================================
# Initialization / reload
# ============================================================================


def reload_constants(config: TemplateConfig | None = None) -> None:
    """(Re-)initialize all constants from *config*.

    Called once at module load and again by
    :func:`template_config.set_template_config`.
    """
    global MOON_COLORS, MOON_COLORS_HEX, MOON_FILLS, MOON_ARC_ADJUSTMENTS
    global MOON_SIZE_EMU, MOON_GROUP, ICON_DEFAULT_SIZE_EMU
    global BULLET_LEVELS, BULLET_CHARS, BULLET_MARGINS
    global BULLET_DEF_TAB_SZ_EMU, BULLET_BU_FONT, BULLET_DEF_RPR

    if config is None:
        config = TEMPLATE_CONFIG

    c = _section(config, "colors")
    f = _section(config, "fonts")
    fs = _section(config, "font_sizes")
    lo = _section(config, "layout")
    dv = _section(config, "dividers")
    bu = _section(config, "bullets")
    td = _section(config, "table_defaults")
    ic = _section(config, "icons")
    mo = _section(config, "moon")
    dc = config.section_or_default("default_colors", _DEFAULT_COLORS)

    # -- Colors --
    Colors.MIDNIGHT = _rgb(c["midnight"])
    Colors.LIGHT_1 = _rgb(c["light_1"])
    Colors.LIGHT_2 = _rgb(c["light_2"])
    Colors.SLATE = _rgb(c["slate"])
    Colors.ELECTRIC_BLUE = _rgb(c["electric_blue"])
    Colors.CYAN = _rgb(c["cyan"])
    Colors.RED = _rgb(c["red"])
    Colors.ORANGE = _rgb(c["orange"])
    Colors.BEIGE = _rgb(c["beige"])
    Colors.LIGHT_GREEN = _rgb(c["light_green"])
    Colors.GREEN = _rgb(c["green"])

    # -- Moon --
    MOON_COLORS.clear()
    MOON_COLORS.update({key: _rgb_optional(val) for key, val in mo["colors"].items()})
    MOON_COLORS_HEX.clear()
    MOON_COLORS_HEX.update(
        {key: (val.lstrip("#") if val else None) for key, val in mo["colors"].items()}
    )
    MOON_FILLS.clear()
    MOON_FILLS.update({key: int(val) for key, val in mo["fills"].items()})
    MOON_ARC_ADJUSTMENTS.clear()
    MOON_ARC_ADJUSTMENTS.update(
        {
            int(key): tuple(val) if val is not None else None
            for key, val in mo["arc_adjustments"].items()
        }
    )
    globals()["MOON_SIZE_EMU"] = int(mo["size_emu"])
    MOON_GROUP.clear()
    MOON_GROUP.update(mo["group"])

    globals()["ICON_DEFAULT_SIZE_EMU"] = int(ic["default_size_emu"])

    # -- Fonts --
    Fonts.HEADLINE = str(f["headline"])
    Fonts.BODY = str(f["body"])

    FontSizes.TITLE = int(fs["title"])
    FontSizes.SUBTITLE = int(fs["subtitle"])
    FontSizes.TABLE_HEADER = int(fs["table_header"])
    FontSizes.TABLE_BODY = int(fs["table_body"])
    FontSizes.FOOTNOTE = int(fs["footnote"])
    FontSizes.TRACKER = int(fs["tracker"])
    FontSizes.DEFAULT = int(fs["default"])

    # -- Layout --
    Layout.SLIDE_WIDTH = Emu(lo["slide_width_emu"])
    Layout.SLIDE_HEIGHT = Emu(lo["slide_height_emu"])
    Layout.LEFT_MARGIN = Emu(lo["left_margin_emu"])
    Layout.RIGHT_MARGIN = Emu(lo["right_margin_emu"])
    Layout.TOP_MARGIN = Emu(lo["top_margin_emu"])
    Layout.TRACKER_Y = Emu(lo["tracker_y_emu"])
    Layout.TITLE_Y = Emu(lo["title_y_emu"])
    Layout.SUBTITLE_Y = Emu(lo["subtitle_y_emu"])
    Layout.HEADER_LINE_Y = Emu(lo["header_line_y_emu"])
    Layout.CONTENT_START_Y = Emu(lo["content_start_y_emu"])
    Layout.FOOTER_LINE_Y = Emu(lo["footer_line_y_emu"])
    Layout.FOOTER_Y = Emu(lo["footer_y_emu"])
    Layout.CONTENT_WIDTH = Emu(
        int(Layout.SLIDE_WIDTH) - int(Layout.LEFT_MARGIN) - int(Layout.RIGHT_MARGIN)
    )
    Layout.CONTENT_HEIGHT = Emu(int(Layout.FOOTER_LINE_Y) - int(Layout.CONTENT_START_Y))

    # -- Dividers --
    Dividers.HEADER = Pt(dv["header_pt"])
    Dividers.ROW = Pt(dv["row_pt"])
    Dividers.FOOTER = Pt(dv["footer_pt"])

    # -- Default colors --
    DefaultColors.BODY_TEXT = str(dc.get("body_text", "tx1"))
    DefaultColors.COL_HEADER = str(dc.get("col_header", "tx1"))
    DefaultColors.COL_SUPERHEADER = str(dc.get("col_superheader", "tx1"))
    DefaultColors.ROW_HEADER = str(dc.get("row_header", "accent2"))
    DefaultColors.ROW_SUPERHEADER = str(dc.get("row_superheader", "tx1"))
    DefaultColors.DIVIDER = str(dc.get("divider", "tx1"))
    DefaultColors.LINK = str(dc.get("link", "accent2"))

    # -- Bullets --
    levels = list(bu["levels"])
    BULLET_LEVELS.clear()
    BULLET_LEVELS.extend(levels)
    BULLET_CHARS.clear()
    BULLET_CHARS.update({int(lv["level"]): lv.get("bullet", "") for lv in levels})
    BULLET_MARGINS.clear()
    BULLET_MARGINS.update(
        {
            int(lv["level"]): (int(lv["mar_l_emu"]), int(lv["indent_emu"]), float(lv["spc_bef_pt"]))
            for lv in levels
        }
    )
    globals()["BULLET_DEF_TAB_SZ_EMU"] = int(bu["def_tab_sz_emu"])
    BULLET_BU_FONT.clear()
    BULLET_BU_FONT.update(bu["bu_font"])
    BULLET_DEF_RPR.clear()
    BULLET_DEF_RPR.update(bu["def_rpr"])

    # -- Table defaults --
    legend = td["legend"]
    TableDefaults.MIN_ROW_HEIGHT = Inches(td["min_row_height_in"])
    TableDefaults.CELL_PADDING = Emu(td["cell_padding_emu"])
    TableDefaults.HEADER_ROW_HEIGHT = Inches(td["header_row_height_in"])
    TableDefaults.WIDTH_STEP = Emu(td["width_step_emu"])
    TableDefaults.DEFAULT_WIDTH = float(td["default_width_in"])
    TableDefaults.ROW_HEADER_WIDTH = float(td["row_header_width_in"])
    TableDefaults.MOON_WIDTH = float(td["moon_width_in"])
    TableDefaults.BULLETS_WIDTH = float(td["bullets_width_in"])
    TableDefaults.SUPERHEADER_PT_BOOST = int(td["superheader_pt_boost"])
    TableDefaults.LINE_SPACING = float(td["line_spacing"])
    TableDefaults.LEGEND_LABEL_PT = int(legend["label_pt"])
    TableDefaults.LEGEND_Y_OFFSET_INCHES = float(legend["y_offset_in"])
    TableDefaults.LEGEND_ROW_HEIGHT_RATIO = float(legend["row_height_ratio"])
    TableDefaults.LEGEND_GAP_RATIO = float(legend["gap_ratio"])
    TableDefaults.LEGEND_CHAR_WIDTH_RATIO = float(legend["char_width_ratio"])
    TableDefaults.SHAPE_ID_START = int(td["shape_id_start"])
    TableDefaults.MAX_HEADER_LINES = int(td["max_header_lines"])
    TableDefaults.ROW_HEADER_TARGET_LINES = int(td["row_header_target_lines"])


# Run initialization at import time.
reload_constants()
