"""Chart style application helpers."""

# pyright: reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false
# pyright: reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false
# pyright: reportMissingParameterType=false
# pyright: reportMissingTypeArgument=false
# pyright: reportArgumentType=false
# pyright: reportAttributeAccessIssue=false
# pyright: reportUnnecessaryIsInstance=false

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_LABEL_POSITION, XL_TICK_LABEL_POSITION, XL_TICK_MARK
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

from .colors import apply_color, hex_to_rgb, normalize_theme_color, resolve_color
from .defaults import DEFAULT_BAR_SERIES_BORDER_COLOR, DEFAULT_WATERFALL_LABEL_FONT_SIZE
from .geometry import normalize_orientation


def apply_series_colors(chart, colors: list[str | None]) -> None:
    for idx, color in enumerate(colors):
        if not color:
            continue
        ser = chart.series[idx]
        theme = normalize_theme_color(color) if isinstance(color, str) else None
        if theme is not None:
            ser.format.fill.solid()
            ser.format.fill.fore_color.theme_color = theme
            continue
        try:
            rgb = hex_to_rgb(color)
        except ValueError:
            continue
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb


def apply_waterfall_style(chart, meta: dict) -> None:
    if not meta:
        return
    offset_idx = meta.get("offset_series_idx", 0)
    offset = chart.series[offset_idx]
    if meta.get("offset_no_fill"):
        offset.format.fill.background()
        offset.format.line.fill.background()
    for idx, color in meta.get("offset_points", []):
        if not color:
            continue
        try:
            rgb = hex_to_rgb(color)
        except ValueError:
            continue
        point = offset.points[idx]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = rgb
        point.format.line.fill.background()


def apply_plot_layout(chart, layout: dict) -> None:
    if not layout:
        return
    chart_space = chart._element
    chart_el = chart_space.find(qn("c:chart"))
    if chart_el is None:
        return
    plot_area = chart_el.find(qn("c:plotArea"))
    if plot_area is None:
        return

    existing = plot_area.find(qn("c:layout"))
    if existing is not None:
        plot_area.remove(existing)

    layout_el = OxmlElement("c:layout")
    manual = OxmlElement("c:manualLayout")

    def add_child(tag: str, val: str) -> None:
        child = OxmlElement(tag)
        child.set("val", val)
        manual.append(child)

    add_child("c:layoutTarget", "inner")
    add_child("c:xMode", "edge")
    add_child("c:yMode", "edge")
    add_child("c:x", str(layout.get("x")))
    add_child("c:y", str(layout.get("y")))
    add_child("c:w", str(layout.get("w")))
    add_child("c:h", str(layout.get("h")))

    layout_el.append(manual)
    plot_area.insert(0, layout_el)


def apply_waterfall_data_labels(chart, meta: dict) -> None:
    if not meta:
        return
    for series in chart.series:
        labels = series.data_labels
        labels.show_value = True
        labels.position = XL_LABEL_POSITION.CENTER
        label_decimals = meta.get("data_label_decimals")
        label_format = meta.get("data_label_format")
        if label_format is None:
            if label_decimals is None:
                label_decimals = 0
            pattern = "0" if int(label_decimals) <= 0 else f"0.{('0' * int(label_decimals))}"
            label_format = f"{pattern};{pattern}"
        labels.number_format = label_format
        labels.number_format_is_linked = False
        labels.font.size = DEFAULT_WATERFALL_LABEL_FONT_SIZE
        color_value = meta.get("data_label_color")
        if color_value is None:
            color_value = "tx1"
        apply_color(labels.font.color, color_value)


def apply_plot_spacing(plot, meta: dict) -> None:
    gap_width = meta.get("gap_width")
    if gap_width is not None:
        plot.gap_width = int(gap_width)
    overlap = meta.get("overlap")
    if overlap is not None:
        plot.overlap = int(overlap)
    plot.vary_by_categories = False


def set_axis_orientation(axis, value: str) -> None:
    ax = axis._element
    scaling = ax.find(qn("c:scaling"))
    if scaling is None:
        scaling = OxmlElement("c:scaling")
        ax.insert(0, scaling)
    orient = scaling.find(qn("c:orientation"))
    if orient is None:
        orient = OxmlElement("c:orientation")
        scaling.insert(0, orient)
    orient.set("val", value)


def apply_axis_style(
    chart,
    axis_min: float | None,
    axis_max: float | None,
    axis_line_color: RGBColor | str | None = None,
    orientation: str | None = None,
) -> None:
    cat_axis = chart.category_axis
    val_axis = chart.value_axis
    orientation = normalize_orientation(orientation)

    cat_axis.has_major_gridlines = False
    cat_axis.has_minor_gridlines = False
    cat_axis.major_tick_mark = XL_TICK_MARK.NONE
    cat_axis.minor_tick_mark = XL_TICK_MARK.NONE
    cat_axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE

    val_axis.has_major_gridlines = False
    val_axis.has_minor_gridlines = False
    val_axis.major_tick_mark = XL_TICK_MARK.NONE
    val_axis.minor_tick_mark = XL_TICK_MARK.NONE
    val_axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE

    if orientation == "horizontal":
        cat_axis.format.line.width = Pt(0.75)
        if axis_line_color is not None:
            apply_color(cat_axis.format.line.color, axis_line_color)
        else:
            cat_axis.format.line.color.rgb = RGBColor(0, 0, 0)
        set_axis_orientation(cat_axis, "maxMin")
        if hasattr(val_axis, "visible"):
            val_axis.visible = False
        else:
            val_axis.format.line.fill.background()
            val_axis.format.line.width = Pt(0)
    else:
        cat_axis.format.line.width = Pt(0.75)
        if axis_line_color is not None:
            apply_color(cat_axis.format.line.color, axis_line_color)
        else:
            cat_axis.format.line.color.rgb = RGBColor(0, 0, 0)
        val_axis.visible = False

    if axis_min is not None:
        val_axis.minimum_scale = float(axis_min)
    if axis_max is not None:
        val_axis.maximum_scale = float(axis_max)


def apply_waterfall_chart_style(chart, meta: dict) -> None:
    if not meta:
        return
    plot = chart.plots[0]
    apply_plot_spacing(plot, meta)
    apply_axis_style(
        chart,
        meta.get("axis_min"),
        meta.get("axis_max"),
        meta.get("axis_line_color"),
        meta.get("orientation"),
    )
    apply_plot_layout(chart, meta.get("plot_layout"))


def apply_bar_chart_style(chart, meta: dict) -> None:
    if not meta:
        return
    plot = chart.plots[0]
    apply_plot_spacing(plot, meta)
    apply_axis_style(
        chart,
        meta.get("axis_min"),
        meta.get("axis_max"),
        meta.get("axis_line_color"),
        meta.get("orientation"),
    )

    border_color = meta.get("series_border_color", DEFAULT_BAR_SERIES_BORDER_COLOR)
    disable_border = False
    if border_color is None or (
        isinstance(border_color, str)
        and border_color.strip().lower()
        in {
            "none",
            "transparent",
            "false",
        }
    ):
        disable_border = True

    for series in chart.series:
        if hasattr(series, "invert_if_negative"):
            series.invert_if_negative = False

    if disable_border:
        for series in chart.series:
            series.format.line.fill.background()
            series.format.line.width = Pt(0)
    else:
        rgb, theme = resolve_color(border_color)
        for series in chart.series:
            series.format.line.width = Pt(0.75)
            if theme is not None:
                series.format.line.color.theme_color = theme
            elif rgb is not None:
                series.format.line.color.rgb = rgb
            else:
                series.format.line.color.rgb = RGBColor(255, 255, 255)

    apply_plot_layout(chart, meta.get("plot_layout"))
