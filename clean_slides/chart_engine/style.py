"""Chart style application helpers."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from typing import Any, Union, cast

from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_LABEL_POSITION, XL_TICK_LABEL_POSITION, XL_TICK_MARK
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

from .colors import apply_color, hex_to_rgb, resolve_color
from .defaults import DEFAULT_BAR_SERIES_BORDER_COLOR, DEFAULT_WATERFALL_LABEL_FONT_SIZE
from .geometry import normalize_orientation

Number = Union[int, float]
ColorValue = Union[RGBColor, str, None]


def _number(value: object, default: Number = 0) -> Number:
    return value if isinstance(value, (int, float)) else default


def _int(value: object, default: int = 0) -> int:
    return int(_number(value, default))


def _float(value: object, default: float = 0.0) -> float:
    return float(_number(value, default))


def _offset_points(value: object) -> list[tuple[int, str]]:
    points: list[tuple[int, str]] = []
    if not isinstance(value, list):
        return points

    for item in cast(list[object], value):
        if isinstance(item, tuple):
            values = cast(tuple[object, ...], item)
        elif isinstance(item, list):
            values = tuple(cast(list[object], item))
        else:
            continue

        if len(values) < 2:
            continue

        idx = values[0]
        color = values[1]
        if isinstance(idx, int) and isinstance(color, str):
            points.append((idx, color))

    return points


def _layout(value: object) -> dict[str, float]:
    if not isinstance(value, dict):
        return {}

    typed_mapping = cast(dict[object, object], value)
    layout = {
        key: float(item)
        for key, item in typed_mapping.items()
        if isinstance(key, str) and isinstance(item, (int, float))
    }
    return layout


def apply_series_colors(chart: Any, colors: Sequence[str | None]) -> None:
    for idx, color in enumerate(colors):
        if not color:
            continue
        if idx >= len(chart.series):
            continue

        series = chart.series[idx]
        rgb, theme = resolve_color(color)
        if theme is not None:
            series.format.fill.solid()
            series.format.fill.fore_color.theme_color = theme
            continue
        if rgb is None:
            continue

        series.format.fill.solid()
        series.format.fill.fore_color.rgb = rgb


def apply_waterfall_style(chart: Any, meta: Mapping[str, object]) -> None:
    if not meta:
        return

    offset_idx = _int(meta.get("offset_series_idx", 0), 0)
    if offset_idx < 0 or offset_idx >= len(chart.series):
        return

    offset_series = chart.series[offset_idx]
    if bool(meta.get("offset_no_fill")):
        offset_series.format.fill.background()
        offset_series.format.line.fill.background()

    for idx, color in _offset_points(meta.get("offset_points")):
        if idx < 0 or idx >= len(offset_series.points):
            continue

        try:
            rgb = hex_to_rgb(color)
        except ValueError:
            continue

        point = offset_series.points[idx]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = rgb
        point.format.line.fill.background()


def apply_plot_layout(chart: Any, layout_value: object) -> None:
    layout = _layout(layout_value)
    if not layout:
        return

    chart_space = chart._element
    chart_element = chart_space.find(qn("c:chart"))
    if chart_element is None:
        return

    plot_area = chart_element.find(qn("c:plotArea"))
    if plot_area is None:
        return

    existing = plot_area.find(qn("c:layout"))
    if existing is not None:
        plot_area.remove(existing)

    layout_element = OxmlElement("c:layout")
    manual_layout = OxmlElement("c:manualLayout")

    def add_child(tag: str, value: str) -> None:
        child = OxmlElement(tag)
        child.set("val", value)
        manual_layout.append(child)

    add_child("c:layoutTarget", "inner")
    add_child("c:xMode", "edge")
    add_child("c:yMode", "edge")
    add_child("c:x", str(layout.get("x")))
    add_child("c:y", str(layout.get("y")))
    add_child("c:w", str(layout.get("w")))
    add_child("c:h", str(layout.get("h")))

    layout_element.append(manual_layout)
    plot_area.insert(0, layout_element)


def apply_waterfall_data_labels(chart: Any, meta: Mapping[str, object]) -> None:
    if not meta:
        return

    for series in chart.series:
        labels = series.data_labels
        labels.show_value = True
        labels.position = XL_LABEL_POSITION.CENTER

        label_decimals = meta.get("data_label_decimals")
        label_format = meta.get("data_label_format")
        if not isinstance(label_format, str):
            decimals = _int(label_decimals, 0)
            pattern = "0" if decimals <= 0 else f"0.{('0' * decimals)}"
            label_format = f"{pattern};{pattern}"

        labels.number_format = label_format
        labels.number_format_is_linked = False
        labels.font.size = DEFAULT_WATERFALL_LABEL_FONT_SIZE

        color_value = meta.get("data_label_color")
        if not isinstance(color_value, (RGBColor, str)):
            color_value = "tx1"
        apply_color(labels.font.color, color_value)


def apply_plot_spacing(plot: Any, meta: Mapping[str, object]) -> None:
    gap_width = meta.get("gap_width")
    if isinstance(gap_width, (int, float)):
        plot.gap_width = int(gap_width)

    overlap = meta.get("overlap")
    if isinstance(overlap, (int, float)):
        plot.overlap = int(overlap)

    plot.vary_by_categories = False


def set_axis_orientation(axis: Any, value: str) -> None:
    axis_element = axis._element
    scaling = axis_element.find(qn("c:scaling"))
    if scaling is None:
        scaling = OxmlElement("c:scaling")
        axis_element.insert(0, scaling)

    orient = scaling.find(qn("c:orientation"))
    if orient is None:
        orient = OxmlElement("c:orientation")
        scaling.insert(0, orient)

    orient.set("val", value)


def apply_axis_style(
    chart: Any,
    axis_min: float | None,
    axis_max: float | None,
    axis_line_color: ColorValue = None,
    orientation: str | None = None,
) -> None:
    category_axis = chart.category_axis
    value_axis = chart.value_axis
    orientation_value = normalize_orientation(orientation)

    category_axis.has_major_gridlines = False
    category_axis.has_minor_gridlines = False
    category_axis.major_tick_mark = XL_TICK_MARK.NONE
    category_axis.minor_tick_mark = XL_TICK_MARK.NONE
    category_axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE

    value_axis.has_major_gridlines = False
    value_axis.has_minor_gridlines = False
    value_axis.major_tick_mark = XL_TICK_MARK.NONE
    value_axis.minor_tick_mark = XL_TICK_MARK.NONE
    value_axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE

    category_axis.format.line.width = Pt(0.75)
    if axis_line_color is not None:
        apply_color(category_axis.format.line.color, axis_line_color)
    else:
        category_axis.format.line.color.rgb = RGBColor(0, 0, 0)

    if orientation_value == "horizontal":
        set_axis_orientation(category_axis, "maxMin")
        if hasattr(value_axis, "visible"):
            value_axis.visible = False
        else:
            value_axis.format.line.fill.background()
            value_axis.format.line.width = Pt(0)
    else:
        value_axis.visible = False

    if axis_min is not None:
        value_axis.minimum_scale = _float(axis_min)
    if axis_max is not None:
        value_axis.maximum_scale = _float(axis_max)


def apply_waterfall_chart_style(chart: Any, meta: Mapping[str, object]) -> None:
    if not meta:
        return

    plot = chart.plots[0]
    apply_plot_spacing(plot, meta)
    apply_axis_style(
        chart,
        _float(meta.get("axis_min")) if meta.get("axis_min") is not None else None,
        _float(meta.get("axis_max")) if meta.get("axis_max") is not None else None,
        cast(ColorValue, meta.get("axis_line_color")),
        cast(str | None, meta.get("orientation")),
    )
    apply_plot_layout(chart, meta.get("plot_layout"))


def apply_bar_chart_style(chart: Any, meta: Mapping[str, object]) -> None:
    if not meta:
        return

    plot = chart.plots[0]
    apply_plot_spacing(plot, meta)
    apply_axis_style(
        chart,
        _float(meta.get("axis_min")) if meta.get("axis_min") is not None else None,
        _float(meta.get("axis_max")) if meta.get("axis_max") is not None else None,
        cast(ColorValue, meta.get("axis_line_color")),
        cast(str | None, meta.get("orientation")),
    )

    border_color_raw = meta.get("series_border_color", DEFAULT_BAR_SERIES_BORDER_COLOR)
    border_color: ColorValue
    if isinstance(border_color_raw, (RGBColor, str)) or border_color_raw is None:
        border_color = border_color_raw
    else:
        border_color = DEFAULT_BAR_SERIES_BORDER_COLOR

    disable_border = False
    if border_color is None or (
        isinstance(border_color, str)
        and border_color.strip().lower() in {"none", "transparent", "false"}
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
        rgb, theme = resolve_color(border_color if isinstance(border_color, str) else None)
        for series in chart.series:
            series.format.line.width = Pt(0.75)
            if isinstance(border_color, RGBColor):
                series.format.line.color.rgb = border_color
            elif theme is not None:
                series.format.line.color.theme_color = theme
            elif rgb is not None:
                series.format.line.color.rgb = rgb
            else:
                series.format.line.color.rgb = RGBColor(255, 255, 255)

    apply_plot_layout(chart, meta.get("plot_layout"))
