"""Connector rendering helpers for waterfall overlays."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from typing import Protocol, Union, cast

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from .colors import apply_color
from .geometry import value_to_x, value_to_y

Number = Union[int, float]
FloatOrNone = Union[float, None]
ColorValue = Union[RGBColor, str, None]
Geometry = Mapping[str, object]


class _ShapeForeColorLike(Protocol):
    rgb: object


class _ShapeFillLike(Protocol):
    fore_color: _ShapeForeColorLike

    def solid(self) -> None: ...


class _ShapeLineFillLike(Protocol):
    def background(self) -> None: ...


class _ShapeLineLike(Protocol):
    fill: _ShapeLineFillLike


class _ShapeLike(Protocol):
    fill: _ShapeFillLike
    line: _ShapeLineLike


class _SlideShapesLike(Protocol):
    def add_shape(
        self,
        shape_type: object,
        left: object,
        top: object,
        width: object,
        height: object,
    ) -> _ShapeLike: ...


class _SlideLike(Protocol):
    shapes: _SlideShapesLike


def _geometry_series(geometry: Geometry, key: str) -> list[float]:
    raw_values = geometry.get(key)
    if not isinstance(raw_values, list):
        return []

    typed_values = cast(list[object], raw_values)
    series: list[float] = []
    for item in typed_values:
        if isinstance(item, (int, float)):
            series.append(float(item))
    return series


def _connector_value(values: Sequence[FloatOrNone], idx: int) -> FloatOrNone:
    if idx < 0 or idx >= len(values):
        return None
    value = values[idx]
    if isinstance(value, (int, float)):
        return float(value)
    return None


def render_waterfall_connectors(
    slide: object,
    categories: Sequence[object],
    connector_values: Sequence[FloatOrNone],
    geometry: Geometry,
    orientation: str,
    axis_min: float,
    axis_max: float,
    plot_left: float,
    plot_top: float,
    plot_width: float,
    plot_height: float,
    connector_style: str,
    connector_inset: Number,
    connector_overlap: Number,
    line_width: int,
    dash_length: Union[int, None],
    dash_gap: int,
    connector_color: ColorValue,
) -> None:
    """Render connector segments between adjacent waterfall categories."""
    slide_like = cast(_SlideLike, slide)

    def add_dash_segment(x: float, y: float, width: float, height: float) -> None:
        if width <= 0 or height <= 0:
            return
        shape = slide_like.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            round(x),
            round(y),
            round(width),
            round(height),
        )
        shape.fill.solid()
        if connector_color is not None:
            apply_color(shape.fill.fore_color, connector_color)
        else:
            shape.fill.fore_color.rgb = RGBColor(0, 0, 0)
        shape.line.fill.background()

    def add_connector(x1: float, y1: float, x2: float, y2: float) -> None:
        if abs(x1 - x2) <= abs(y1 - y2):
            x = min(x1, x2) - line_width / 2
            start = min(y1, y2)
            total = abs(y2 - y1)
            if dash_length is None or dash_length <= 0:
                add_dash_segment(x, start, line_width, total)
                return

            pos = 0.0
            while pos < total:
                segment_length = min(dash_length, total - pos)
                add_dash_segment(x, start + pos, line_width, segment_length)
                pos += segment_length + dash_gap
            return

        y = min(y1, y2) - line_width / 2
        start = min(x1, x2)
        total = abs(x2 - x1)
        if dash_length is None or dash_length <= 0:
            add_dash_segment(start, y, total, line_width)
            return

        pos = 0.0
        while pos < total:
            segment_length = min(dash_length, total - pos)
            add_dash_segment(start + pos, y, segment_length, line_width)
            pos += segment_length + dash_gap

    bar_bottoms = _geometry_series(geometry, "bar_bottoms")
    bar_tops = _geometry_series(geometry, "bar_tops")
    bar_rights = _geometry_series(geometry, "bar_rights")
    bar_lefts = _geometry_series(geometry, "bar_lefts")

    inset = float(connector_inset)
    overlap = float(connector_overlap)

    for idx in range(len(categories) - 1):
        current_value = _connector_value(connector_values, idx)
        if current_value is None:
            continue

        next_value = _connector_value(connector_values, idx + 1)

        if orientation == "horizontal":
            if idx >= len(bar_bottoms) or idx + 1 >= len(bar_tops):
                continue

            x_pos = value_to_x(current_value, axis_min, axis_max, plot_left, plot_width) - inset
            base_start = bar_bottoms[idx]
            base_end = bar_tops[idx + 1]
            y_start = base_start - overlap
            y_end = base_end + overlap
            add_connector(x_pos, y_start, x_pos, y_end)

            if connector_style == "step" and next_value is not None:
                next_x = value_to_x(next_value, axis_min, axis_max, plot_left, plot_width) - inset
                if next_x != x_pos:
                    add_connector(x_pos, base_end, next_x, base_end)
            continue

        if idx >= len(bar_rights) or idx + 1 >= len(bar_lefts):
            continue

        y_pos = value_to_y(current_value, axis_min, axis_max, plot_top, plot_height)
        base_start = bar_rights[idx]
        base_end = bar_lefts[idx + 1]
        x_start = base_start - overlap
        x_end = base_end + overlap
        add_connector(x_start, y_pos, x_end, y_pos)

        if connector_style == "step" and next_value is not None:
            next_y = value_to_y(next_value, axis_min, axis_max, plot_top, plot_height)
            if next_y != y_pos:
                add_connector(base_end, y_pos, base_end, next_y)
