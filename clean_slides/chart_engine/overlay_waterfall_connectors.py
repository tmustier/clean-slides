"""Connector rendering helpers for waterfall overlays."""

# pyright: reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false
# pyright: reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false
# pyright: reportMissingParameterType=false
# pyright: reportMissingTypeArgument=false
# pyright: reportGeneralTypeIssues=false
# pyright: reportArgumentType=false
# pyright: reportCallIssue=false
# pyright: reportAttributeAccessIssue=false
# pyright: reportIndexIssue=false
# pyright: reportOperatorIssue=false

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from .colors import apply_color
from .geometry import value_to_x, value_to_y


def render_waterfall_connectors(
    slide,
    categories: list[object],
    connector_values: list[float | None],
    geometry: dict,
    orientation: str,
    axis_min: float,
    axis_max: float,
    plot_left: float,
    plot_top: float,
    plot_width: float,
    plot_height: float,
    connector_style: str,
    connector_inset: int,
    connector_overlap: int,
    line_width: int,
    dash_length: int | None,
    dash_gap: int,
    connector_color,
) -> None:
    """Render connector segments between adjacent waterfall categories."""

    def add_dash_segment(x: float, y: float, width: float, height: float) -> None:
        if width <= 0 or height <= 0:
            return
        shape = slide.shapes.add_shape(
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
                seg_len = min(dash_length, total - pos)
                add_dash_segment(x, start + pos, line_width, seg_len)
                pos += seg_len + dash_gap
        else:
            y = min(y1, y2) - line_width / 2
            start = min(x1, x2)
            total = abs(x2 - x1)
            if dash_length is None or dash_length <= 0:
                add_dash_segment(start, y, total, line_width)
                return
            pos = 0.0
            while pos < total:
                seg_len = min(dash_length, total - pos)
                add_dash_segment(start + pos, y, seg_len, line_width)
                pos += seg_len + dash_gap

    bar_bottoms = geometry.get("bar_bottoms") or []
    bar_tops = geometry.get("bar_tops") or []
    bar_rights = geometry.get("bar_rights") or []
    bar_lefts = geometry.get("bar_lefts") or []

    for idx in range(len(categories) - 1):
        if idx >= len(connector_values):
            continue
        current_value = connector_values[idx]
        if current_value is None:
            continue
        next_value = connector_values[idx + 1] if idx + 1 < len(connector_values) else None

        if orientation == "horizontal":
            if idx >= len(bar_bottoms) or idx + 1 >= len(bar_tops):
                continue
            x_pos = value_to_x(current_value, axis_min, axis_max, plot_left, plot_width)
            x_pos -= connector_inset
            base_start = bar_bottoms[idx]
            base_end = bar_tops[idx + 1]
            y_start = base_start - connector_overlap
            y_end = base_end + connector_overlap
            add_connector(x_pos, y_start, x_pos, y_end)
            if connector_style == "step" and next_value is not None:
                next_x = value_to_x(next_value, axis_min, axis_max, plot_left, plot_width)
                next_x -= connector_inset
                if next_x != x_pos:
                    add_connector(x_pos, base_end, next_x, base_end)
        else:
            if idx >= len(bar_rights) or idx + 1 >= len(bar_lefts):
                continue
            y_pos = value_to_y(current_value, axis_min, axis_max, plot_top, plot_height)
            base_start = bar_rights[idx]
            base_end = bar_lefts[idx + 1]
            x_start = base_start - connector_overlap
            x_end = base_end + connector_overlap
            add_connector(x_start, y_pos, x_end, y_pos)
            if connector_style == "step" and next_value is not None:
                next_y = value_to_y(next_value, axis_min, axis_max, plot_top, plot_height)
                if next_y != y_pos:
                    add_connector(base_end, y_pos, base_end, next_y)
