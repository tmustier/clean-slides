"""Waterfall overlay orchestration."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any, Union, cast

from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt

from .defaults import (
    DEFAULT_WATERFALL_CATEGORY_OFFSET_RATIO,
    DEFAULT_WATERFALL_CONNECTOR_DASH_GAP,
    DEFAULT_WATERFALL_CONNECTOR_DASH_LENGTH,
    DEFAULT_WATERFALL_CONNECTOR_DOT_GAP,
    DEFAULT_WATERFALL_CONNECTOR_DOT_LENGTH,
    DEFAULT_WATERFALL_CONNECTOR_INSET,
    DEFAULT_WATERFALL_CONNECTOR_OVERLAP,
    DEFAULT_WATERFALL_LABEL_GAP,
    DEFAULT_WATERFALL_LABEL_OFFSET_RATIO,
    DEFAULT_WATERFALL_PLOT_LAYOUT,
)
from .geometry import compute_category_geometry, normalize_orientation, resolve_label_collisions
from .overlay_waterfall_connectors import render_waterfall_connectors
from .overlay_waterfall_labels import (
    add_waterfall_category_labels,
    add_waterfall_series_labels,
    add_waterfall_value_labels,
    build_waterfall_value_label_specs,
)
from .units import coerce_emu, coerce_line_width, coerce_offset_value

Number = Union[int, float]
FloatOrNone = Union[float, None]
ColorValue = Union[RGBColor, str, None]
SlideSize = Union[tuple[int, int], None]


def _number(value: object, default: Number = 0) -> Number:
    return value if isinstance(value, (int, float)) else default


def _float(value: object, default: float = 0.0) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    return default


def _str(value: object, default: str) -> str:
    return value.strip().lower() if isinstance(value, str) else default


def _mapping(value: object) -> dict[str, object]:
    if not isinstance(value, dict):
        return {}

    typed_mapping = cast(dict[object, object], value)
    mapped: dict[str, object] = {}
    for key, item in typed_mapping.items():
        if isinstance(key, str):
            mapped[key] = item
    return mapped


def _list(value: object) -> list[object]:
    if isinstance(value, list):
        return cast(list[object], value)
    if isinstance(value, tuple):
        return list(cast(tuple[object, ...], value))
    return []


def _float_or_none(value: object) -> FloatOrNone:
    if isinstance(value, (int, float)):
        return float(value)
    return None


def _float_or_none_list(value: object) -> list[FloatOrNone]:
    result: list[FloatOrNone] = []
    for item in _list(value):
        result.append(_float_or_none(item))
    return result


def _color(value: object) -> ColorValue:
    if isinstance(value, RGBColor):
        return value
    if isinstance(value, str):
        return value
    return None


def _index_set(value: object) -> set[int]:
    indices: set[int] = set()
    items: list[object]
    if isinstance(value, set):
        items = list(cast(set[object], value))
    elif isinstance(value, list):
        items = cast(list[object], value)
    elif isinstance(value, tuple):
        items = list(cast(tuple[object, ...], value))
    else:
        return indices

    for item in items:
        if isinstance(item, int):
            indices.add(item)
    return indices


def add_waterfall_overlays(
    slide: Any,
    chart_box: tuple[int, int, int, int],
    meta: Mapping[str, object],
    slide_size: SlideSize = None,
) -> None:
    overlay_raw = meta.get("overlay") if meta else None
    overlay = _mapping(overlay_raw)
    if not overlay:
        return

    categories = _list(overlay.get("categories"))
    cumulative_totals = _float_or_none_list(overlay.get("cumulative_totals"))
    delta_values = _float_or_none_list(overlay.get("delta_values"))
    label_tops = _float_or_none_list(overlay.get("label_tops"))
    label_bottoms = _float_or_none_list(overlay.get("label_bottoms"))
    total_categories = _index_set(overlay.get("total_categories"))
    chart_series_names = _list(overlay.get("chart_series"))

    segment_values_raw = overlay.get("segment_values")
    segment_values: Mapping[object, object]
    if isinstance(segment_values_raw, dict):
        segment_values = cast(dict[object, object], segment_values_raw)
    else:
        segment_values = {}

    axis_min = _float(meta.get("axis_min", 0), 0.0)
    axis_max = _float(meta.get("axis_max", 0), 0.0)
    gap_width = int(_number(meta.get("gap_width", 80), 80))

    plot_layout_raw = meta.get("plot_layout")
    if isinstance(plot_layout_raw, dict):
        typed_plot_layout = cast(dict[object, object], plot_layout_raw)
        plot_layout = {
            key: float(value)
            for key, value in typed_plot_layout.items()
            if isinstance(key, str) and isinstance(value, (int, float))
        }
        if not plot_layout:
            plot_layout = DEFAULT_WATERFALL_PLOT_LAYOUT
    else:
        plot_layout = DEFAULT_WATERFALL_PLOT_LAYOUT

    orientation = normalize_orientation(_str(meta.get("orientation"), "vertical"))

    geometry = compute_category_geometry(
        chart_box,
        plot_layout,
        categories,
        gap_width,
        orientation,
    )

    plot_top = _float(geometry.get("plot_top"), 0.0)
    plot_height = _float(geometry.get("plot_height"), 0.0)
    plot_left = _float(geometry.get("plot_left"), 0.0)
    plot_width = _float(geometry.get("plot_width"), 0.0)
    plot_bottom = plot_top + plot_height

    slide_width = slide_size[0] if slide_size else None

    axis_span = chart_box[3] if orientation == "vertical" else chart_box[2]
    label_gap_raw = meta.get("label_gap")
    if label_gap_raw is None:
        label_gap_raw = meta.get("label_offset")

    if label_gap_raw is not None:
        label_gap_value = coerce_emu(label_gap_raw) or 0
    else:
        ratio = meta.get("label_offset_ratio")
        if isinstance(ratio, (int, float)):
            label_gap_value = axis_span * float(ratio)
        elif orientation == "vertical":
            label_gap_value = int(DEFAULT_WATERFALL_LABEL_GAP)
        else:
            label_gap_value = axis_span * float(DEFAULT_WATERFALL_LABEL_OFFSET_RATIO)

    label_gap = int(label_gap_value)
    if orientation == "horizontal":
        label_offset = max(label_gap, 0)
    else:
        label_gap = max(label_gap, 0)
        label_offset = label_gap

    connector_style = _str(meta.get("connector_style"), "gap")
    connector_value_mode = _str(meta.get("connector_value"), "totals")
    if connector_value_mode in {"totals", "total", "running", "end"}:
        connector_values = cumulative_totals
    else:
        connector_values = label_tops

    connector_width = meta.get("connector_line_width")
    connector_color = _color(meta.get("connector_line_color"))
    connector_dash = _str(meta.get("connector_dash_style"), "long_dash")

    connector_overlap_raw = meta.get("connector_overlap")
    if connector_overlap_raw is None:
        connector_overlap = int(DEFAULT_WATERFALL_CONNECTOR_OVERLAP)
    else:
        connector_overlap = coerce_emu(connector_overlap_raw) or 0

    connector_inset_raw = meta.get("connector_inset")
    if connector_inset_raw is None:
        connector_inset = (
            int(DEFAULT_WATERFALL_CONNECTOR_INSET) if orientation == "horizontal" else 0
        )
    else:
        connector_inset = coerce_emu(connector_inset_raw) or 0

    if connector_width is not None:
        line_width_value = coerce_line_width(connector_width)
    else:
        line_width_value = int(Pt(0.25))
    if line_width_value is None:
        line_width_value = int(Pt(0.25))
    line_width = max(int(line_width_value), int(Emu(6000)))

    if connector_dash == "solid":
        dash_length = None
        dash_gap = 0
    elif connector_dash == "dot":
        dash_length = int(DEFAULT_WATERFALL_CONNECTOR_DOT_LENGTH)
        dash_gap = int(DEFAULT_WATERFALL_CONNECTOR_DOT_GAP)
    else:
        dash_length = int(DEFAULT_WATERFALL_CONNECTOR_DASH_LENGTH)
        dash_gap = int(DEFAULT_WATERFALL_CONNECTOR_DASH_GAP)

    render_waterfall_connectors(
        slide,
        categories=categories,
        connector_values=connector_values,
        geometry=geometry,
        orientation=orientation,
        axis_min=axis_min,
        axis_max=axis_max,
        plot_left=plot_left,
        plot_top=plot_top,
        plot_width=plot_width,
        plot_height=plot_height,
        connector_style=connector_style,
        connector_inset=connector_inset,
        connector_overlap=connector_overlap,
        line_width=line_width,
        dash_length=dash_length,
        dash_gap=dash_gap,
        connector_color=connector_color,
    )

    label_specs = build_waterfall_value_label_specs(
        meta,
        categories=categories,
        cumulative_totals=cumulative_totals,
        delta_values=delta_values,
        total_categories=total_categories,
        label_tops=label_tops,
        label_bottoms=label_bottoms,
        chart_box=chart_box,
        geometry=geometry,
        orientation=orientation,
        axis_min=axis_min,
        axis_max=axis_max,
        plot_left=plot_left,
        plot_top=plot_top,
        plot_width=plot_width,
        plot_height=plot_height,
        label_gap=label_gap,
        label_offset=label_offset,
        slide_width=slide_width,
    )

    if meta.get("label_collision") and label_specs:
        gap = coerce_offset_value(meta.get("label_collision_gap"))
        if orientation == "horizontal":
            resolve_label_collisions(label_specs, axis="x", min_gap=gap, direction=1)
        else:
            resolve_label_collisions(label_specs, axis="y", min_gap=gap, direction=-1)

    add_waterfall_value_labels(slide, label_specs)

    category_offset_raw = meta.get("category_label_offset")
    if category_offset_raw is None:
        category_offset_raw = meta.get("category_offset")

    if category_offset_raw is not None:
        category_offset = coerce_emu(category_offset_raw) or 0
    else:
        category_offset = axis_span * DEFAULT_WATERFALL_CATEGORY_OFFSET_RATIO

    add_waterfall_category_labels(
        slide,
        categories=categories,
        chart_box=chart_box,
        geometry=geometry,
        orientation=orientation,
        category_offset=category_offset,
        plot_left=plot_left,
        plot_bottom=plot_bottom,
    )

    add_waterfall_series_labels(
        slide,
        chart_box=chart_box,
        chart_series_names=chart_series_names,
        segment_values=segment_values,
        orientation=orientation,
        meta=meta,
        axis_min=axis_min,
        axis_max=axis_max,
        plot_top=plot_top,
        plot_height=plot_height,
    )
