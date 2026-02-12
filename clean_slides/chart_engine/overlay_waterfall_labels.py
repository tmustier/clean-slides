"""Label rendering helpers for waterfall overlays."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from typing import Any, Callable, Union, cast

from pptx.enum.text import PP_ALIGN

from . import annotations as _annotations
from . import overlay_waterfall_data_labels as _waterfall_data_labels
from .defaults import (
    DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT,
    DEFAULT_WATERFALL_LABEL_MARGIN,
    DEFAULT_WATERFALL_SERIES_LABEL_INSET,
    DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT,
)
from .geometry import value_to_x, value_to_y
from .spec_utils import format_label
from .units import emu_or_default

Number = Union[int, float]
FloatOrNone = Union[float, None]
LabelSpec = dict[str, Union[float, int, str]]
Geometry = Mapping[str, object]
MetaSpec = Mapping[str, object]

AddTextLabelFn = Callable[..., object]
MeasureLabelWidthFn = Callable[[str], int]


def _require_attr(module: object, name: str) -> object:
    value = getattr(module, name, None)
    if value is None:
        raise AttributeError(f"{module!r} does not expose {name}")
    return value


add_text_label = cast(AddTextLabelFn, _require_attr(_annotations, "add_text_label"))
measure_label_width = cast(
    MeasureLabelWidthFn,
    _require_attr(_waterfall_data_labels, "measure_label_width"),
)


def _number(value: object, default: Number = 0) -> Number:
    return value if isinstance(value, (int, float)) else default


def _float_or_none(value: object) -> FloatOrNone:
    if isinstance(value, (int, float)):
        return float(value)
    return None


def _float_list(value: object) -> list[float]:
    if not isinstance(value, list):
        return []

    typed_values = cast(list[object], value)
    result: list[float] = []
    for item in typed_values:
        if isinstance(item, (int, float)):
            result.append(float(item))
    return result


def _value_at(values: Sequence[FloatOrNone], idx: int) -> FloatOrNone:
    if idx < 0 or idx >= len(values):
        return None
    value = values[idx]
    if isinstance(value, (int, float)):
        return float(value)
    return None


def _optional_float_list(value: object) -> list[FloatOrNone]:
    if not isinstance(value, list):
        return []

    typed_values = cast(list[object], value)
    result: list[FloatOrNone] = []
    for item in typed_values:
        result.append(_float_or_none(item))
    return result


def _mapping(value: object) -> dict[object, object]:
    if not isinstance(value, dict):
        return {}
    return cast(dict[object, object], value)


def build_waterfall_value_label_specs(
    meta: MetaSpec,
    categories: Sequence[object],
    cumulative_totals: Sequence[FloatOrNone],
    delta_values: Sequence[FloatOrNone],
    total_categories: set[int],
    label_tops: Sequence[FloatOrNone],
    label_bottoms: Sequence[FloatOrNone],
    chart_box: tuple[int, int, int, int],
    geometry: Geometry,
    orientation: str,
    axis_min: float,
    axis_max: float,
    plot_left: float,
    plot_top: float,
    plot_width: float,
    plot_height: float,
    label_gap: int,
    label_offset: int,
    slide_width: Union[int, None],
) -> list[LabelSpec]:
    """Build value-label box specs for waterfall overlays."""
    label_specs: list[LabelSpec] = []

    bar_centers = _float_list(geometry.get("bar_centers"))
    label_decimals = int(_number(meta.get("label_decimals", 0), 0))

    for idx, _category in enumerate(categories):
        total_value = _value_at(cumulative_totals, idx)
        if total_value is None:
            continue

        if idx == 0 or idx in total_categories:
            label_value = total_value
        else:
            label_value = _value_at(delta_values, idx)

        text = format_label(label_value, decimals=label_decimals)
        if text is None:
            continue

        if idx >= len(bar_centers):
            continue

        label_width = measure_label_width(text)
        label_height = int(DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT)

        top_value = _value_at(label_tops, idx)
        if top_value is None:
            top_value = total_value

        bottom_value = _value_at(label_bottoms, idx)
        if bottom_value is None:
            bottom_value = top_value

        if orientation == "horizontal":
            anchor_value = top_value
            if label_value is not None and label_value < 0:
                anchor_value = bottom_value

            x_base = value_to_x(anchor_value, axis_min, axis_max, plot_left, plot_width)
            if label_value is not None and label_value < 0:
                x = x_base - label_width - label_offset
            else:
                x = x_base + label_offset

            min_x = chart_box[0] - label_width
            max_x = (
                slide_width - label_width
                if slide_width is not None
                else chart_box[0] + chart_box[2] + label_offset - label_width
            )
            if x < min_x:
                x = float(min_x)
            if x > max_x:
                x = float(max_x)

            y = bar_centers[idx] - label_height / 2
            anchor = "middle"
        else:
            x = bar_centers[idx] - label_width / 2
            if label_value is not None and label_value < 0:
                y_base = value_to_y(bottom_value, axis_min, axis_max, plot_top, plot_height)
                y = y_base + label_gap
                anchor = "top"
            else:
                y_base = value_to_y(top_value, axis_min, axis_max, plot_top, plot_height)
                y = y_base - label_gap - label_height
                anchor = "bottom"

        label_specs.append(
            {
                "text": text,
                "x": float(x),
                "y": float(y),
                "width": int(label_width),
                "height": int(label_height),
                "vertical_anchor": anchor,
            }
        )

    return label_specs


def add_waterfall_value_labels(slide: Any, label_specs: Sequence[LabelSpec]) -> None:
    """Render prepared value-label specs."""
    for spec in label_specs:
        text = spec.get("text")
        x = spec.get("x")
        y = spec.get("y")
        width = spec.get("width")
        height = spec.get("height")
        vertical_anchor = spec.get("vertical_anchor")

        if not isinstance(text, str):
            continue
        if not isinstance(x, (int, float)):
            continue
        if not isinstance(y, (int, float)):
            continue
        if not isinstance(width, (int, float)):
            continue
        if not isinstance(height, (int, float)):
            continue

        add_text_label(
            slide,
            text,
            x,
            y,
            width,
            height,
            margin_left=DEFAULT_WATERFALL_LABEL_MARGIN,
            margin_right=DEFAULT_WATERFALL_LABEL_MARGIN,
            vertical_anchor=vertical_anchor if isinstance(vertical_anchor, str) else None,
        )


def add_waterfall_category_labels(
    slide: Any,
    categories: Sequence[object],
    chart_box: tuple[int, int, int, int],
    geometry: Geometry,
    orientation: str,
    category_offset: Number,
    plot_left: float,
    plot_bottom: float,
) -> None:
    """Render category labels for waterfall overlays."""
    bar_centers = _float_list(geometry.get("bar_centers"))

    if orientation == "horizontal":
        for idx, label in enumerate(categories):
            if idx >= len(bar_centers):
                continue

            text = str(label)
            label_width = measure_label_width(text)
            x = plot_left - category_offset - label_width
            min_x = max(0, chart_box[0] - label_width)
            if x < min_x:
                x = float(min_x)

            y = bar_centers[idx] - int(DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT) / 2
            add_text_label(
                slide,
                text,
                x,
                y,
                label_width,
                int(DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT),
                align=PP_ALIGN.RIGHT,
            )
        return

    slot_width_obj = geometry.get("slot_width")
    if not isinstance(slot_width_obj, (int, float)):
        return

    slot_width = float(slot_width_obj)
    category_y = plot_bottom + category_offset

    plot_left_obj = geometry.get("plot_left", plot_left)
    plot_left_value = plot_left_obj if isinstance(plot_left_obj, (int, float)) else plot_left

    for idx, label in enumerate(categories):
        x = plot_left_value + slot_width * idx
        add_text_label(
            slide,
            str(label),
            x,
            category_y,
            slot_width,
            int(DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT),
        )


def add_waterfall_series_labels(
    slide: Any,
    chart_box: tuple[int, int, int, int],
    chart_series_names: Sequence[object],
    segment_values: Mapping[object, object],
    orientation: str,
    meta: MetaSpec,
    axis_min: float,
    axis_max: float,
    plot_top: float,
    plot_height: float,
) -> None:
    """Render stacked-series labels shown left of vertical waterfall charts."""
    if not chart_series_names or orientation != "vertical":
        return

    series_label_inset = emu_or_default(
        meta.get("series_label_inset"),
        int(DEFAULT_WATERFALL_SERIES_LABEL_INSET),
    )
    series_label_left = emu_or_default(meta.get("series_label_left"), 0)

    segment_values_dict = _mapping(segment_values)

    stacked: list[tuple[str, float]] = []
    for name in chart_series_names:
        values_obj = segment_values_dict.get(name, [])
        values = _optional_float_list(values_obj)
        first_value = values[0] if values else None
        magnitude = abs(first_value) if first_value is not None else 0.0
        stacked.append((str(name), magnitude))

    current = 0.0
    for name, magnitude in stacked:
        if magnitude <= 0:
            continue

        center_val = current + magnitude / 2
        current += magnitude
        y = value_to_y(center_val, axis_min, axis_max, plot_top, plot_height)

        text = str(name)
        width = measure_label_width(text)
        label_right = chart_box[0] + series_label_inset
        x = max(label_right - width, series_label_left)
        add_text_label(
            slide,
            text,
            x,
            y - int(DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT) / 2,
            width,
            int(DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT),
            align=PP_ALIGN.RIGHT,
        )
