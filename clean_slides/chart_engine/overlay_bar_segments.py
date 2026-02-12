"""Segment label overlay helpers for bar charts."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, Callable, Union, cast

from pptx.dml.color import RGBColor

from . import annotations as _annotations
from . import text_templates as _text_templates
from .defaults import (
    DEFAULT_BAR_SEGMENT_LABEL_FONT_SIZE,
    DEFAULT_BAR_SEGMENT_LABEL_HEIGHT,
    DEFAULT_BAR_SEGMENT_LABEL_OFFSET_RATIO,
    DEFAULT_BAR_SEGMENT_LABEL_WIDTH,
)
from .geometry import value_to_x, value_to_y
from .spec_utils import (
    format_label,
    normalize_category_indices,
    normalize_list,
    numeric_value,
    safe_value,
)
from .units import coerce_offset_value, emu_or_default, normalize_offset_matrix

Number = Union[int, float]
NumberOrBool = Union[int, float, bool]
StrOrNone = Union[str, None]
PathOrNone = Union[Path, None]
ColorValue = Union[RGBColor, str, None]
OverlaySpec = Mapping[str, object]
Geometry = Mapping[str, object]
TemplateMap = Mapping[str, object]

AddTextLabelFn = Callable[..., object]
ResolveTemplateFn = Callable[[PathOrNone, str, object], object]


def _require_attr(module: object, name: str) -> object:
    value = getattr(module, name, None)
    if value is None:
        raise AttributeError(f"{module!r} does not expose {name}")
    return value


add_text_label = cast(AddTextLabelFn, _require_attr(_annotations, "add_text_label"))
resolve_txbody_template = cast(
    ResolveTemplateFn,
    _require_attr(_text_templates, "resolve_txbody_template"),
)


def _mapping(value: object) -> dict[str, object]:
    if not isinstance(value, dict):
        return {}

    typed_mapping = cast(dict[object, object], value)
    mapped: dict[str, object] = {}
    for key, item in typed_mapping.items():
        if isinstance(key, str):
            mapped[key] = item
    return mapped


def _number(value: object, default: Number = 0) -> Number:
    return value if isinstance(value, (int, float)) else default


def _int(value: object, default: int = 0) -> int:
    return int(_number(value, default))


def _bool(value: object, default: bool) -> bool:
    return value if isinstance(value, bool) else default


def _geometry_float(geometry: Geometry, key: str) -> float:
    value = geometry.get(key)
    if isinstance(value, (int, float)):
        return float(value)
    return 0.0


def _geometry_series(geometry: Geometry, key: str) -> list[float]:
    values = geometry.get(key)
    if not isinstance(values, list):
        return []

    typed_values = cast(list[object], values)
    result: list[float] = []
    for item in typed_values:
        if isinstance(item, (int, float)):
            result.append(float(item))
    return result


def _coerce_fill_color(fill_value: object, series_color: StrOrNone) -> ColorValue:
    if fill_value == "series":
        return series_color
    if fill_value in {None, "none", "transparent", False}:
        return None
    if isinstance(fill_value, (RGBColor, str)):
        return fill_value
    return None


def add_bar_segment_labels(
    slide: Any,
    *,
    overlay: OverlaySpec,
    categories: Sequence[object],
    series_names: Sequence[str],
    series_colors: Sequence[StrOrNone],
    segment_values: Sequence[Sequence[object]],
    orientation: str,
    axis_min: float,
    axis_max: float,
    plot_left: float,
    plot_top: float,
    plot_width: float,
    plot_height: float,
    geometry: Geometry,
    template_path: PathOrNone,
    templates: TemplateMap,
) -> None:
    segment_configs = normalize_list(overlay.get("segment_labels"))
    categories_list = list(categories)

    bar_tops = _geometry_series(geometry, "bar_tops")
    bar_height = _geometry_float(geometry, "bar_height")
    bar_lefts = _geometry_series(geometry, "bar_lefts")
    bar_width = _geometry_float(geometry, "bar_width")

    for raw_segment_cfg in segment_configs:
        segment_cfg = _mapping(raw_segment_cfg)
        if not segment_cfg:
            continue
        if not _bool(segment_cfg.get("show", True), True):
            continue

        segment_series = normalize_list(
            segment_cfg.get("series_indices")
            or segment_cfg.get("series")
            or segment_cfg.get("series_names")
        )

        segment_series_indices: list[int] = []
        for item in segment_series:
            if isinstance(item, int):
                segment_series_indices.append(item)
            elif isinstance(item, str) and item in series_names:
                segment_series_indices.append(series_names.index(item))

        segment_series_indices = [
            idx for idx in segment_series_indices if 0 <= idx < len(segment_values)
        ]
        if not segment_series_indices:
            continue

        category_filter = normalize_category_indices(
            categories_list,
            segment_cfg.get("categories") or segment_cfg.get("category_indices"),
        )
        category_filter_set: Union[set[int], None]
        category_filter_set = set(category_filter) if category_filter else None

        positions = segment_cfg.get("positions")
        if positions is not None:
            ratios = [float(cast(NumberOrBool, val)) for val in normalize_list(positions)]
        else:
            offset_ratio = float(
                _number(segment_cfg.get("offset_ratio"), DEFAULT_BAR_SEGMENT_LABEL_OFFSET_RATIO)
            )
            if len(segment_series_indices) == 1:
                ratios = [0.5]
            else:
                start = 0.5 - offset_ratio
                end = 0.5 + offset_ratio
                step = (end - start) / (len(segment_series_indices) - 1)
                ratios = [start + step * idx for idx in range(len(segment_series_indices))]

        ratio_map = {
            series_idx: ratios[idx]
            for idx, series_idx in enumerate(segment_series_indices)
            if idx < len(ratios)
        }

        segment_width = emu_or_default(
            segment_cfg.get("width"),
            int(DEFAULT_BAR_SEGMENT_LABEL_WIDTH),
        )
        segment_height = emu_or_default(
            segment_cfg.get("height"),
            int(DEFAULT_BAR_SEGMENT_LABEL_HEIGHT),
        )
        segment_font = segment_cfg.get("font_size", DEFAULT_BAR_SEGMENT_LABEL_FONT_SIZE)
        segment_color = segment_cfg.get("text_color", "bg1")
        segment_fill = segment_cfg.get("fill", "series")
        segment_decimals = _int(segment_cfg.get("decimals", 0), 0)

        segment_offsets_x = [
            coerce_offset_value(val) for val in normalize_list(segment_cfg.get("offsets_x"))
        ]
        segment_offsets_y = [
            coerce_offset_value(val) for val in normalize_list(segment_cfg.get("offsets_y"))
        ]
        segment_offsets_x_by_category = normalize_offset_matrix(
            segment_cfg.get("offsets_x_by_category") or segment_cfg.get("offsets_x_matrix")
        )
        segment_offsets_y_by_category = normalize_offset_matrix(
            segment_cfg.get("offsets_y_by_category") or segment_cfg.get("offsets_y_matrix")
        )

        for cat_idx in range(len(categories_list)):
            if category_filter_set and cat_idx not in category_filter_set:
                continue

            running = 0.0
            for series_idx, series_vals in enumerate(segment_values):
                value = numeric_value(safe_value(series_vals, cat_idx))
                if value is None:
                    continue

                x_start = 0.0
                x_end = 0.0
                y_bottom = 0.0
                y_top = 0.0
                if orientation == "horizontal":
                    x_start = value_to_x(running, axis_min, axis_max, plot_left, plot_width)
                    x_end = value_to_x(running + value, axis_min, axis_max, plot_left, plot_width)
                else:
                    y_bottom = value_to_y(running, axis_min, axis_max, plot_top, plot_height)
                    y_top = value_to_y(running + value, axis_min, axis_max, plot_top, plot_height)

                if series_idx in ratio_map:
                    text = format_label(value, decimals=segment_decimals)
                    if text is not None:
                        ratio = ratio_map[series_idx]

                        if orientation == "horizontal":
                            if cat_idx >= len(bar_tops):
                                running += value
                                continue

                            x_center = (x_start + x_end) / 2
                            y_center = bar_tops[cat_idx] + bar_height * ratio
                        else:
                            if cat_idx >= len(bar_lefts):
                                running += value
                                continue

                            x_center = bar_lefts[cat_idx] + bar_width * ratio
                            y_center = (y_bottom + y_top) / 2

                        series_color = (
                            series_colors[series_idx] if series_idx < len(series_colors) else None
                        )
                        fill_color = _coerce_fill_color(segment_fill, series_color)

                        offset_idx = segment_series_indices.index(series_idx)
                        offset_x = (
                            segment_offsets_x[offset_idx]
                            if offset_idx < len(segment_offsets_x)
                            else 0
                        )
                        offset_y = (
                            segment_offsets_y[offset_idx]
                            if offset_idx < len(segment_offsets_y)
                            else 0
                        )

                        if segment_offsets_x_by_category and cat_idx < len(
                            segment_offsets_x_by_category
                        ):
                            row = segment_offsets_x_by_category[cat_idx]
                            if offset_idx < len(row):
                                offset_x += row[offset_idx]

                        if segment_offsets_y_by_category and cat_idx < len(
                            segment_offsets_y_by_category
                        ):
                            row = segment_offsets_y_by_category[cat_idx]
                            if offset_idx < len(row):
                                offset_y += row[offset_idx]

                        label = add_text_label(
                            slide,
                            text,
                            x_center - segment_width / 2 + offset_x,
                            y_center - segment_height / 2 + offset_y,
                            segment_width,
                            segment_height,
                            font_size=segment_font,
                            color=segment_color,
                            fill_color=fill_color,
                            margin_left=segment_cfg.get("margin_left", 25400),
                            margin_right=segment_cfg.get("margin_right", 25400),
                            margin_top=segment_cfg.get("margin_top", 0),
                            margin_bottom=segment_cfg.get("margin_bottom", 0),
                            vertical_anchor=segment_cfg.get("vertical_anchor", "center"),
                            bold=segment_cfg.get("bold"),
                            bw_mode=segment_cfg.get("bw_mode", "gray"),
                            txbody_template=resolve_txbody_template(
                                template_path,
                                text,
                                templates.get("segment"),
                            ),
                        )
                        if label is not None:
                            rendered_width = getattr(label, "width", None)
                            if isinstance(rendered_width, int):
                                segment_width = rendered_width

                running += value
