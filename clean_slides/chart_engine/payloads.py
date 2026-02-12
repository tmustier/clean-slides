"""Chart payload builders for bar and waterfall chart specs."""

from __future__ import annotations

import math
from collections.abc import Mapping
from pathlib import Path
from typing import Any, Union, cast

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE

from .defaults import (
    DEFAULT_BAR_AXIS_PADDING,
    DEFAULT_BAR_CATEGORY_LABEL_FONT_SIZE,
    DEFAULT_BAR_CATEGORY_LABEL_HEIGHT,
    DEFAULT_BAR_CATEGORY_LABEL_OFFSET,
    DEFAULT_BAR_CATEGORY_LABEL_WIDTH,
    DEFAULT_BAR_LEGEND_LABEL_FONT_SIZE,
    DEFAULT_BAR_LEGEND_LABEL_HEIGHT,
    DEFAULT_BAR_LEGEND_LABEL_OFFSET,
    DEFAULT_BAR_LEGEND_LABEL_WIDTH,
    DEFAULT_BAR_LEGEND_LEFT_RATIO,
    DEFAULT_BAR_LEGEND_MARKER_HEIGHT,
    DEFAULT_BAR_LEGEND_MARKER_LEFT_RATIO,
    DEFAULT_BAR_LEGEND_MARKER_STEP_RATIO,
    DEFAULT_BAR_LEGEND_MARKER_WIDTH,
    DEFAULT_BAR_LEGEND_MARKER_Y_OFFSET,
    DEFAULT_BAR_LEGEND_STEP_RATIO,
    DEFAULT_BAR_OVERLAY_BAND_EXTRA,
    DEFAULT_BAR_PLOT_LAYOUT,
    DEFAULT_BAR_SERIES_BORDER_COLOR,
    DEFAULT_BAR_TOTAL_LABEL_FONT_SIZE,
    DEFAULT_BAR_TOTAL_LABEL_HEIGHT,
    DEFAULT_BAR_TOTAL_LABEL_OFFSET,
    DEFAULT_BAR_TOTAL_LABEL_WIDTH,
    DEFAULT_WATERFALL_PLOT_LAYOUT,
    DEFAULT_WATERFALL_START_COLOR,
    DEFAULT_WATERFALL_TOTAL_COLOR,
)
from .geometry import normalize_orientation
from .spec_utils import (
    infer_label_decimals,
    infer_total_categories,
    is_range_series,
    is_total_series,
    normalize_category_set,
    normalize_series_set,
    normalize_total_series,
    numeric_value,
    safe_value,
    sum_numeric,
)
from .units import emu_or_default, resolve_path

Number = Union[int, float]
SeriesEntry = dict[str, object]
SpecMap = Mapping[str, object]
BaseDirValue = Union[str, Path, None]

BAR_TUNING_KEYS = {
    "totals": "total_label_offset",
    "legend": "legend_label_offset",
    "category": "category_label_offset",
}


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


def _series_entries(value: object) -> list[SeriesEntry]:
    entries: list[SeriesEntry] = []
    for item in _list(value):
        mapped = _mapping(item)
        if mapped:
            entries.append(mapped)
    return entries


def _str(value: object, default: str = "") -> str:
    return value if isinstance(value, str) else default


def _str_or_none(value: object) -> str | None:
    return value if isinstance(value, str) else None


def _number(value: object, default: Number = 0) -> Number:
    return value if isinstance(value, (int, float)) else default


def _int(value: object, default: int = 0) -> int:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    return default


def _bool(value: object, default: bool) -> bool:
    return value if isinstance(value, bool) else default


def _color_str(value: object) -> str | None:
    return value if isinstance(value, str) else None


def _base_dir(value: object) -> BaseDirValue:
    if isinstance(value, (str, Path)):
        return value
    return None


def apply_bar_tuning(bar_cfg: Mapping[str, object]) -> dict[str, object]:
    if not bar_cfg:
        return {}

    tuning = _mapping(bar_cfg.get("tuning"))
    if not tuning:
        return dict(bar_cfg)

    updated = dict(bar_cfg)
    for source_key, target_key in BAR_TUNING_KEYS.items():
        if source_key in tuning and target_key not in updated:
            updated[target_key] = tuning.get(source_key)
    updated.pop("tuning", None)
    return updated


def build_bar_payload(spec: SpecMap) -> tuple[XL_CHART_TYPE, CategoryChartData, dict[str, object]]:
    categories = _list(spec.get("categories"))
    series = _series_entries(spec.get("series"))
    chart_type_key = _str(spec.get("type"), "clustered")
    bar_cfg = apply_bar_tuning(_mapping(spec.get("bar")))
    orientation = normalize_orientation(
        _str_or_none(bar_cfg.get("orientation")) or _str_or_none(spec.get("orientation"))
    )

    stacked = chart_type_key == "stacked"
    if orientation == "horizontal":
        chart_type = XL_CHART_TYPE.BAR_STACKED if stacked else XL_CHART_TYPE.BAR_CLUSTERED
    else:
        chart_type = XL_CHART_TYPE.COLUMN_STACKED if stacked else XL_CHART_TYPE.COLUMN_CLUSTERED

    chart_data = CategoryChartData()
    chart_data.categories = cast(Any, categories)

    for entry in series:
        name = _str(entry.get("name"), "Series")
        series_values = _list(entry.get("values"))
        chart_data.add_series(name, cast(Any, series_values))

    series_colors = [_color_str(entry.get("color")) for entry in series]

    totals: list[float | None] = []
    total_label_tops: list[float | None] = []
    min_val = 0.0
    max_val = 0.0

    for idx in range(len(categories)):
        values: list[float] = []
        for entry in series:
            series_values = _list(entry.get("values"))
            value = numeric_value(safe_value(series_values, idx))
            if value is not None:
                values.append(value)

        if not values:
            totals.append(None)
            total_label_tops.append(None)
            continue

        if stacked:
            sum_pos = sum(val for val in values if val > 0)
            sum_neg = sum(val for val in values if val < 0)
            total = sum_pos + sum_neg
            totals.append(total)
            total_label_tops.append(sum_pos if total >= 0 else sum_neg)
            min_val = min(min_val, sum_neg)
            max_val = max(max_val, sum_pos)
        else:
            val_min = min(values)
            val_max = max(values)
            min_val = min(min_val, val_min)
            max_val = max(max_val, val_max)
            totals.append(None)
            total_label_tops.append(None)

    axis_padding = float(_number(bar_cfg.get("axis_padding"), DEFAULT_BAR_AXIS_PADDING))
    axis_min = min_val if min_val < 0 else 0.0
    axis_max = max_val if max_val > 0 else 0.0
    if axis_max > 0:
        axis_max *= axis_padding
    if axis_min < 0:
        axis_min *= axis_padding
    if axis_max == axis_min:
        axis_max = axis_min + 1

    overlap_value = bar_cfg.get("overlap")
    if isinstance(overlap_value, (int, float)):
        overlap = int(overlap_value)
    else:
        overlap = 100 if stacked else 0

    base_dir = _base_dir(spec.get("_base_dir"))
    chart_template_value = bar_cfg.get("chart_template")
    chart_template: object
    if isinstance(chart_template_value, Path):
        chart_template = str(resolve_path(str(chart_template_value), base_dir))
    elif isinstance(chart_template_value, str):
        if chart_template_value.strip():
            chart_template = str(resolve_path(chart_template_value, base_dir))
        else:
            chart_template = chart_template_value
    else:
        chart_template = chart_template_value

    return (
        chart_type,
        chart_data,
        {
            "series_colors": series_colors,
            "bar": {
                "axis_min": axis_min,
                "axis_max": axis_max,
                "orientation": orientation,
                "axis_line_color": bar_cfg.get("axis_line_color"),
                "gap_width": _int(bar_cfg.get("gap_width"), 80),
                "overlap": overlap,
                "plot_layout": bar_cfg.get("plot_layout") or DEFAULT_BAR_PLOT_LAYOUT,
                "series_border_color": bar_cfg.get(
                    "series_border_color", DEFAULT_BAR_SERIES_BORDER_COLOR
                ),
                "chart_template": chart_template,
                "chart_template_slide": _int(bar_cfg.get("chart_template_slide"), 1),
                "chart_template_chart_index": _int(bar_cfg.get("chart_template_chart_index"), 0),
                "chart_template_series_index": _int(bar_cfg.get("chart_template_series_index"), 0),
                "chart_template_copy": _bool(bar_cfg.get("chart_template_copy"), False),
                "overlay": {
                    "categories": categories,
                    "series_names": [_str(entry.get("name"), "Series") for entry in series],
                    "series_colors": series_colors,
                    "segment_values": [_list(entry.get("values")) for entry in series],
                    "segment_labels": bar_cfg.get("segment_labels"),
                    "annotations": bar_cfg.get("annotations"),
                    "text_style_template": bar_cfg.get("text_style_template"),
                    "text_style_map": bar_cfg.get("text_style_map"),
                    "_base_dir": base_dir,
                    "totals": totals,
                    "total_label_tops": total_label_tops,
                    "show_totals": _bool(bar_cfg.get("show_totals"), stacked),
                    "show_legend_labels": _bool(bar_cfg.get("show_legend_labels"), True),
                    "category_label_width": emu_or_default(
                        bar_cfg.get("category_label_width"),
                        int(DEFAULT_BAR_CATEGORY_LABEL_WIDTH),
                    ),
                    "category_label_widths": bar_cfg.get("category_label_widths"),
                    "category_label_offset": emu_or_default(
                        bar_cfg.get("category_label_offset"),
                        int(DEFAULT_BAR_CATEGORY_LABEL_OFFSET),
                    ),
                    "category_label_offsets": bar_cfg.get("category_label_offsets"),
                    "legend_label_width": emu_or_default(
                        bar_cfg.get("legend_label_width"),
                        int(DEFAULT_BAR_LEGEND_LABEL_WIDTH),
                    ),
                    "legend_label_offset": emu_or_default(
                        bar_cfg.get("legend_label_offset"),
                        int(DEFAULT_BAR_LEGEND_LABEL_OFFSET),
                    ),
                    "total_label_width": emu_or_default(
                        bar_cfg.get("total_label_width"),
                        int(DEFAULT_BAR_TOTAL_LABEL_WIDTH),
                    ),
                    "total_label_widths": bar_cfg.get("total_label_widths"),
                    "total_label_offsets": bar_cfg.get("total_label_offsets"),
                    "total_label_offsets_x": bar_cfg.get("total_label_offsets_x"),
                    "category_label_height": emu_or_default(
                        bar_cfg.get("category_label_height"),
                        int(DEFAULT_BAR_CATEGORY_LABEL_HEIGHT),
                    ),
                    "category_label_heights": bar_cfg.get("category_label_heights"),
                    "legend_label_height": emu_or_default(
                        bar_cfg.get("legend_label_height"),
                        int(DEFAULT_BAR_LEGEND_LABEL_HEIGHT),
                    ),
                    "total_label_height": emu_or_default(
                        bar_cfg.get("total_label_height"),
                        int(DEFAULT_BAR_TOTAL_LABEL_HEIGHT),
                    ),
                    "category_label_color": bar_cfg.get("category_label_color", RGBColor(0, 0, 0)),
                    "legend_label_color": bar_cfg.get("legend_label_color", RGBColor(0, 0, 0)),
                    "total_label_color": bar_cfg.get("total_label_color", RGBColor(0, 0, 0)),
                    "category_label_font": bar_cfg.get(
                        "category_label_font_size", DEFAULT_BAR_CATEGORY_LABEL_FONT_SIZE
                    ),
                    "legend_label_font": bar_cfg.get(
                        "legend_label_font_size", DEFAULT_BAR_LEGEND_LABEL_FONT_SIZE
                    ),
                    "total_label_font": bar_cfg.get(
                        "total_label_font_size", DEFAULT_BAR_TOTAL_LABEL_FONT_SIZE
                    ),
                    "legend_left_ratio": _number(
                        bar_cfg.get("legend_left_ratio"),
                        DEFAULT_BAR_LEGEND_LEFT_RATIO,
                    ),
                    "legend_step_ratio": _number(
                        bar_cfg.get("legend_step_ratio"),
                        DEFAULT_BAR_LEGEND_STEP_RATIO,
                    ),
                    "legend_layout": bar_cfg.get("legend_layout"),
                    "legend_left_offset": emu_or_default(bar_cfg.get("legend_left_offset"), 0),
                    "legend_top_offset": emu_or_default(bar_cfg.get("legend_top_offset"), 0),
                    "legend_step": emu_or_default(bar_cfg.get("legend_step"), 0),
                    "legend_alignment": bar_cfg.get("legend_alignment"),
                    "legend_show_markers": _bool(bar_cfg.get("legend_show_markers"), True),
                    "legend_marker_left_ratio": _number(
                        bar_cfg.get("legend_marker_left_ratio"),
                        DEFAULT_BAR_LEGEND_MARKER_LEFT_RATIO,
                    ),
                    "legend_marker_step_ratio": _number(
                        bar_cfg.get("legend_marker_step_ratio"),
                        DEFAULT_BAR_LEGEND_MARKER_STEP_RATIO,
                    ),
                    "legend_marker_width": emu_or_default(
                        bar_cfg.get("legend_marker_width"),
                        int(DEFAULT_BAR_LEGEND_MARKER_WIDTH),
                    ),
                    "legend_marker_height": emu_or_default(
                        bar_cfg.get("legend_marker_height"),
                        int(DEFAULT_BAR_LEGEND_MARKER_HEIGHT),
                    ),
                    "legend_marker_y_offset": emu_or_default(
                        bar_cfg.get("legend_marker_y_offset"),
                        int(DEFAULT_BAR_LEGEND_MARKER_Y_OFFSET),
                    ),
                    "total_label_offset": emu_or_default(
                        bar_cfg.get("total_label_offset"),
                        int(DEFAULT_BAR_TOTAL_LABEL_OFFSET),
                    ),
                    "overlay_band_extra": emu_or_default(
                        bar_cfg.get("overlay_band_extra"),
                        int(DEFAULT_BAR_OVERLAY_BAND_EXTRA),
                    ),
                },
            },
        },
    )


def build_waterfall_payload(
    spec: SpecMap,
) -> tuple[XL_CHART_TYPE, CategoryChartData, dict[str, object]]:
    categories = _list(spec.get("categories"))
    series_entries = _series_entries(spec.get("series"))
    wf = _mapping(spec.get("waterfall"))
    orientation = normalize_orientation(
        _str_or_none(wf.get("orientation")) or _str_or_none(spec.get("orientation"))
    )

    funnel = _bool(wf.get("funnel"), False) or _str(spec.get("type")) == "waterfall-funnel"
    reuse_start_base = _bool(wf.get("reuse_start_base"), True)
    total_series_names = normalize_total_series(wf.get("total_series"))
    total_categories = normalize_category_set(categories, wf.get("total_categories"))
    decrease_categories = normalize_category_set(categories, wf.get("decrease_categories"))
    plot_layout = wf.get("plot_layout") or DEFAULT_WATERFALL_PLOT_LAYOUT
    allow_total_override = _bool(wf.get("total_override"), False)

    total_entry: SeriesEntry | None = None
    segment_entries: list[SeriesEntry] = []
    for entry in series_entries:
        if total_entry is None and is_total_series(entry, total_series_names):
            total_entry = entry
        else:
            segment_entries.append(entry)

    if not total_categories:
        total_categories = infer_total_categories(
            categories,
            cast(list[Mapping[str, object]], segment_entries),
            cast(Mapping[str, object] | None, total_entry),
        )

    range_series_names = normalize_series_set(
        cast(list[Mapping[str, object]], segment_entries),
        wf.get("range_series"),
    )
    range_series_names.update(
        _str(entry.get("name"), "Series")
        for entry in segment_entries
        if _str(entry.get("role")).lower() in {"range", "band"}
    )

    base_entries = [
        entry for entry in segment_entries if not is_range_series(entry, range_series_names)
    ]
    base_series_names = {_str(entry.get("name"), "Series") for entry in base_entries}

    chart_series = list(reversed(segment_entries))

    row0: list[float | None] = [None for _ in categories]
    row_values: dict[str, list[float | None]] = {
        _str(entry.get("name"), "Series"): [None for _ in categories] for entry in chart_series
    }

    start_color = _color_str(wf.get("start_color")) if "start_color" in wf else None

    if "total_color" in wf:
        total_color = _color_str(wf.get("total_color"))
    elif total_entry and _color_str(total_entry.get("color")):
        total_color = _color_str(total_entry.get("color"))
    elif segment_entries and _color_str(segment_entries[0].get("color")):
        total_color = _color_str(segment_entries[0].get("color"))
    else:
        total_color = DEFAULT_WATERFALL_TOTAL_COLOR

    def value_for(entry: SeriesEntry, idx: int) -> float | None:
        return numeric_value(safe_value(_list(entry.get("values")), idx))

    def signed_value(entry: SeriesEntry, idx: int) -> float | None:
        value = value_for(entry, idx)
        if value is None:
            return None
        if idx in decrease_categories:
            return -abs(value)
        return value

    def total_override(idx: int) -> float | None:
        if total_entry is not None:
            override = value_for(total_entry, idx)
            if override is not None:
                return override

        if allow_total_override:
            for entry in base_entries:
                override = value_for(entry, idx)
                if override is not None:
                    return override

        return None

    segment_values = {
        _str(entry.get("name"), "Series"): [
            signed_value(entry, idx) for idx in range(len(categories))
        ]
        for entry in segment_entries
    }

    value_samples = [
        value for values in segment_values.values() for value in values if value is not None
    ]

    label_decimals_raw = wf.get("label_decimals")
    if isinstance(label_decimals_raw, (int, float)):
        label_decimals = int(label_decimals_raw)
    else:
        label_decimals = infer_label_decimals(value_samples)

    data_label_decimals_raw = wf.get("data_label_decimals")
    if isinstance(data_label_decimals_raw, (int, float)):
        data_label_decimals = int(data_label_decimals_raw)
    else:
        data_label_decimals = label_decimals

    cumulative_totals: list[float | None] = [None for _ in categories]
    delta_values: list[float | None] = [None for _ in categories]
    label_tops: list[float | None] = [None for _ in categories]
    label_bottoms: list[float | None] = [None for _ in categories]

    funnel_max_total: float | None = None
    if funnel:
        max_total_raw = wf.get("max_total")
        if isinstance(max_total_raw, (int, float)):
            funnel_max_total = float(max_total_raw)
        else:
            totals: list[float] = []
            for idx in range(len(categories)):
                if idx in total_categories:
                    continue
                totals.append(sum_numeric(signed_value(entry, idx) for entry in base_entries))
            funnel_max_total = max(totals) if totals else 0.0

    prev_total = 0.0
    min_total = 0.0
    max_total = 0.0
    seen_normal = False
    offset_points: list[tuple[int, str]] = []
    offset_label_indices: list[int] = []

    for idx in range(len(categories)):
        segment_vals = {
            _str(entry.get("name"), "Series"): signed_value(entry, idx) for entry in segment_entries
        }

        sum_values = sum_numeric(segment_vals.get(name) for name in base_series_names)
        is_total = idx in total_categories
        delta_values[idx] = sum_values if not is_total else None

        stack_height = sum(abs(val) for val in segment_vals.values() if val is not None)
        range_height = sum(
            abs(val)
            for name, val in segment_vals.items()
            if name in range_series_names and val is not None
        )

        if is_total:
            override_total = total_override(idx)
            if override_total is None:
                total_value = sum_values if not seen_normal else prev_total
            else:
                total_value = override_total

            row0[idx] = total_value
            cumulative_totals[idx] = total_value
            label_tops[idx] = total_value
            label_bottoms[idx] = total_value

            if total_color:
                offset_points.append((idx, total_color))

            if range_series_names:
                for entry in chart_series:
                    name = _str(entry.get("name"), "Series")
                    if name not in range_series_names:
                        continue
                    val = segment_vals.get(name)
                    row_values[name][idx] = abs(val) if val is not None else None

            if not funnel:
                prev_total = total_value
            min_total = min(min_total, total_value)
            max_total = max(max_total, total_value + range_height)
            continue

        new_total = prev_total
        if funnel:
            total_value = sum_values
            base = (funnel_max_total - total_value) / 2 if funnel_max_total is not None else 0.0
            label_tops[idx] = base + total_value
            label_bottoms[idx] = base
        else:
            new_total = prev_total + sum_values
            total_value = new_total
            base = min(prev_total, new_total)
            label_tops[idx] = max(prev_total, new_total)
            label_bottoms[idx] = base

        row0[idx] = base
        cumulative_totals[idx] = total_value

        for entry in chart_series:
            name = _str(entry.get("name"), "Series")
            val = segment_vals.get(name)
            row_values[name][idx] = abs(val) if val is not None else None

        should_reuse = (funnel and base == 0) or (
            not funnel and reuse_start_base and not seen_normal and base == 0
        )

        if should_reuse and chart_series:
            first_entry = next(
                (
                    entry
                    for entry in chart_series
                    if _str(entry.get("name"), "Series") not in range_series_names
                ),
                chart_series[0],
            )

            first_name = _str(first_entry.get("name"), "Series")
            first_val = segment_vals.get(first_name)
            if first_val is not None and first_val != 0:
                row0[idx] = abs(first_val)
                row_values[first_name][idx] = None
                offset_label_indices.append(idx)

                point_color = start_color
                if point_color is None:
                    point_color = (
                        _color_str(first_entry.get("color")) or DEFAULT_WATERFALL_START_COLOR
                    )
                if point_color:
                    offset_points.append((idx, point_color))

        if not funnel:
            prev_total = new_total
        min_total = min(min_total, base, total_value)
        max_total = max(max_total, base + stack_height, total_value)
        seen_normal = True

    chart_data = CategoryChartData()
    chart_data.categories = cast(Any, categories)

    chart_data.add_series("Offset", cast(Any, row0))
    for entry in chart_series:
        name = _str(entry.get("name"), "Series")
        chart_data.add_series(name, cast(Any, row_values[name]))

    series_colors = [None] + [_color_str(entry.get("color")) for entry in chart_series]

    axis_min = min_total
    axis_max = max_total
    if axis_min > 0:
        axis_min = 0.0
    if axis_max < 0:
        axis_max = 0.0

    axis_min = math.floor(axis_min)
    axis_max = math.ceil(axis_max)
    if axis_max == axis_min:
        axis_max = axis_min + 1

    chart_type = (
        XL_CHART_TYPE.BAR_STACKED if orientation == "horizontal" else XL_CHART_TYPE.COLUMN_STACKED
    )

    return (
        chart_type,
        chart_data,
        {
            "series_colors": series_colors,
            "waterfall": {
                "offset_series_idx": 0,
                "offset_no_fill": True,
                "offset_points": offset_points,
                "axis_min": axis_min,
                "axis_max": axis_max,
                "orientation": orientation,
                "gap_width": _int(wf.get("gap_width"), 80),
                "overlap": _int(wf.get("overlap"), 100),
                "plot_layout": plot_layout,
                "funnel": funnel,
                "label_offset_ratio": wf.get("label_offset_ratio"),
                "label_offset": wf.get("label_offset"),
                "label_gap": wf.get("label_gap"),
                "label_decimals": label_decimals,
                "data_label_decimals": data_label_decimals,
                "data_label_color": wf.get("data_label_color"),
                "data_label_format": wf.get("data_label_format"),
                "offset_label_indices": offset_label_indices,
                "category_label_offset": wf.get("category_label_offset"),
                "category_offset": wf.get("category_offset"),
                "connector_style": wf.get("connector_style"),
                "connector_value": wf.get("connector_value"),
                "connector_line_width": wf.get("connector_line_width"),
                "connector_line_color": wf.get("connector_line_color"),
                "connector_dash_style": wf.get("connector_dash_style"),
                "connector_overlap": wf.get("connector_overlap"),
                "connector_inset": wf.get("connector_inset"),
                "label_collision": wf.get("label_collision"),
                "label_collision_gap": wf.get("label_collision_gap"),
                "series_label_inset": wf.get("series_label_inset"),
                "series_label_left": wf.get("series_label_left"),
                "overlay": {
                    "categories": categories,
                    "chart_series": [_str(entry.get("name"), "Series") for entry in chart_series],
                    "segment_values": segment_values,
                    "cumulative_totals": cumulative_totals,
                    "delta_values": delta_values,
                    "label_tops": label_tops,
                    "label_bottoms": label_bottoms,
                    "total_categories": total_categories,
                },
            },
        },
    )
