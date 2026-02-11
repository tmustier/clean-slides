"""Segment label overlay helpers for bar charts."""

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
# pyright: reportPossiblyUnboundVariable=false

from __future__ import annotations

from pathlib import Path

from pptx.oxml.xmlchemy import OxmlElement

from .annotations import add_text_label
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
from .text_templates import resolve_txbody_template
from .units import (
    coerce_offset_value,
    emu_or_default,
    normalize_offset_matrix,
)


def add_bar_segment_labels(
    slide,
    *,
    overlay: dict,
    categories: list,
    series_names: list[str],
    series_colors: list[str | None],
    segment_values: list,
    orientation: str,
    axis_min: float,
    axis_max: float,
    plot_left: float,
    plot_top: float,
    plot_width: float,
    plot_height: float,
    geometry: dict,
    template_path: Path | None,
    templates: dict[str, OxmlElement | None],
) -> None:
    segment_configs = normalize_list(overlay.get("segment_labels"))
    for segment_cfg in segment_configs:
        if not segment_cfg or not segment_cfg.get("show", True):
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
            categories,
            segment_cfg.get("categories") or segment_cfg.get("category_indices"),
        )
        category_filter_set = set(category_filter) if category_filter else None

        positions = segment_cfg.get("positions")
        if positions is not None:
            ratios = [float(val) for val in normalize_list(positions)]
        else:
            offset_ratio = float(
                segment_cfg.get("offset_ratio", DEFAULT_BAR_SEGMENT_LABEL_OFFSET_RATIO)
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
            DEFAULT_BAR_SEGMENT_LABEL_WIDTH,
        )
        segment_height = emu_or_default(
            segment_cfg.get("height"),
            DEFAULT_BAR_SEGMENT_LABEL_HEIGHT,
        )
        segment_font = segment_cfg.get("font_size", DEFAULT_BAR_SEGMENT_LABEL_FONT_SIZE)
        segment_color = segment_cfg.get("text_color", "bg1")
        segment_fill = segment_cfg.get("fill", "series")
        segment_decimals = segment_cfg.get("decimals", 0)
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

        for cat_idx in range(len(categories)):
            if category_filter_set and cat_idx not in category_filter_set:
                continue
            running = 0.0
            for series_idx, series_vals in enumerate(segment_values):
                value = numeric_value(safe_value(series_vals, cat_idx))
                if value is None:
                    continue
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
                            x_center = (x_start + x_end) / 2
                            y_center = (
                                geometry["bar_tops"][cat_idx] + geometry["bar_height"] * ratio
                            )
                        else:
                            x_center = (
                                geometry["bar_lefts"][cat_idx] + geometry["bar_width"] * ratio
                            )
                            y_center = (y_bottom + y_top) / 2
                        fill_color = None
                        if segment_fill == "series" and series_idx < len(series_colors):
                            fill_color = series_colors[series_idx]
                        elif segment_fill not in {None, "none", "transparent", False}:
                            fill_color = segment_fill

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
                            segment_width = label.width
                running += value
