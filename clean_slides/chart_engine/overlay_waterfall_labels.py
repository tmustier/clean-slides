"""Label rendering helpers for waterfall overlays."""

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

from pptx.enum.text import PP_ALIGN

from .annotations import add_text_label
from .defaults import (
    DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT,
    DEFAULT_WATERFALL_LABEL_MARGIN,
    DEFAULT_WATERFALL_SERIES_LABEL_INSET,
    DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT,
)
from .geometry import value_to_x, value_to_y
from .overlay_waterfall_data_labels import measure_label_width
from .spec_utils import format_label
from .units import emu_or_default


def build_waterfall_value_label_specs(
    meta: dict,
    categories: list[object],
    cumulative_totals: list[float | None],
    delta_values: list[float | None],
    total_categories,
    label_tops: list[float | None],
    label_bottoms: list[float | None],
    chart_box: tuple[int, int, int, int],
    geometry: dict,
    orientation: str,
    axis_min: float,
    axis_max: float,
    plot_left: float,
    plot_top: float,
    plot_width: float,
    plot_height: float,
    label_gap: int,
    label_offset: int,
    slide_width: int | None,
) -> list[dict[str, object]]:
    """Build value-label box specs for waterfall overlays."""
    label_specs: list[dict[str, object]] = []

    bar_centers = geometry.get("bar_centers") or []

    for idx, _category in enumerate(categories):
        total_value = cumulative_totals[idx] if idx < len(cumulative_totals) else None
        if total_value is None:
            continue
        if idx == 0 or idx in total_categories:
            label_value = total_value
        else:
            label_value = delta_values[idx] if idx < len(delta_values) else None
        label_decimals = meta.get("label_decimals", 0)
        text = format_label(label_value, decimals=label_decimals)
        if text is None:
            continue
        label_width = measure_label_width(text)
        label_height = DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT
        top_value = label_tops[idx] if idx < len(label_tops) else total_value
        bottom_value = label_bottoms[idx] if idx < len(label_bottoms) else None
        if bottom_value is None:
            bottom_value = top_value
        if orientation == "horizontal":
            if idx >= len(bar_centers):
                continue
            anchor_value = top_value
            if label_value is not None and label_value < 0:
                anchor_value = bottom_value
            x_base = value_to_x(anchor_value, axis_min, axis_max, plot_left, plot_width)
            if label_value is not None and label_value < 0:
                x = x_base - label_width - label_offset
            else:
                x = x_base + label_offset
            min_x = chart_box[0] - label_width
            if slide_width is not None:
                max_x = slide_width - label_width
            else:
                max_x = chart_box[0] + chart_box[2] + label_offset - label_width
            if x < min_x:
                x = min_x
            if x > max_x:
                x = max_x
            y = bar_centers[idx] - label_height / 2
        else:
            if idx >= len(bar_centers):
                continue
            x = bar_centers[idx] - label_width / 2
            if label_value is not None and label_value < 0 and bottom_value is not None:
                y_base = value_to_y(bottom_value, axis_min, axis_max, plot_top, plot_height)
                y = y_base + label_gap
            else:
                y_base = value_to_y(top_value, axis_min, axis_max, plot_top, plot_height)
                y = y_base - label_gap - label_height

        anchor = None
        if orientation == "horizontal":
            anchor = "middle"
        elif orientation == "vertical":
            if label_value is not None and label_value < 0:
                anchor = "top"
            else:
                anchor = "bottom"

        label_specs.append(
            {
                "text": text,
                "x": x,
                "y": y,
                "width": label_width,
                "height": label_height,
                "vertical_anchor": anchor,
            }
        )

    return label_specs


def add_waterfall_value_labels(slide, label_specs: list[dict[str, object]]) -> None:
    """Render prepared value-label specs."""
    for spec in label_specs:
        add_text_label(
            slide,
            spec["text"],
            spec["x"],
            spec["y"],
            spec["width"],
            spec["height"],
            margin_left=DEFAULT_WATERFALL_LABEL_MARGIN,
            margin_right=DEFAULT_WATERFALL_LABEL_MARGIN,
            vertical_anchor=spec.get("vertical_anchor"),
        )


def add_waterfall_category_labels(
    slide,
    categories: list[object],
    chart_box: tuple[int, int, int, int],
    geometry: dict,
    orientation: str,
    category_offset: int,
    plot_left: float,
    plot_bottom: float,
) -> None:
    """Render category labels for waterfall overlays."""
    bar_centers = geometry.get("bar_centers") or []

    if orientation == "horizontal":
        for idx, label in enumerate(categories):
            if idx >= len(bar_centers):
                continue
            text = str(label)
            label_width = measure_label_width(text)
            x = plot_left - category_offset - label_width
            min_x = max(0, chart_box[0] - label_width)
            if x < min_x:
                x = min_x
            y = bar_centers[idx] - DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT / 2
            add_text_label(
                slide,
                text,
                x,
                y,
                label_width,
                DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT,
                align=PP_ALIGN.RIGHT,
            )
        return

    slot_width = geometry.get("slot_width")
    if slot_width is None:
        return
    category_y = plot_bottom + category_offset
    plot_left_value = geometry.get("plot_left", plot_left)
    for idx, label in enumerate(categories):
        x = plot_left_value + slot_width * idx
        add_text_label(
            slide,
            str(label),
            x,
            category_y,
            slot_width,
            DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT,
        )


def add_waterfall_series_labels(
    slide,
    chart_box: tuple[int, int, int, int],
    chart_series_names: list[object],
    segment_values: dict,
    orientation: str,
    meta: dict,
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

    stacked = []
    for name in chart_series_names:
        values = segment_values.get(name, [])
        val = values[0] if values else None
        magnitude = abs(val) if val is not None else 0
        stacked.append((name, magnitude))

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
            y - DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT / 2,
            width,
            DEFAULT_WATERFALL_CATEGORY_LABEL_HEIGHT,
            align=PP_ALIGN.RIGHT,
        )
