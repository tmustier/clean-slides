"""Waterfall overlay orchestration."""

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
# pyright: reportUnknownLambdaType=false
# pyright: reportPossiblyUnboundVariable=false

from __future__ import annotations

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


def add_waterfall_overlays(
    slide,
    chart_box: tuple,
    meta: dict,
    slide_size: tuple[int, int] | None = None,
) -> None:
    overlay = meta.get("overlay") if meta else None
    if not overlay:
        return

    categories = overlay.get("categories", [])
    cumulative_totals = overlay.get("cumulative_totals", [])
    delta_values = overlay.get("delta_values", [])
    label_tops = overlay.get("label_tops", [])
    label_bottoms = overlay.get("label_bottoms", [])
    total_categories = overlay.get("total_categories", set())
    chart_series_names = overlay.get("chart_series", [])
    segment_values = overlay.get("segment_values", {})

    axis_min = meta.get("axis_min", 0)
    axis_max = meta.get("axis_max", 0)
    gap_width = int(meta.get("gap_width", 80))
    plot_layout = meta.get("plot_layout") or DEFAULT_WATERFALL_PLOT_LAYOUT
    orientation = normalize_orientation(meta.get("orientation"))

    geometry = compute_category_geometry(
        chart_box,
        plot_layout,
        categories,
        gap_width,
        orientation,
    )
    plot_top = geometry["plot_top"]
    plot_height = geometry["plot_height"]
    plot_left = geometry["plot_left"]
    plot_width = geometry["plot_width"]
    plot_bottom = plot_top + plot_height

    slide_width = slide_size[0] if slide_size else None
    axis_span = chart_box[3] if orientation == "vertical" else chart_box[2]
    label_gap = meta.get("label_gap")
    if label_gap is None:
        label_gap = meta.get("label_offset")
    if label_gap is not None:
        label_gap = coerce_emu(label_gap) or 0
    else:
        ratio = meta.get("label_offset_ratio")
        if ratio is not None:
            label_gap = axis_span * float(ratio)
        elif orientation == "vertical":
            label_gap = DEFAULT_WATERFALL_LABEL_GAP
        else:
            label_gap = axis_span * float(DEFAULT_WATERFALL_LABEL_OFFSET_RATIO)
    if orientation == "horizontal":
        label_offset = max(int(label_gap), 0)
    else:
        label_gap = max(int(label_gap), 0)
        label_offset = label_gap

    connector_style = str(meta.get("connector_style") or "gap").lower()
    connector_value_mode = str(meta.get("connector_value") or "totals").lower()
    if connector_value_mode in {"totals", "total", "running", "end"}:
        connector_values = cumulative_totals
    else:
        connector_values = label_tops

    connector_width = meta.get("connector_line_width")
    connector_color = meta.get("connector_line_color")
    connector_dash = str(meta.get("connector_dash_style") or "long_dash").lower()
    connector_overlap = meta.get("connector_overlap")
    if connector_overlap is None:
        connector_overlap = DEFAULT_WATERFALL_CONNECTOR_OVERLAP
    else:
        connector_overlap = coerce_emu(connector_overlap) or 0

    connector_inset = meta.get("connector_inset")
    if connector_inset is None:
        connector_inset = DEFAULT_WATERFALL_CONNECTOR_INSET if orientation == "horizontal" else 0
    else:
        connector_inset = coerce_emu(connector_inset) or 0

    if connector_width is not None:
        line_width_value = coerce_line_width(connector_width)
    else:
        line_width_value = Pt(0.25)
    if line_width_value is None:
        line_width_value = Pt(0.25)
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
        label_gap=int(label_gap),
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

    category_offset = meta.get("category_label_offset")
    if category_offset is None:
        category_offset = meta.get("category_offset")
    if category_offset is not None:
        category_offset = coerce_emu(category_offset) or 0
    else:
        axis_span = chart_box[3] if orientation == "vertical" else chart_box[2]
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
