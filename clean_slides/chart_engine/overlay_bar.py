"""Chart overlay and manual label layout helpers."""

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

from pathlib import Path

from pptx.oxml.xmlchemy import OxmlElement

from .annotations import add_line_annotation, add_shape_annotation
from .defaults import (
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
    DEFAULT_BAR_PLOT_LAYOUT,
    DEFAULT_BAR_TOTAL_LABEL_FONT_SIZE,
    DEFAULT_BAR_TOTAL_LABEL_HEIGHT,
    DEFAULT_BAR_TOTAL_LABEL_OFFSET,
    DEFAULT_BAR_TOTAL_LABEL_WIDTH,
)
from .geometry import compute_category_geometry, normalize_orientation
from .overlay_bar_legend import add_bar_legend
from .overlay_bar_segments import add_bar_segment_labels
from .overlay_bar_totals_categories import add_bar_category_labels, add_bar_total_labels
from .text_templates import load_txbody_template
from .units import resolve_path


def add_bar_overlays(slide, chart_box: tuple, meta: dict) -> None:
    overlay = meta.get("overlay") if meta else None
    if not overlay:
        return

    base_dir = overlay.get("_base_dir")

    categories = overlay.get("categories", [])
    totals = overlay.get("totals", [])
    total_label_tops = overlay.get("total_label_tops") or []
    series_names = overlay.get("series_names", [])
    series_colors = overlay.get("series_colors", [])
    show_totals = overlay.get("show_totals", False)
    show_legend = overlay.get("show_legend_labels", True)

    axis_min = meta.get("axis_min", 0)
    axis_max = meta.get("axis_max", 0)
    gap_width = int(meta.get("gap_width", 80))
    plot_layout = meta.get("plot_layout") or DEFAULT_BAR_PLOT_LAYOUT
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

    category_width = overlay.get("category_label_width", DEFAULT_BAR_CATEGORY_LABEL_WIDTH)
    category_widths = overlay.get("category_label_widths") or []
    legend_width = overlay.get("legend_label_width", DEFAULT_BAR_LEGEND_LABEL_WIDTH)
    total_width = overlay.get("total_label_width", DEFAULT_BAR_TOTAL_LABEL_WIDTH)
    total_widths = overlay.get("total_label_widths") or []
    category_offsets = overlay.get("category_label_offsets") or []

    category_height = overlay.get("category_label_height", DEFAULT_BAR_CATEGORY_LABEL_HEIGHT)
    category_heights = overlay.get("category_label_heights") or []
    legend_height = overlay.get("legend_label_height", DEFAULT_BAR_LEGEND_LABEL_HEIGHT)
    total_height = overlay.get("total_label_height", DEFAULT_BAR_TOTAL_LABEL_HEIGHT)

    category_font = overlay.get("category_label_font", DEFAULT_BAR_CATEGORY_LABEL_FONT_SIZE)
    legend_font = overlay.get("legend_label_font", DEFAULT_BAR_LEGEND_LABEL_FONT_SIZE)
    total_font = overlay.get("total_label_font", DEFAULT_BAR_TOTAL_LABEL_FONT_SIZE)

    category_color = overlay.get("category_label_color")
    legend_color = overlay.get("legend_label_color")
    total_color = overlay.get("total_label_color")

    total_label_offsets = overlay.get("total_label_offsets") or []

    legend_left_ratio = overlay.get("legend_left_ratio", DEFAULT_BAR_LEGEND_LEFT_RATIO)
    legend_step_ratio = overlay.get("legend_step_ratio", DEFAULT_BAR_LEGEND_STEP_RATIO)
    marker_left_ratio = overlay.get(
        "legend_marker_left_ratio", DEFAULT_BAR_LEGEND_MARKER_LEFT_RATIO
    )
    marker_step_ratio = overlay.get(
        "legend_marker_step_ratio", DEFAULT_BAR_LEGEND_MARKER_STEP_RATIO
    )
    marker_width = overlay.get("legend_marker_width", DEFAULT_BAR_LEGEND_MARKER_WIDTH)
    marker_height = overlay.get("legend_marker_height", DEFAULT_BAR_LEGEND_MARKER_HEIGHT)
    marker_y_offset = overlay.get("legend_marker_y_offset", DEFAULT_BAR_LEGEND_MARKER_Y_OFFSET)
    total_label_offset = overlay.get("total_label_offset", DEFAULT_BAR_TOTAL_LABEL_OFFSET)
    total_label_offsets_x = overlay.get("total_label_offsets_x") or []
    category_offset = overlay.get("category_label_offset", DEFAULT_BAR_CATEGORY_LABEL_OFFSET)
    legend_offset = overlay.get("legend_label_offset", DEFAULT_BAR_LEGEND_LABEL_OFFSET)

    annotations = overlay.get("annotations") or []
    for annotation in annotations:
        if not isinstance(annotation, dict):
            continue
        annotation_spec = dict(annotation)
        if base_dir is not None and "_base_dir" not in annotation_spec:
            annotation_spec["_base_dir"] = base_dir

        kind = (annotation_spec.get("type") or "shape").lower()
        if kind == "line":
            add_line_annotation(slide, annotation_spec)
        else:
            add_shape_annotation(slide, annotation_spec)

    text_style_template = overlay.get("text_style_template")
    text_style_map = overlay.get("text_style_map") or {}
    templates: dict[str, OxmlElement | None] = {}
    template_path: Path | None = None
    if isinstance(text_style_template, str):
        template_path = resolve_path(text_style_template, base_dir)
        for key, sample in text_style_map.items():
            if isinstance(sample, str):
                templates[key] = load_txbody_template(template_path, sample)

    segment_values = overlay.get("segment_values") or []
    add_bar_segment_labels(
        slide,
        overlay=overlay,
        categories=categories,
        series_names=series_names,
        series_colors=series_colors,
        segment_values=segment_values,
        orientation=orientation,
        axis_min=axis_min,
        axis_max=axis_max,
        plot_left=plot_left,
        plot_top=plot_top,
        plot_width=plot_width,
        plot_height=plot_height,
        geometry=geometry,
        template_path=template_path,
        templates=templates,
    )

    if show_totals:
        add_bar_total_labels(
            slide,
            overlay=overlay,
            totals=totals,
            total_label_tops=total_label_tops,
            total_width=total_width,
            total_widths=total_widths,
            total_height=total_height,
            total_font=total_font,
            total_color=total_color,
            total_label_offsets=total_label_offsets,
            total_label_offset=total_label_offset,
            total_label_offsets_x=total_label_offsets_x,
            orientation=orientation,
            axis_min=axis_min,
            axis_max=axis_max,
            plot_left=plot_left,
            plot_top=plot_top,
            plot_width=plot_width,
            plot_height=plot_height,
            geometry=geometry,
            template_path=template_path,
            templates=templates,
        )

    add_bar_category_labels(
        slide,
        overlay=overlay,
        categories=categories,
        orientation=orientation,
        plot_left=plot_left,
        plot_bottom=plot_bottom,
        geometry=geometry,
        category_width=category_width,
        category_widths=category_widths,
        category_height=category_height,
        category_heights=category_heights,
        category_offsets=category_offsets,
        category_offset=category_offset,
        category_font=category_font,
        category_color=category_color,
        template_path=template_path,
        templates=templates,
    )

    if show_legend and series_names:
        add_bar_legend(
            slide,
            overlay=overlay,
            chart_box=chart_box,
            plot_bottom=plot_bottom,
            geometry=geometry,
            series_names=series_names,
            series_colors=series_colors,
            legend_width=legend_width,
            legend_height=legend_height,
            legend_font=legend_font,
            legend_color=legend_color,
            legend_offset=legend_offset,
            legend_left_ratio=legend_left_ratio,
            legend_step_ratio=legend_step_ratio,
            marker_left_ratio=marker_left_ratio,
            marker_step_ratio=marker_step_ratio,
            marker_width=marker_width,
            marker_height=marker_height,
            marker_y_offset=marker_y_offset,
            template_path=template_path,
            templates=templates,
        )
