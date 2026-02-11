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

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement

from .annotations import (
    add_line_annotation,
    add_shape_annotation,
    add_text_label,
)
from .colors import resolve_color
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
from .geometry import (
    compute_category_geometry,
    normalize_orientation,
    value_to_x,
    value_to_y,
)
from .overlay_bar_segments import add_bar_segment_labels
from .spec_utils import (
    format_label,
)
from .text_style import normalize_alignment
from .text_templates import load_txbody_template, resolve_txbody_template
from .units import (
    resolve_path,
)


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

    # total labels (stacked)
    if show_totals:
        for idx, total_value in enumerate(totals):
            if total_value is None:
                continue
            text = format_label(total_value)
            if text is None:
                continue
            label_width = total_widths[idx] if idx < len(total_widths) else total_width
            offset_axis = total_label_offsets_x[idx] if idx < len(total_label_offsets_x) else 0
            offset = (
                total_label_offsets[idx] if idx < len(total_label_offsets) else total_label_offset
            )
            top_value = (
                total_label_tops[idx]
                if idx < len(total_label_tops) and total_label_tops[idx] is not None
                else total_value
            )
            if orientation == "horizontal":
                x_base = value_to_x(top_value, axis_min, axis_max, plot_left, plot_width)
                if top_value is not None and top_value < 0:
                    x = x_base - label_width - offset
                else:
                    x = x_base + offset
                y = geometry["bar_centers"][idx] - total_height / 2 + offset_axis
            else:
                x = geometry["bar_centers"][idx] - label_width / 2 - offset_axis
                y_base = value_to_y(top_value, axis_min, axis_max, plot_top, plot_height)
                if top_value is not None and top_value < 0:
                    y = y_base + offset
                else:
                    y = y_base - total_height - offset
            label = add_text_label(
                slide,
                text,
                x,
                y,
                label_width,
                total_height,
                font_size=total_font,
                color=total_color,
                margin_left=overlay.get("total_label_margin_left", 25400),
                margin_right=overlay.get("total_label_margin_right", 25400),
                margin_top=overlay.get("total_label_margin_top", 0),
                margin_bottom=overlay.get("total_label_margin_bottom", 0),
                vertical_anchor=overlay.get("total_label_anchor", "bottom"),
                bw_mode=overlay.get("total_label_bw_mode", "gray"),
                txbody_template=resolve_txbody_template(
                    template_path,
                    text,
                    templates.get("total"),
                ),
            )
            if label is not None and idx < len(total_widths):
                total_widths[idx] = label.width

    # category labels
    if orientation == "horizontal":
        for idx, label in enumerate(categories):
            text = str(label)
            label_width = category_widths[idx] if idx < len(category_widths) else category_width
            label_height = category_heights[idx] if idx < len(category_heights) else category_height
            offset = category_offsets[idx] if idx < len(category_offsets) else 0
            x = plot_left + category_offset - label_width
            y = geometry["bar_centers"][idx] - label_height / 2 + offset
            add_text_label(
                slide,
                text,
                x,
                y,
                label_width,
                label_height,
                align=PP_ALIGN.RIGHT,
                font_size=category_font,
                color=category_color,
                margin_left=overlay.get("category_label_margin_left", 0),
                margin_right=overlay.get("category_label_margin_right", 0),
                margin_top=overlay.get("category_label_margin_top", 0),
                margin_bottom=overlay.get("category_label_margin_bottom", 0),
                vertical_anchor=overlay.get("category_label_anchor", "center"),
                bw_mode=overlay.get("category_label_bw_mode", "auto"),
                txbody_template=resolve_txbody_template(
                    template_path,
                    text,
                    templates.get("category"),
                ),
            )
    else:
        category_y = plot_bottom + category_offset
        for idx, label in enumerate(categories):
            text = str(label)
            label_width = category_widths[idx] if idx < len(category_widths) else category_width
            label_height = category_heights[idx] if idx < len(category_heights) else category_height
            offset = category_offsets[idx] if idx < len(category_offsets) else 0
            x = geometry["bar_centers"][idx] - label_width / 2 - offset
            add_text_label(
                slide,
                text,
                x,
                category_y,
                label_width,
                label_height,
                font_size=category_font,
                color=category_color,
                margin_left=overlay.get("category_label_margin_left", 0),
                margin_right=overlay.get("category_label_margin_right", 0),
                margin_top=overlay.get("category_label_margin_top", 0),
                margin_bottom=overlay.get("category_label_margin_bottom", 0),
                vertical_anchor=overlay.get("category_label_anchor", "top"),
                bw_mode=overlay.get("category_label_bw_mode", "auto"),
                txbody_template=resolve_txbody_template(
                    template_path,
                    text,
                    templates.get("category"),
                ),
            )

    # legend labels + color markers
    if show_legend and series_names:
        legend_layout = overlay.get("legend_layout")
        legend_align = normalize_alignment(overlay.get("legend_alignment"))
        legend_show_markers = overlay.get("legend_show_markers", True)

        if legend_layout == "left":
            legend_left_offset = overlay.get("legend_left_offset", 0)
            legend_top_offset = overlay.get("legend_top_offset", 0)
            legend_step = overlay.get("legend_step", legend_height)
            legend_x = chart_box[0] + legend_left_offset
            legend_y = plot_bottom + legend_top_offset
            if legend_align is None:
                legend_align = PP_ALIGN.RIGHT

            for idx, name in enumerate(series_names):
                text = str(name)
                y = legend_y + legend_step * idx
                add_text_label(
                    slide,
                    text,
                    legend_x,
                    y,
                    legend_width,
                    legend_height,
                    align=legend_align,
                    font_size=legend_font,
                    color=legend_color,
                    margin_left=overlay.get("legend_label_margin_left", 0),
                    margin_right=overlay.get("legend_label_margin_right", 0),
                    margin_top=overlay.get("legend_label_margin_top", 0),
                    margin_bottom=overlay.get("legend_label_margin_bottom", 0),
                    vertical_anchor=overlay.get("legend_label_anchor", "center"),
                    bw_mode=overlay.get("legend_label_bw_mode", "auto"),
                    txbody_template=resolve_txbody_template(
                        template_path,
                        text,
                        templates.get("legend"),
                    ),
                )
        else:
            legend_y = plot_bottom + legend_offset
            if legend_align is None:
                legend_align = PP_ALIGN.LEFT
            for idx, name in enumerate(series_names):
                text = str(name)
                label_width = legend_width
                x = geometry["plot_left"] + geometry["plot_width"] * (
                    legend_left_ratio + legend_step_ratio * idx
                )
                add_text_label(
                    slide,
                    text,
                    x,
                    legend_y,
                    label_width,
                    legend_height,
                    align=legend_align,
                    font_size=legend_font,
                    color=legend_color,
                    margin_left=overlay.get("legend_label_margin_left", 0),
                    margin_right=overlay.get("legend_label_margin_right", 0),
                    margin_top=overlay.get("legend_label_margin_top", 0),
                    margin_bottom=overlay.get("legend_label_margin_bottom", 0),
                    vertical_anchor=overlay.get("legend_label_anchor", "center"),
                    bw_mode=overlay.get("legend_label_bw_mode", "auto"),
                    txbody_template=resolve_txbody_template(
                        template_path,
                        text,
                        templates.get("legend"),
                    ),
                )

                if legend_show_markers and idx < len(series_colors) and series_colors[idx]:
                    marker_x = geometry["plot_left"] + geometry["plot_width"] * (
                        marker_left_ratio + marker_step_ratio * idx
                    )
                    marker_y = legend_y + marker_y_offset
                    shape = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE,
                        int(marker_x),
                        int(marker_y),
                        int(marker_width),
                        int(marker_height),
                    )
                    shape.fill.solid()
                    rgb, theme = resolve_color(series_colors[idx])
                    if theme is not None:
                        shape.fill.fore_color.theme_color = theme
                    elif rgb is not None:
                        shape.fill.fore_color.rgb = rgb
                    shape.line.fill.background()
