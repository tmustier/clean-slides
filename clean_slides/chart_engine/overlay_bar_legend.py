"""Legend overlay helpers for bar charts."""

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

from __future__ import annotations

from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement

from .annotations import add_text_label
from .colors import resolve_color
from .text_style import normalize_alignment
from .text_templates import resolve_txbody_template


def add_bar_legend(
    slide,
    *,
    overlay: dict,
    chart_box: tuple,
    plot_bottom: float,
    geometry: dict,
    series_names: list[str],
    series_colors: list[str | None],
    legend_width: int,
    legend_height: int,
    legend_font: int,
    legend_color,
    legend_offset: int,
    legend_left_ratio: float,
    legend_step_ratio: float,
    marker_left_ratio: float,
    marker_step_ratio: float,
    marker_width: int,
    marker_height: int,
    marker_y_offset: int,
    template_path: Path | None,
    templates: dict[str, OxmlElement | None],
) -> None:
    """Render legend labels and optional color markers."""
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
        return

    legend_y = plot_bottom + legend_offset
    if legend_align is None:
        legend_align = PP_ALIGN.LEFT

    for idx, name in enumerate(series_names):
        text = str(name)
        x = geometry["plot_left"] + geometry["plot_width"] * (
            legend_left_ratio + legend_step_ratio * idx
        )
        add_text_label(
            slide,
            text,
            x,
            legend_y,
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
