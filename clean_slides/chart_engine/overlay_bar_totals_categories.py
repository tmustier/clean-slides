"""Total/category label overlay helpers for bar charts."""

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

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement

from .annotations import add_text_label
from .geometry import value_to_x, value_to_y
from .spec_utils import format_label
from .text_templates import resolve_txbody_template


def add_bar_total_labels(
    slide,
    *,
    overlay: dict,
    totals: list,
    total_label_tops: list,
    total_width: int,
    total_widths: list,
    total_height: int,
    total_font: int,
    total_color,
    total_label_offsets: list,
    total_label_offset: int,
    total_label_offsets_x: list,
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
    """Render bar total labels."""

    for idx, total_value in enumerate(totals):
        if total_value is None:
            continue
        text = format_label(total_value)
        if text is None:
            continue

        label_width = (
            total_widths[idx]
            if idx < len(total_widths) and total_widths[idx] is not None
            else total_width
        )
        offset_axis = total_label_offsets_x[idx] if idx < len(total_label_offsets_x) else 0
        offset = (
            total_label_offsets[idx]
            if idx < len(total_label_offsets) and total_label_offsets[idx] is not None
            else total_label_offset
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


def add_bar_category_labels(
    slide,
    *,
    overlay: dict,
    categories: list,
    orientation: str,
    plot_left: float,
    plot_bottom: float,
    geometry: dict,
    category_width: int,
    category_widths: list,
    category_height: int,
    category_heights: list,
    category_offsets: list,
    category_offset: int,
    category_font: int,
    category_color,
    template_path: Path | None,
    templates: dict[str, OxmlElement | None],
) -> None:
    """Render category labels around bars."""

    if orientation == "horizontal":
        for idx, label in enumerate(categories):
            text = str(label)
            label_width = (
                category_widths[idx]
                if idx < len(category_widths) and category_widths[idx] is not None
                else category_width
            )
            label_height = (
                category_heights[idx]
                if idx < len(category_heights) and category_heights[idx] is not None
                else category_height
            )
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
        return

    category_y = plot_bottom + category_offset
    for idx, label in enumerate(categories):
        text = str(label)
        label_width = (
            category_widths[idx]
            if idx < len(category_widths) and category_widths[idx] is not None
            else category_width
        )
        label_height = (
            category_heights[idx]
            if idx < len(category_heights) and category_heights[idx] is not None
            else category_height
        )
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
