"""Total/category label overlay helpers for bar charts."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, Callable, Union, cast

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

from . import annotations as _annotations
from . import text_templates as _text_templates
from .geometry import value_to_x, value_to_y
from .spec_utils import format_label

Number = Union[int, float]
PathOrNone = Union[Path, None]
FloatOrNone = Union[float, None]
NumberOrNone = Union[Number, None]
IntOrNone = Union[int, None]
ColorValue = Union[RGBColor, str, None]
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


def _number(value: object, default: Number = 0) -> Number:
    return value if isinstance(value, (int, float)) else default


def _str_or_default(value: object, default: str) -> str:
    return value if isinstance(value, str) else default


def _bar_center(geometry: Geometry, idx: int) -> FloatOrNone:
    bar_centers_obj = geometry.get("bar_centers")
    if not isinstance(bar_centers_obj, list):
        return None

    bar_centers = cast(list[object], bar_centers_obj)
    if idx < 0 or idx >= len(bar_centers):
        return None

    center = bar_centers[idx]
    if isinstance(center, (int, float)):
        return float(center)
    return None


def add_bar_total_labels(
    slide: Any,
    *,
    overlay: Mapping[str, object],
    totals: Sequence[FloatOrNone],
    total_label_tops: Sequence[FloatOrNone],
    total_width: int,
    total_widths: list[IntOrNone],
    total_height: int,
    total_font: Union[Pt, int, float],
    total_color: ColorValue,
    total_label_offsets: Sequence[NumberOrNone],
    total_label_offset: Number,
    total_label_offsets_x: Sequence[Number],
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
    """Render bar total labels."""

    margin_left = _number(overlay.get("total_label_margin_left", 25400), 25400)
    margin_right = _number(overlay.get("total_label_margin_right", 25400), 25400)
    margin_top = _number(overlay.get("total_label_margin_top", 0), 0)
    margin_bottom = _number(overlay.get("total_label_margin_bottom", 0), 0)
    vertical_anchor = _str_or_default(overlay.get("total_label_anchor"), "bottom")
    bw_mode = _str_or_default(overlay.get("total_label_bw_mode"), "gray")

    for idx, total_value in enumerate(totals):
        if total_value is None:
            continue

        text = format_label(total_value)
        if text is None:
            continue

        label_width_value: Number = total_width
        if idx < len(total_widths):
            configured_width = total_widths[idx]
            if isinstance(configured_width, (int, float)):
                label_width_value = configured_width
        label_width = float(label_width_value)

        offset_axis = _number(
            total_label_offsets_x[idx] if idx < len(total_label_offsets_x) else 0,
            0,
        )

        offset = _number(
            total_label_offsets[idx] if idx < len(total_label_offsets) else None,
            total_label_offset,
        )

        top_value = total_label_tops[idx] if idx < len(total_label_tops) else total_value
        if top_value is None:
            top_value = total_value

        bar_center = _bar_center(geometry, idx)
        if bar_center is None:
            continue

        if orientation == "horizontal":
            x_base = value_to_x(top_value, axis_min, axis_max, plot_left, plot_width)
            if top_value < 0:
                x = x_base - label_width - offset
            else:
                x = x_base + offset
            y = bar_center - total_height / 2 + offset_axis
        else:
            x = bar_center - label_width / 2 - offset_axis
            y_base = value_to_y(top_value, axis_min, axis_max, plot_top, plot_height)
            if top_value < 0:
                y = y_base + offset
            else:
                y = y_base - total_height - offset

        label = add_text_label(
            slide,
            text,
            x,
            y,
            label_width,
            float(total_height),
            font_size=total_font,
            color=total_color,
            margin_left=margin_left,
            margin_right=margin_right,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
            vertical_anchor=vertical_anchor,
            bw_mode=bw_mode,
            txbody_template=resolve_txbody_template(
                template_path,
                text,
                templates.get("total"),
            ),
        )

        if label is not None and idx < len(total_widths):
            rendered_width = getattr(label, "width", None)
            if isinstance(rendered_width, int):
                total_widths[idx] = rendered_width


def add_bar_category_labels(
    slide: Any,
    *,
    overlay: Mapping[str, object],
    categories: Sequence[object],
    orientation: str,
    plot_left: float,
    plot_bottom: float,
    geometry: Geometry,
    category_width: int,
    category_widths: Sequence[NumberOrNone],
    category_height: int,
    category_heights: Sequence[NumberOrNone],
    category_offsets: Sequence[Number],
    category_offset: Number,
    category_font: Union[Pt, int, float],
    category_color: ColorValue,
    template_path: PathOrNone,
    templates: TemplateMap,
) -> None:
    """Render category labels around bars."""

    margin_left = _number(overlay.get("category_label_margin_left", 0), 0)
    margin_right = _number(overlay.get("category_label_margin_right", 0), 0)
    margin_top = _number(overlay.get("category_label_margin_top", 0), 0)
    margin_bottom = _number(overlay.get("category_label_margin_bottom", 0), 0)
    bw_mode = _str_or_default(overlay.get("category_label_bw_mode"), "auto")

    if orientation == "horizontal":
        vertical_anchor = _str_or_default(overlay.get("category_label_anchor"), "center")

        for idx, label in enumerate(categories):
            text = str(label)

            configured_width = category_widths[idx] if idx < len(category_widths) else None
            label_width = float(_number(configured_width, category_width))

            configured_height = category_heights[idx] if idx < len(category_heights) else None
            label_height = float(_number(configured_height, category_height))

            offset = _number(category_offsets[idx] if idx < len(category_offsets) else 0, 0)
            bar_center = _bar_center(geometry, idx)
            if bar_center is None:
                continue

            x = plot_left + category_offset - label_width
            y = bar_center - label_height / 2 + offset
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
                margin_left=margin_left,
                margin_right=margin_right,
                margin_top=margin_top,
                margin_bottom=margin_bottom,
                vertical_anchor=vertical_anchor,
                bw_mode=bw_mode,
                txbody_template=resolve_txbody_template(
                    template_path,
                    text,
                    templates.get("category"),
                ),
            )
        return

    vertical_anchor = _str_or_default(overlay.get("category_label_anchor"), "top")
    category_y = plot_bottom + category_offset

    for idx, label in enumerate(categories):
        text = str(label)

        configured_width = category_widths[idx] if idx < len(category_widths) else None
        label_width = float(_number(configured_width, category_width))

        configured_height = category_heights[idx] if idx < len(category_heights) else None
        label_height = float(_number(configured_height, category_height))

        offset = _number(category_offsets[idx] if idx < len(category_offsets) else 0, 0)
        bar_center = _bar_center(geometry, idx)
        if bar_center is None:
            continue

        x = bar_center - label_width / 2 - offset
        add_text_label(
            slide,
            text,
            x,
            category_y,
            label_width,
            label_height,
            font_size=category_font,
            color=category_color,
            margin_left=margin_left,
            margin_right=margin_right,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
            vertical_anchor=vertical_anchor,
            bw_mode=bw_mode,
            txbody_template=resolve_txbody_template(
                template_path,
                text,
                templates.get("category"),
            ),
        )
