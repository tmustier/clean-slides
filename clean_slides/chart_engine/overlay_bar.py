"""Chart overlay and manual label layout helpers."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, Callable, Union, cast

from pptx.dml.color import RGBColor
from pptx.util import Pt

from . import annotations as _annotations
from . import text_templates as _text_templates
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
from .units import resolve_path

Number = Union[int, float]
FloatOrNone = Union[float, None]
NumberOrNone = Union[Number, None]
IntOrNone = Union[int, None]
StrOrNone = Union[str, None]
PathOrNone = Union[Path, None]
BaseDirValue = Union[str, Path, None]
ColorValue = Union[RGBColor, str, None]
FontValue = Union[Pt, int, float]
OverlaySpec = Mapping[str, object]
MetaSpec = Mapping[str, object]

AddLineAnnotationFn = Callable[[Any, dict[str, object]], None]
AddShapeAnnotationFn = Callable[[Any, dict[str, object]], None]
LoadTemplateFn = Callable[[Path, str], object]


def _require_attr(module: object, name: str) -> object:
    value = getattr(module, name, None)
    if value is None:
        raise AttributeError(f"{module!r} does not expose {name}")
    return value


add_line_annotation = cast(
    AddLineAnnotationFn,
    _require_attr(_annotations, "add_line_annotation"),
)
add_shape_annotation = cast(
    AddShapeAnnotationFn,
    _require_attr(_annotations, "add_shape_annotation"),
)
load_txbody_template = cast(
    LoadTemplateFn,
    _require_attr(_text_templates, "load_txbody_template"),
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


def _list(value: object) -> list[object]:
    if isinstance(value, list):
        return cast(list[object], value)
    if isinstance(value, tuple):
        return list(cast(tuple[object, ...], value))
    return []


def _sequence_matrix(value: object) -> list[Sequence[object]]:
    matrix: list[Sequence[object]] = []
    for row in _list(value):
        if isinstance(row, list):
            matrix.append(cast(list[object], row))
        elif isinstance(row, tuple):
            matrix.append(list(cast(tuple[object, ...], row)))
    return matrix


def _number(value: object, default: Number = 0) -> Number:
    return value if isinstance(value, (int, float)) else default


def _int(value: object, default: int = 0) -> int:
    return int(_number(value, default))


def _float(value: object, default: float = 0.0) -> float:
    return float(_number(value, default))


def _bool(value: object, default: bool) -> bool:
    return value if isinstance(value, bool) else default


def _optional_str(value: object) -> StrOrNone:
    return value if isinstance(value, str) else None


def _base_dir(value: object) -> BaseDirValue:
    if isinstance(value, (str, Path)):
        return value
    return None


def _font_value(value: object, default: FontValue) -> FontValue:
    if isinstance(value, Pt):
        return value
    if isinstance(value, (int, float)):
        return value
    return default


def _color(value: object) -> ColorValue:
    if isinstance(value, (RGBColor, str)):
        return value
    return None


def _string_list(value: object) -> list[str]:
    result: list[str] = []
    for item in _list(value):
        if isinstance(item, str):
            result.append(item)
    return result


def _string_or_none_list(value: object) -> list[StrOrNone]:
    result: list[StrOrNone] = []
    for item in _list(value):
        if isinstance(item, str):
            result.append(item)
        elif item is None:
            result.append(None)
    return result


def _float_or_none_list(value: object) -> list[FloatOrNone]:
    result: list[FloatOrNone] = []
    for item in _list(value):
        if isinstance(item, (int, float)):
            result.append(float(item))
        else:
            result.append(None)
    return result


def _number_or_none_list(value: object) -> list[NumberOrNone]:
    result: list[NumberOrNone] = []
    for item in _list(value):
        if isinstance(item, (int, float)):
            result.append(item)
        else:
            result.append(None)
    return result


def _int_or_none_list(value: object) -> list[IntOrNone]:
    result: list[IntOrNone] = []
    for item in _list(value):
        if isinstance(item, (int, float)):
            result.append(int(item))
        else:
            result.append(None)
    return result


def _number_list(value: object) -> list[Number]:
    result: list[Number] = []
    for item in _list(value):
        if isinstance(item, (int, float)):
            result.append(item)
    return result


def _plot_layout(value: object) -> dict[str, float]:
    if isinstance(value, dict):
        typed_mapping = cast(dict[object, object], value)
        layout = {
            key: float(item)
            for key, item in typed_mapping.items()
            if isinstance(key, str) and isinstance(item, (int, float))
        }
        if layout:
            return layout

    return dict(DEFAULT_BAR_PLOT_LAYOUT)


def add_bar_overlays(
    slide: Any,
    chart_box: tuple[int, int, int, int],
    meta: MetaSpec,
) -> None:
    overlay = _mapping(meta.get("overlay")) if meta else {}
    if not overlay:
        return

    base_dir = _base_dir(overlay.get("_base_dir"))

    categories = _list(overlay.get("categories"))
    totals = _float_or_none_list(overlay.get("totals"))
    total_label_tops = _float_or_none_list(overlay.get("total_label_tops"))
    series_names = _string_list(overlay.get("series_names"))
    series_colors = _string_or_none_list(overlay.get("series_colors"))
    show_totals = _bool(overlay.get("show_totals", False), False)
    show_legend = _bool(overlay.get("show_legend_labels", True), True)

    axis_min = _float(meta.get("axis_min", 0), 0.0)
    axis_max = _float(meta.get("axis_max", 0), 0.0)
    gap_width = _int(meta.get("gap_width", 80), 80)
    plot_layout = _plot_layout(meta.get("plot_layout"))
    orientation = normalize_orientation(_optional_str(meta.get("orientation")))

    geometry = compute_category_geometry(
        chart_box,
        plot_layout,
        categories,
        gap_width,
        orientation,
    )
    plot_top = _float(geometry.get("plot_top"), 0.0)
    plot_height = _float(geometry.get("plot_height"), 0.0)
    plot_left = _float(geometry.get("plot_left"), 0.0)
    plot_width = _float(geometry.get("plot_width"), 0.0)
    plot_bottom = plot_top + plot_height

    category_width = _int(
        overlay.get("category_label_width", DEFAULT_BAR_CATEGORY_LABEL_WIDTH),
        int(DEFAULT_BAR_CATEGORY_LABEL_WIDTH),
    )
    category_widths = _number_or_none_list(overlay.get("category_label_widths"))
    legend_width = _int(
        overlay.get("legend_label_width", DEFAULT_BAR_LEGEND_LABEL_WIDTH),
        int(DEFAULT_BAR_LEGEND_LABEL_WIDTH),
    )
    total_width = _int(
        overlay.get("total_label_width", DEFAULT_BAR_TOTAL_LABEL_WIDTH),
        int(DEFAULT_BAR_TOTAL_LABEL_WIDTH),
    )
    total_widths = _int_or_none_list(overlay.get("total_label_widths"))
    category_offsets = _number_list(overlay.get("category_label_offsets"))

    category_height = _int(
        overlay.get("category_label_height", DEFAULT_BAR_CATEGORY_LABEL_HEIGHT),
        int(DEFAULT_BAR_CATEGORY_LABEL_HEIGHT),
    )
    category_heights = _number_or_none_list(overlay.get("category_label_heights"))
    legend_height = _int(
        overlay.get("legend_label_height", DEFAULT_BAR_LEGEND_LABEL_HEIGHT),
        int(DEFAULT_BAR_LEGEND_LABEL_HEIGHT),
    )
    total_height = _int(
        overlay.get("total_label_height", DEFAULT_BAR_TOTAL_LABEL_HEIGHT),
        int(DEFAULT_BAR_TOTAL_LABEL_HEIGHT),
    )

    category_font = _font_value(
        overlay.get("category_label_font", DEFAULT_BAR_CATEGORY_LABEL_FONT_SIZE),
        DEFAULT_BAR_CATEGORY_LABEL_FONT_SIZE,
    )
    legend_font = _font_value(
        overlay.get("legend_label_font", DEFAULT_BAR_LEGEND_LABEL_FONT_SIZE),
        DEFAULT_BAR_LEGEND_LABEL_FONT_SIZE,
    )
    total_font = _font_value(
        overlay.get("total_label_font", DEFAULT_BAR_TOTAL_LABEL_FONT_SIZE),
        DEFAULT_BAR_TOTAL_LABEL_FONT_SIZE,
    )

    category_color = _color(overlay.get("category_label_color"))
    legend_color = _color(overlay.get("legend_label_color"))
    total_color = _color(overlay.get("total_label_color"))

    total_label_offsets = _number_or_none_list(overlay.get("total_label_offsets"))

    legend_left_ratio = _float(
        overlay.get("legend_left_ratio", DEFAULT_BAR_LEGEND_LEFT_RATIO),
        float(DEFAULT_BAR_LEGEND_LEFT_RATIO),
    )
    legend_step_ratio = _float(
        overlay.get("legend_step_ratio", DEFAULT_BAR_LEGEND_STEP_RATIO),
        float(DEFAULT_BAR_LEGEND_STEP_RATIO),
    )
    marker_left_ratio = _float(
        overlay.get("legend_marker_left_ratio", DEFAULT_BAR_LEGEND_MARKER_LEFT_RATIO),
        float(DEFAULT_BAR_LEGEND_MARKER_LEFT_RATIO),
    )
    marker_step_ratio = _float(
        overlay.get("legend_marker_step_ratio", DEFAULT_BAR_LEGEND_MARKER_STEP_RATIO),
        float(DEFAULT_BAR_LEGEND_MARKER_STEP_RATIO),
    )
    marker_width = _int(
        overlay.get("legend_marker_width", DEFAULT_BAR_LEGEND_MARKER_WIDTH),
        int(DEFAULT_BAR_LEGEND_MARKER_WIDTH),
    )
    marker_height = _int(
        overlay.get("legend_marker_height", DEFAULT_BAR_LEGEND_MARKER_HEIGHT),
        int(DEFAULT_BAR_LEGEND_MARKER_HEIGHT),
    )
    marker_y_offset = _int(
        overlay.get("legend_marker_y_offset", DEFAULT_BAR_LEGEND_MARKER_Y_OFFSET),
        int(DEFAULT_BAR_LEGEND_MARKER_Y_OFFSET),
    )
    total_label_offset = _int(
        overlay.get("total_label_offset", DEFAULT_BAR_TOTAL_LABEL_OFFSET),
        int(DEFAULT_BAR_TOTAL_LABEL_OFFSET),
    )
    total_label_offsets_x = _number_list(overlay.get("total_label_offsets_x"))
    category_offset = _int(
        overlay.get("category_label_offset", DEFAULT_BAR_CATEGORY_LABEL_OFFSET),
        int(DEFAULT_BAR_CATEGORY_LABEL_OFFSET),
    )
    legend_offset = _int(
        overlay.get("legend_label_offset", DEFAULT_BAR_LEGEND_LABEL_OFFSET),
        int(DEFAULT_BAR_LEGEND_LABEL_OFFSET),
    )

    annotations = _list(overlay.get("annotations"))
    for annotation in annotations:
        annotation_spec = _mapping(annotation)
        if not annotation_spec:
            continue
        if base_dir is not None and "_base_dir" not in annotation_spec:
            annotation_spec["_base_dir"] = base_dir

        kind_raw = annotation_spec.get("type")
        kind = kind_raw.lower() if isinstance(kind_raw, str) else "shape"
        if kind == "line":
            add_line_annotation(slide, annotation_spec)
        else:
            add_shape_annotation(slide, annotation_spec)

    text_style_template = overlay.get("text_style_template")
    text_style_map = _mapping(overlay.get("text_style_map"))

    templates: dict[str, object] = {}
    template_path: PathOrNone = None

    if isinstance(text_style_template, Path):
        template_path = resolve_path(str(text_style_template), base_dir)
    elif isinstance(text_style_template, str):
        template_path = resolve_path(text_style_template, base_dir)

    if template_path is not None:
        for key, sample in text_style_map.items():
            if isinstance(sample, str):
                templates[key] = load_txbody_template(template_path, sample)

    segment_values = _sequence_matrix(overlay.get("segment_values"))
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
            geometry={
                "plot_left": plot_left,
                "plot_width": plot_width,
            },
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
