"""Legend overlay helpers for bar charts."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Callable, Protocol, Union, cast

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

from . import annotations as _annotations
from . import text_templates as _text_templates
from .colors import resolve_color
from .text_style import normalize_alignment

Number = Union[int, float]
PathOrNone = Union[Path, None]
StrOrNone = Union[str, None]
ColorValue = Union[RGBColor, str, None]
TemplateMap = Mapping[str, object]
LegendGeometry = Mapping[str, float]


class _ShapeForeColorLike(Protocol):
    theme_color: object
    rgb: object


class _ShapeFillLike(Protocol):
    fore_color: _ShapeForeColorLike

    def solid(self) -> None: ...


class _ShapeLineFillLike(Protocol):
    def background(self) -> None: ...


class _ShapeLineLike(Protocol):
    fill: _ShapeLineFillLike


class _ShapeLike(Protocol):
    fill: _ShapeFillLike
    line: _ShapeLineLike


class _SlideShapesLike(Protocol):
    def add_shape(
        self,
        shape_type: object,
        left: object,
        top: object,
        width: object,
        height: object,
    ) -> _ShapeLike: ...


class _SlideLike(Protocol):
    shapes: _SlideShapesLike


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


def _optional_str(value: object) -> StrOrNone:
    return value if isinstance(value, str) else None


def _bool(value: object, default: bool) -> bool:
    return value if isinstance(value, bool) else default


def add_bar_legend(
    slide: object,
    *,
    overlay: Mapping[str, object],
    chart_box: tuple[int, int, int, int],
    plot_bottom: float,
    geometry: LegendGeometry,
    series_names: Sequence[str],
    series_colors: Sequence[StrOrNone],
    legend_width: int,
    legend_height: int,
    legend_font: Union[Pt, int, float],
    legend_color: ColorValue,
    legend_offset: int,
    legend_left_ratio: float,
    legend_step_ratio: float,
    marker_left_ratio: float,
    marker_step_ratio: float,
    marker_width: int,
    marker_height: int,
    marker_y_offset: int,
    template_path: PathOrNone,
    templates: TemplateMap,
) -> None:
    """Render legend labels and optional color markers."""
    slide_like = cast(_SlideLike, slide)

    legend_layout = _optional_str(overlay.get("legend_layout"))
    legend_align = normalize_alignment(_optional_str(overlay.get("legend_alignment")))
    legend_show_markers = _bool(overlay.get("legend_show_markers", True), True)

    margin_left = _number(overlay.get("legend_label_margin_left", 0), 0)
    margin_right = _number(overlay.get("legend_label_margin_right", 0), 0)
    margin_top = _number(overlay.get("legend_label_margin_top", 0), 0)
    margin_bottom = _number(overlay.get("legend_label_margin_bottom", 0), 0)
    vertical_anchor = _optional_str(overlay.get("legend_label_anchor")) or "center"
    bw_mode = _optional_str(overlay.get("legend_label_bw_mode")) or "auto"

    if legend_layout == "left":
        legend_left_offset = _number(overlay.get("legend_left_offset", 0), 0)
        legend_top_offset = _number(overlay.get("legend_top_offset", 0), 0)
        legend_step = _number(overlay.get("legend_step", legend_height), legend_height)
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
                margin_left=margin_left,
                margin_right=margin_right,
                margin_top=margin_top,
                margin_bottom=margin_bottom,
                vertical_anchor=vertical_anchor,
                bw_mode=bw_mode,
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

    plot_left = geometry["plot_left"]
    plot_width = geometry["plot_width"]

    for idx, name in enumerate(series_names):
        text = str(name)
        x = plot_left + plot_width * (legend_left_ratio + legend_step_ratio * idx)
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
            margin_left=margin_left,
            margin_right=margin_right,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
            vertical_anchor=vertical_anchor,
            bw_mode=bw_mode,
            txbody_template=resolve_txbody_template(
                template_path,
                text,
                templates.get("legend"),
            ),
        )

        series_color = series_colors[idx] if idx < len(series_colors) else None
        if legend_show_markers and series_color:
            marker_x = plot_left + plot_width * (marker_left_ratio + marker_step_ratio * idx)
            marker_y = legend_y + marker_y_offset
            shape = slide_like.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                int(marker_x),
                int(marker_y),
                int(marker_width),
                int(marker_height),
            )
            shape.fill.solid()
            rgb, theme = resolve_color(series_color)
            if theme is not None:
                shape.fill.fore_color.theme_color = theme
            elif rgb is not None:
                shape.fill.fore_color.rgb = rgb
            shape.line.fill.background()
