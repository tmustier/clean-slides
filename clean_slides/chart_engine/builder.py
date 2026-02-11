"""Chart slide construction orchestration."""

# pyright: reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false
# pyright: reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false
# pyright: reportMissingParameterType=false
# pyright: reportMissingTypeArgument=false
# pyright: reportGeneralTypeIssues=false
# pyright: reportAttributeAccessIssue=false
# pyright: reportUnknownLambdaType=false
# pyright: reportUnnecessaryIsInstance=false

from __future__ import annotations

import copy
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from .annotations import add_waterfall_title
from .colors import apply_color
from .defaults import (
    CLEAN_SLIDES_CONTENT_BOX,
    CLEAN_SLIDES_TEMPLATE_PATH,
    DEFAULT_BAR_DATA_LABEL_FONT_SIZE,
    DEFAULT_BAR_DATA_LABEL_FORMAT,
    DEFAULT_BAR_LEGEND_LABEL_HEIGHT,
    DEFAULT_BAR_LEGEND_LABEL_OFFSET,
    DEFAULT_BAR_LEGEND_MARKER_HEIGHT,
    DEFAULT_BAR_LEGEND_MARKER_Y_OFFSET,
    DEFAULT_BAR_OVERLAY_BAND_EXTRA,
    DEFAULT_WATERFALL_CHART_BOX,
    DEFAULT_WATERFALL_TITLE_OFFSET,
    WATERFALL_TYPES,
)
from .overlays import (
    add_bar_overlays,
    add_waterfall_overlays,
    apply_waterfall_data_label_layout,
    get_chart_series,
)
from .payloads import build_bar_payload, build_waterfall_payload
from .spec_utils import normalize_list
from .style import (
    apply_bar_chart_style,
    apply_series_colors,
    apply_waterfall_chart_style,
    apply_waterfall_data_labels,
    apply_waterfall_style,
)
from .template_ops import ChartTemplateReplacement, apply_chart_template_replacements
from .text_style import normalize_label_position
from .units import coerce_emu


def apply_chart_template_dlbls(
    target_chart,
    template_path: Path,
    slide_index: int = 0,
    chart_index: int = 0,
    series_index: int = 0,
) -> None:
    prs = Presentation(str(template_path))
    if slide_index < 0 or slide_index >= len(prs.slides):
        return
    slide = prs.slides[slide_index]
    template_charts = [shape.chart for shape in slide.shapes if shape.has_chart]
    if chart_index < 0 or chart_index >= len(template_charts):
        return
    template_chart = template_charts[chart_index]

    template_series = get_chart_series(template_chart._chartSpace)
    target_series = get_chart_series(target_chart._chartSpace)
    if not template_series or not target_series:
        return
    if series_index < 0 or series_index >= len(template_series):
        return
    if series_index >= len(target_series):
        return

    tmpl_ser = template_series[series_index]
    tgt_ser = target_series[series_index]

    tmpl_dlbls = tmpl_ser.find(qn("c:dLbls"))
    if tmpl_dlbls is None:
        return
    existing = tgt_ser.find(qn("c:dLbls"))
    if existing is not None:
        tgt_ser.remove(existing)
    tgt_ser.insert(len(tgt_ser), copy.deepcopy(tmpl_dlbls))


def resolve_series_indices(series_spec, series_names: list[str]) -> list[int]:
    indices: list[int] = []
    for item in normalize_list(series_spec):
        if isinstance(item, int):
            indices.append(item)
        elif isinstance(item, str) and item in series_names:
            indices.append(series_names.index(item))
    return list(dict.fromkeys(indices))


def apply_data_label_style(labels, data_cfg: dict) -> None:
    labels.number_format = data_cfg.get("format", DEFAULT_BAR_DATA_LABEL_FORMAT)
    labels.number_format_is_linked = False

    font_size = data_cfg.get("font_size", DEFAULT_BAR_DATA_LABEL_FONT_SIZE)
    if hasattr(font_size, "pt"):
        labels.font.size = font_size
    else:
        labels.font.size = Pt(float(font_size))

    label_position = normalize_label_position(data_cfg.get("position"))
    if label_position is not None:
        labels.position = label_position
    if hasattr(labels, "show_value"):
        labels.show_value = True
    color_value = data_cfg.get("color")
    if color_value is not None:
        apply_color(labels.font.color, color_value)


def chart_box_from_spec(raw) -> tuple[int, int, int, int] | None:
    if not raw:
        return None
    if isinstance(raw, dict):
        values = [raw.get("x"), raw.get("y"), raw.get("cx"), raw.get("cy")]
    else:
        values = list(raw)
    if len(values) != 4:
        return None

    emu_box: list[int] = []
    for value in values:
        emu_value = coerce_emu(value)
        if emu_value is None:
            return None
        emu_box.append(int(emu_value))

    return (emu_box[0], emu_box[1], emu_box[2], emu_box[3])


def template_content_box(slide, template_path: Path | None) -> tuple | None:
    if template_path is None:
        return None
    try:
        if template_path.resolve() == CLEAN_SLIDES_TEMPLATE_PATH.resolve():
            return CLEAN_SLIDES_CONTENT_BOX
    except FileNotFoundError:
        if template_path == CLEAN_SLIDES_TEMPLATE_PATH:
            return CLEAN_SLIDES_CONTENT_BOX

    placeholder = find_content_placeholder(slide)
    if placeholder is not None:
        return (placeholder.left, placeholder.top, placeholder.width, placeholder.height)
    return None


def find_content_placeholder(slide):
    candidates = []
    for placeholder in slide.placeholders:
        ph_type = placeholder.placeholder_format.type
        if ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT):
            candidates.append(placeholder)
    if not candidates:
        return None
    return max(candidates, key=lambda ph: ph.width * ph.height)


def remove_shape(shape) -> None:
    element = shape._element
    element.getparent().remove(element)


def adjust_bar_chart_box_for_overlays(chart_box: tuple, bar_meta: dict | None) -> tuple:
    if not bar_meta:
        return chart_box
    overlay = bar_meta.get("overlay") if bar_meta else None
    if not overlay:
        return chart_box

    legend_offset = overlay.get("legend_label_offset", DEFAULT_BAR_LEGEND_LABEL_OFFSET)
    legend_height = DEFAULT_BAR_LEGEND_LABEL_HEIGHT
    marker_height = overlay.get("legend_marker_height", DEFAULT_BAR_LEGEND_MARKER_HEIGHT)
    marker_y_offset = overlay.get("legend_marker_y_offset", DEFAULT_BAR_LEGEND_MARKER_Y_OFFSET)
    overlay_extra = overlay.get("overlay_band_extra", DEFAULT_BAR_OVERLAY_BAND_EXTRA)

    bottom_band = max(
        0,
        int(legend_offset) + int(legend_height),
        int(legend_offset) + int(marker_y_offset) + int(marker_height),
    ) + int(overlay_extra)
    if bottom_band <= 0:
        return chart_box

    x, y, cx, cy = chart_box
    new_cy = max(Emu(100000), int(cy) - bottom_band)
    return (int(x), int(y), int(cx), int(new_cy))


def find_layout(prs: Presentation, name: str | None):
    if not name:
        return None
    target = name.strip().lower()
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() == target:
            return layout
    return None


def apply_template_placeholders(slide, title: str | None, subtitle: str | None) -> None:
    for placeholder in slide.placeholders:
        if not placeholder.has_text_frame:
            continue
        ph_type = placeholder.placeholder_format.type
        if ph_type == PP_PLACEHOLDER.TITLE:
            placeholder.text = title or ""
        elif ph_type == PP_PLACEHOLDER.SUBTITLE:
            placeholder.text = subtitle or ""
        else:
            placeholder.text = ""


def add_hidden_anchor(slide) -> None:
    # tiny rectangle at top-left as a stand-in for the OLE anchor
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.02),
        Inches(0.02),
        Inches(0.02),
        Inches(0.02),
    )
    shape.name = "chart data - do not delete"
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.fill.background()


def add_overlay_labels(slide, categories, chart_box):
    x, y, cx, cy = chart_box
    label_y = y + cy + Inches(0.1)
    if not categories:
        return
    label_w = cx / len(categories)
    for idx, label in enumerate(categories):
        tb = slide.shapes.add_textbox(x + (label_w * idx), label_y, label_w, Inches(0.3))
        tb.name = f"tc_category_label_{idx}"
        tf = tb.text_frame
        tf.text = str(label)
        p = tf.paragraphs[0]
        p.font.size = Pt(10)


def select_slide(
    prs: Presentation,
    slide_layout,
    use_template: bool,
    spec: dict,
):
    template_slide_index = spec.get("template_slide_index")
    append_slide = spec.get("append_slide", False)

    if use_template and template_slide_index is not None:
        try:
            index = int(template_slide_index) - 1
        except (TypeError, ValueError) as exc:
            raise ValueError("template_slide_index must be an integer (1-based)") from exc
        if index < 0 or index >= len(prs.slides):
            raise ValueError(
                f"template_slide_index {template_slide_index} is out of range (1-{len(prs.slides)})"
            )
        return prs.slides[index]

    if use_template and not append_slide and len(prs.slides) == 1:
        return prs.slides[0]

    # If the deck has exactly one slide with no chart content yet
    # (just the blank template slide), reuse it even when append_slide
    # is set — avoids creating a blank leading slide.
    if use_template and append_slide and len(prs.slides) == 1:
        slide = prs.slides[0]
        has_chart = any(s.has_chart for s in slide.shapes)
        if not has_chart:
            return slide

    return prs.slides.add_slide(slide_layout)


def build_chart(
    prs: Presentation,
    spec: dict,
    output_path: Path,
    template_path: Path | None = None,
    layout_name: str | None = None,
    save: bool = True,
    defer_template_copy: bool = False,
) -> list[ChartTemplateReplacement]:
    slide_layout = find_layout(prs, layout_name)
    if slide_layout is None:
        slide_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    use_template = template_path is not None
    slide = select_slide(prs, slide_layout, use_template, spec)

    add_hidden_anchor(slide)

    chart_type_key = spec.get("type", "clustered")
    is_waterfall = chart_type_key in WATERFALL_TYPES
    if is_waterfall:
        chart_type, chart_data, style = build_waterfall_payload(spec)
    else:
        chart_type, chart_data, style = build_bar_payload(spec)

    title = spec.get("title")
    subtitle = spec.get("subtitle")
    use_template = template_path is not None
    content_placeholder = None
    remove_placeholder = spec.get("remove_content_placeholder")
    if remove_placeholder is None:
        remove_placeholder = use_template
    if use_template:
        apply_template_placeholders(slide, title, subtitle)
        content_placeholder = find_content_placeholder(slide)

    default_box = (
        DEFAULT_WATERFALL_CHART_BOX
        if is_waterfall
        else (Inches(0.5), Inches(1.0), Inches(9.0), Inches(4.5))
    )
    chart_box = chart_box_from_spec(spec.get("chart_box")) or default_box
    wf_chart_box = (spec.get("waterfall") or {}).get("chart_box") if is_waterfall else None
    if is_waterfall:
        chart_box = chart_box_from_spec(wf_chart_box) or chart_box
    template_box = template_content_box(slide, template_path)
    if content_placeholder is not None and remove_placeholder:
        remove_shape(content_placeholder)
    if template_box and wf_chart_box is None and spec.get("chart_box") is None:
        chart_box = template_box
        if style.get("bar") and spec.get("add_overlay_labels", False):
            chart_box = adjust_bar_chart_box_for_overlays(chart_box, style.get("bar"))

    x, y, cx, cy = chart_box
    chart = slide.shapes.add_chart(chart_type, x, y, cx, cy, chart_data).chart
    chart_part = chart.part.partname.lstrip("/")

    if title and not is_waterfall and not use_template:
        chart.has_title = True
        chart.chart_title.text_frame.text = title

    show_legend = spec.get("show_legend")
    if show_legend is None:
        show_legend = not (is_waterfall or style.get("bar"))
    chart.has_legend = bool(show_legend)
    if chart.has_legend:
        chart.legend.include_in_layout = False

    if is_waterfall:
        show_segment_labels = spec.get("show_data_labels", True)
        if show_segment_labels:
            apply_waterfall_data_labels(chart, style.get("waterfall", {}))
    elif spec.get("show_data_labels", False):
        data_cfg = spec.get("data_labels") or {}
        series_selector = (
            data_cfg.get("series_indices") or data_cfg.get("series") or data_cfg.get("series_names")
        )
        if series_selector:
            series_names = [series.name for series in chart.series]
            indices = resolve_series_indices(series_selector, series_names)
            for idx in indices:
                if 0 <= idx < len(chart.series):
                    series = chart.series[idx]
                    series.has_data_labels = True
                    apply_data_label_style(series.data_labels, data_cfg)
        else:
            plot = chart.plots[0]
            plot.has_data_labels = True
            apply_data_label_style(plot.data_labels, data_cfg)

    apply_series_colors(chart, style.get("series_colors", []))
    if style.get("waterfall"):
        apply_waterfall_style(chart, style["waterfall"])
        apply_waterfall_chart_style(chart, style["waterfall"])
        if is_waterfall and spec.get("show_data_labels", True):
            apply_waterfall_data_label_layout(
                chart,
                (int(x), int(y), int(cx), int(cy)),
                style["waterfall"],
            )
    if style.get("bar"):
        apply_bar_chart_style(chart, style["bar"])
        bar_template = style["bar"].get("chart_template")
        if bar_template:
            apply_chart_template_dlbls(
                chart,
                Path(bar_template),
                slide_index=int(style["bar"].get("chart_template_slide", 1)) - 1,
                chart_index=int(style["bar"].get("chart_template_chart_index", 0)),
                series_index=int(style["bar"].get("chart_template_series_index", 0)),
            )

    overlay_title = (spec.get("waterfall") or {}).get("overlay_title")
    if overlay_title is None:
        overlay_title = not use_template

    if is_waterfall and title and overlay_title:
        title_offset = (spec.get("waterfall") or {}).get("title_offset")
        if title_offset is None:
            title_offset = DEFAULT_WATERFALL_TITLE_OFFSET
        else:
            title_offset = Inches(float(title_offset))
        add_waterfall_title(slide, (int(x), int(y), int(cx), int(cy)), title, title_offset)

    if spec.get("add_overlay_labels", False):
        chart_box_emu = (int(x), int(y), int(cx), int(cy))
        if is_waterfall:
            add_waterfall_overlays(
                slide,
                chart_box_emu,
                style.get("waterfall", {}),
                slide_size=(prs.slide_width, prs.slide_height),
            )
        elif style.get("bar"):
            add_bar_overlays(slide, chart_box_emu, style.get("bar", {}))
        else:
            add_overlay_labels(slide, spec.get("categories", []), (x, y, cx, cy))

    replacements: list[ChartTemplateReplacement] = []
    bar_meta = style.get("bar") or {}
    if bar_meta.get("chart_template_copy"):
        chart_template = bar_meta.get("chart_template")
        if chart_template:
            replacements.append(
                ChartTemplateReplacement(
                    chart_part=chart_part,
                    template_path=Path(chart_template),
                    template_slide_index=int(bar_meta.get("chart_template_slide", 1)) - 1,
                    template_chart_index=int(bar_meta.get("chart_template_chart_index", 0)),
                )
            )

    if save:
        prs.save(output_path)
        if replacements and not defer_template_copy:
            apply_chart_template_replacements(output_path, replacements)
            replacements = []

    return replacements
