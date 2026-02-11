"""Chart annotation and text-shape helpers."""

# pyright: reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false
# pyright: reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false
# pyright: reportMissingParameterType=false
# pyright: reportMissingTypeArgument=false
# pyright: reportGeneralTypeIssues=false
# pyright: reportPrivateUsage=false
# pyright: reportArgumentType=false
# pyright: reportCallIssue=false
# pyright: reportAttributeAccessIssue=false
# pyright: reportUnnecessaryIsInstance=false

from __future__ import annotations

from typing import Any

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.oxml.ns import _nsmap, qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

from .colors import apply_color, normalize_theme_color, resolve_color
from .defaults import (
    DEFAULT_BAR_SEGMENT_LABEL_FONT_SIZE,
    DEFAULT_WATERFALL_LABEL_FONT_SIZE,
    DEFAULT_WATERFALL_TITLE_FONT_SIZE,
    DEFAULT_WATERFALL_TITLE_HEIGHT,
)
from .text_style import normalize_alignment, normalize_vertical_anchor
from .text_templates import apply_txbody_template, resolve_txbody_template
from .units import coerce_emu, coerce_line_width, resolve_path

_NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
if "a14" not in _nsmap:
    _nsmap["a14"] = _NS_A14

_DEFAULT_TEXT_COLOR = RGBColor(0, 0, 0)


def add_text_label(
    slide,
    text: str,
    x: float,
    y: float,
    width: float,
    height: float,
    align: PP_ALIGN = PP_ALIGN.CENTER,
    color: RGBColor | str | None = _DEFAULT_TEXT_COLOR,
    font_size: Pt = DEFAULT_WATERFALL_LABEL_FONT_SIZE,
    fill_color: RGBColor | str | None = None,
    shape_type: MSO_SHAPE | None = None,
    margin_left: int | float | None = None,
    margin_right: int | float | None = None,
    margin_top: int | float | None = None,
    margin_bottom: int | float | None = None,
    vertical_anchor: MSO_VERTICAL_ANCHOR | str | None = None,
    bold: bool | None = None,
    line_color: RGBColor | str | None = None,
    line_width: int | float | None = None,
    bw_mode: str | None = None,
    txbody_template: OxmlElement | None = None,
) -> Any:
    x_int = round(x)
    y_int = round(y)
    w_int = round(width)
    h_int = round(height)
    if shape_type is None:
        box = slide.shapes.add_textbox(x_int, y_int, w_int, h_int)
    else:
        box = slide.shapes.add_shape(shape_type, x_int, y_int, w_int, h_int)
    box.text_frame.text = text
    box.text_frame.word_wrap = False
    box.text_frame.auto_size = MSO_AUTO_SIZE.NONE

    if vertical_anchor is not None:
        anchor = (
            normalize_vertical_anchor(vertical_anchor)
            if isinstance(vertical_anchor, str)
            else vertical_anchor
        )
        if anchor is not None:
            box.text_frame.vertical_anchor = anchor

    if margin_left is not None:
        box.text_frame.margin_left = coerce_emu(margin_left) or 0
    if margin_right is not None:
        box.text_frame.margin_right = coerce_emu(margin_right) or 0
    if margin_top is not None:
        box.text_frame.margin_top = coerce_emu(margin_top) or 0
    if margin_bottom is not None:
        box.text_frame.margin_bottom = coerce_emu(margin_bottom) or 0

    p = box.text_frame.paragraphs[0]
    p.alignment = align
    if isinstance(font_size, Pt):
        p.font.size = font_size
    else:
        p.font.size = Pt(float(font_size))
    if bold is not None:
        p.font.bold = bool(bold)
    if color is not None:
        apply_color(p.font.color, color)

    if fill_color is None:
        box.fill.background()
    else:
        box.fill.solid()
        if not apply_color(box.fill.fore_color, fill_color):
            box.fill.background()

    if line_color is None:
        box.line.fill.background()
    else:
        box.line.fill.solid()
        apply_color(box.line.color, line_color)
        if line_width is not None:
            lw = coerce_line_width(line_width)
            if lw is not None:
                box.line.width = lw

    resolved_bw_mode = bw_mode
    if resolved_bw_mode is None and shape_type is None:
        resolved_bw_mode = "gray" if fill_color is not None else "auto"
    if resolved_bw_mode:
        sp_pr = box._element.find(qn("p:spPr"))
        if sp_pr is not None:
            sp_pr.set("bwMode", resolved_bw_mode)

    if shape_type is None:
        sp_pr = box._element.find(qn("p:spPr"))
        if sp_pr is not None:
            prst_geom = sp_pr.find(qn("a:prstGeom"))
            if prst_geom is not None:
                prst_geom.set("prst", "rect")

            ext_uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}"
            ext_lst = sp_pr.find(qn("a:extLst"))
            if fill_color is None:
                if ext_lst is None:
                    ext_lst = OxmlElement("a:extLst")
                    sp_pr.append(ext_lst)
                existing = None
                for child in ext_lst:
                    if child.tag == qn("a:ext") and child.get("uri") == ext_uri:
                        existing = child
                        break
                if existing is None:
                    ext = OxmlElement("a:ext")
                    ext.set("uri", ext_uri)
                    hidden_fill = OxmlElement("a14:hiddenFill")
                    solid = OxmlElement("a:solidFill")
                    scheme = OxmlElement("a:schemeClr")
                    scheme.set("val", "accent1")
                    solid.append(scheme)
                    hidden_fill.append(solid)
                    ext.append(hidden_fill)
                    ext_lst.append(ext)

            effect_lst = sp_pr.find(qn("a:effectLst"))
            if effect_lst is None:
                effect_lst = OxmlElement("a:effectLst")
                if ext_lst is not None:
                    sp_pr.insert(sp_pr.index(ext_lst), effect_lst)
                else:
                    sp_pr.append(effect_lst)

    if txbody_template is not None:
        apply_txbody_template(box, txbody_template, text)

    return box


def set_line_endings(line, head: dict | None = None, tail: dict | None = None) -> None:
    ln = line._get_or_add_ln()

    def apply_end(tag: str, data: dict) -> None:
        element = ln.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            ln.append(element)
        element.attrib.clear()
        element.set("type", data.get("type", "none"))
        element.set("w", data.get("w", "med"))
        element.set("len", data.get("len", "med"))

    if head is not None:
        apply_end("a:headEnd", head)
    if tail is not None:
        apply_end("a:tailEnd", tail)


def add_line_annotation(slide, spec: dict) -> None:
    x = coerce_emu(spec.get("x"))
    y = coerce_emu(spec.get("y"))
    raw_width = spec.get("w") if spec.get("w") is not None else spec.get("width")
    raw_height = spec.get("h") if spec.get("h") is not None else spec.get("height")
    width = coerce_emu(raw_width)
    height = coerce_emu(raw_height)
    if None in (x, y, width, height):
        return
    x_int = round(x)
    y_int = round(y)
    w_int = round(width)
    h_int = round(height)
    # python-pptx skips zero-length connectors; create a minimal line then patch extents.
    line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        x_int,
        y_int,
        x_int + (w_int or 1),
        y_int + (h_int or 1),
    )
    sp_pr = line._element.find(qn("p:spPr"))
    if sp_pr is not None:
        sp_pr.set("bwMode", "auto")
        xfrm = sp_pr.find(qn("a:xfrm"))
        if xfrm is not None:
            off = xfrm.find(qn("a:off"))
            ext = xfrm.find(qn("a:ext"))
            if off is not None:
                off.set("x", str(x_int))
                off.set("y", str(y_int))
            if ext is not None:
                ext.set("cx", str(w_int))
                ext.set("cy", str(h_int))
    line_width = spec.get("line_width")
    if line_width is not None:
        lw = coerce_line_width(line_width)
        if lw is not None:
            line.line.width = lw
    line_color = spec.get("line_color")
    if line_color is not None:
        line.line.fill.solid()
        apply_color(line.line.color, line_color)
    dash = spec.get("dash_style")
    if dash:
        dash_map = {
            "solid": MSO_LINE_DASH_STYLE.SOLID,
            "dash": MSO_LINE_DASH_STYLE.DASH,
            "long_dash": MSO_LINE_DASH_STYLE.LONG_DASH,
            "dot": MSO_LINE_DASH_STYLE.ROUND_DOT,
        }
        dash_key = str(dash).lower()
        dash_style = dash_map.get(dash_key)
        if dash_style is not None:
            line.line.dash_style = dash_style
        # Ensure prstDash node exists for solid styles.
        if dash_key in {"solid", "dash", "long_dash", "dot"}:
            ln = line.line._get_or_add_ln()
            prst = ln.find(qn("a:prstDash"))
            if prst is None:
                prst = OxmlElement("a:prstDash")
                ln.append(prst)
            prst.set("val", "solid" if dash_key == "solid" else dash_key)

    if spec.get("round"):
        ln = line.line._get_or_add_ln()
        if ln.find(qn("a:round")) is None:
            ln.append(OxmlElement("a:round"))

    head = spec.get("head_end")
    tail = spec.get("tail_end")
    if head or tail:
        set_line_endings(line.line, head=head, tail=tail)

    cap = spec.get("cap")
    if cap:
        ln = line.line._get_or_add_ln()
        ln.set("cap", cap)
    cmpd = spec.get("cmpd")
    if cmpd:
        ln = line.line._get_or_add_ln()
        ln.set("cmpd", cmpd)
    algn = spec.get("algn")
    if algn:
        ln = line.line._get_or_add_ln()
        ln.set("algn", algn)

    if spec.get("flip_v"):
        xfrm = line._element.find(qn("p:spPr"))
        if xfrm is not None:
            xfrm = xfrm.find(qn("a:xfrm"))
        if xfrm is not None:
            xfrm.set("flipV", "1")

    if spec.get("flip_h"):
        xfrm = line._element.find(qn("p:spPr"))
        if xfrm is not None:
            xfrm = xfrm.find(qn("a:xfrm"))
        if xfrm is not None:
            xfrm.set("flipH", "1")

    sp_pr = line._element.find(qn("p:spPr"))
    if sp_pr is not None and sp_pr.find(qn("a:effectLst")) is None:
        sp_pr.append(OxmlElement("a:effectLst"))


def add_shape_annotation(slide, spec: dict) -> None:
    shape_type = spec.get("shape")
    if shape_type == "ellipse":
        shape_enum = MSO_SHAPE.OVAL
    elif shape_type == "rectangle":
        shape_enum = MSO_SHAPE.RECTANGLE
    else:
        shape_enum = None

    x = coerce_emu(spec.get("x"))
    y = coerce_emu(spec.get("y"))
    raw_width = spec.get("w") if spec.get("w") is not None else spec.get("width")
    raw_height = spec.get("h") if spec.get("h") is not None else spec.get("height")
    width = coerce_emu(raw_width)
    height = coerce_emu(raw_height)
    if None in (x, y, width, height):
        return

    # For ellipses, mirror the text box + ellipse geometry from sample exports.
    use_text_box = shape_type == "ellipse"
    font_size_value = spec.get("font_size", DEFAULT_BAR_SEGMENT_LABEL_FONT_SIZE)
    if isinstance(font_size_value, Pt):
        font_size = font_size_value
    else:
        font_size = Pt(float(font_size_value))

    txbody_template = None
    base_dir = spec.get("_base_dir")
    text_style_template = spec.get("text_style_template")
    if isinstance(text_style_template, str) and spec.get("text"):
        template_path = resolve_path(text_style_template, base_dir)
        txbody_template = resolve_txbody_template(template_path, spec.get("text", ""), None)

    text_value = spec.get("text", "")
    shape = add_text_label(
        slide,
        text_value,
        x,
        y,
        width,
        height,
        align=normalize_alignment(spec.get("align")) or PP_ALIGN.CENTER,
        color=spec.get("text_color"),
        font_size=font_size,
        fill_color=spec.get("fill_color"),
        shape_type=None if use_text_box else shape_enum,
        margin_left=spec.get("margin_left"),
        margin_right=spec.get("margin_right"),
        margin_top=spec.get("margin_top"),
        margin_bottom=spec.get("margin_bottom"),
        vertical_anchor=spec.get("vertical_anchor"),
        bold=spec.get("bold"),
        line_color=spec.get("line_color"),
        line_width=spec.get("line_width"),
    )

    if shape is not None and txbody_template is not None:
        template_text = "".join(t_elem.text or "" for t_elem in txbody_template.iter(qn("a:t")))
        override_text = None if template_text == text_value else text_value
        apply_txbody_template(shape, txbody_template, override_text)

    if use_text_box and shape is not None:
        sp_pr = shape._element.find(qn("p:spPr"))
        if sp_pr is not None:
            sp_pr.set("bwMode", "auto")
            prst_geom = sp_pr.find(qn("a:prstGeom"))
            if prst_geom is not None:
                prst_geom.set("prst", "ellipse")

            line_color = spec.get("line_color")
            if line_color is not None:
                ln = sp_pr.find(qn("a:ln"))
                if ln is None:
                    ln = OxmlElement("a:ln")
                    sp_pr.append(ln)
                if spec.get("line_width") is not None:
                    lw = coerce_line_width(spec.get("line_width"))
                    if lw is not None:
                        ln.set("w", str(lw))
                if spec.get("cmpd"):
                    ln.set("cmpd", spec.get("cmpd"))
                for child in list(ln):
                    if child.tag == qn("a:solidFill"):
                        ln.remove(child)
                no_fill = ln.find(qn("a:noFill"))
                if no_fill is None:
                    ln.append(OxmlElement("a:noFill"))

                ext_uri = "{91240B29-F687-4F45-9708-019B960494DF}"
                ext_lst = sp_pr.find(qn("a:extLst"))
                if ext_lst is None:
                    ext_lst = OxmlElement("a:extLst")
                    sp_pr.append(ext_lst)
                ext = None
                for child in ext_lst:
                    if child.tag == qn("a:ext") and child.get("uri") == ext_uri:
                        ext = child
                        break
                if ext is None:
                    ext = OxmlElement("a:ext")
                    ext.set("uri", ext_uri)
                    ext_lst.append(ext)

                hidden_line = ext.find(qn("a14:hiddenLine"))
                if hidden_line is None:
                    hidden_line = OxmlElement("a14:hiddenLine")
                    ext.append(hidden_line)
                if spec.get("line_width") is not None:
                    lw = coerce_line_width(spec.get("line_width"))
                    if lw is not None:
                        hidden_line.set("w", str(lw))
                if spec.get("cmpd"):
                    hidden_line.set("cmpd", spec.get("cmpd"))

                solid = hidden_line.find(qn("a:solidFill"))
                if solid is None:
                    solid = OxmlElement("a:solidFill")
                    hidden_line.append(solid)
                if isinstance(line_color, str) and normalize_theme_color(line_color) is not None:
                    scheme = solid.find(qn("a:schemeClr"))
                    if scheme is None:
                        scheme = OxmlElement("a:schemeClr")
                        solid.append(scheme)
                    scheme.set("val", line_color)
                else:
                    rgb, theme = resolve_color(line_color)
                    if theme is not None:
                        scheme = solid.find(qn("a:schemeClr"))
                        if scheme is None:
                            scheme = OxmlElement("a:schemeClr")
                            solid.append(scheme)
                        scheme.set("val", theme)
                    elif rgb is not None:
                        srgb = solid.find(qn("a:srgbClr"))
                        if srgb is None:
                            srgb = OxmlElement("a:srgbClr")
                            solid.append(srgb)
                        srgb.set("val", rgb.hex)


def add_waterfall_title(slide, chart_box: tuple, title: str, offset: float) -> None:
    x, y, cx, _ = chart_box
    title_y = y - offset
    add_text_label(
        slide,
        title,
        x,
        title_y,
        cx,
        DEFAULT_WATERFALL_TITLE_HEIGHT,
        align=PP_ALIGN.CENTER,
        font_size=DEFAULT_WATERFALL_TITLE_FONT_SIZE,
    )
