"""
XML helpers for generating proper OOXML for PowerPoint.

Two categories of utilities:
  1. Shape generation — create lstStyle, moons, lines from scratch
  2. XML mutation — idempotent find-or-create, safe removal (prevents duplicate elements)
"""

# pyright: reportPrivateUsage=false

from __future__ import annotations

from lxml import etree
from pptx.oxml.ns import qn
from pptx.shapes.base import BaseShape

from .constants import (
    BULLET_BU_FONT,
    BULLET_DEF_RPR,
    BULLET_DEF_TAB_SZ_EMU,
    BULLET_LEVELS,
    MOON_ARC_ADJUSTMENTS,
    MOON_COLORS_HEX,
    MOON_FILLS,
    MOON_GROUP,
    MOON_SIZE_EMU,
)

# OOXML namespace URIs
NSMAP: dict[str, str] = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def _el(tag: str, **attribs: str) -> etree._Element:
    """Create an element with namespace prefix"""
    ns, local = tag.split(":")
    return etree.Element(f"{{{NSMAP[ns]}}}{local}", attribs)


def _sub(parent: etree._Element, tag: str, **attribs: str) -> etree._Element:
    """Create and append a subelement"""
    el = _el(tag, **attribs)
    parent.append(el)
    return el


# ============================================================================
# TEXT PLACEHOLDER XML
# ============================================================================


def create_placeholder_lstStyle_xml() -> etree._Element:
    """
    Create the lstStyle XML with all 9 bullet levels matching the template.
    This is the key to proper bullet formatting.
    """
    lstStyle = _el("a:lstStyle")

    for level in BULLET_LEVELS:
        lvl = int(level["level"])
        marL = int(level["mar_l_emu"])
        indent = int(level["indent_emu"])
        spc_bef = int(float(level["spc_bef_pt"]) * 100)
        spc_aft = int(float(level["spc_aft_pt"]) * 100)
        bullet = level.get("bullet", "")
        lnSpc = int(level["ln_spc_pct"])

        pPr = _sub(
            lstStyle,
            f"a:lvl{lvl}pPr",
            marL=str(marL),
            indent=str(indent),
            algn="l",
            defTabSz=str(BULLET_DEF_TAB_SZ_EMU),
            rtl="0",
            eaLnBrk="1",
            latinLnBrk="0",
            hangingPunct="1",
        )

        # Line spacing
        lnSpcEl = _sub(pPr, "a:lnSpc")
        _sub(lnSpcEl, "a:spcPct", val=str(lnSpc))

        # Space before
        spcBefEl = _sub(pPr, "a:spcBef")
        _sub(spcBefEl, "a:spcPts", val=str(spc_bef))

        # Space after
        spcAft = _sub(pPr, "a:spcAft")
        _sub(spcAft, "a:spcPts", val=str(spc_aft))

        # Bullet
        if bullet:
            _sub(
                pPr,
                "a:buFont",
                typeface=BULLET_BU_FONT["typeface"],
                pitchFamily=BULLET_BU_FONT["pitch_family"],
                charset=BULLET_BU_FONT["charset"],
            )
            _sub(pPr, "a:buChar", char=bullet)
        else:
            _sub(pPr, "a:buNone")

        # Default run properties
        def_sz = int(float(BULLET_DEF_RPR["size_pt"]) * 100)
        def_kern = int(float(BULLET_DEF_RPR["kern_pt"]) * 100)
        defRPr = _sub(pPr, "a:defRPr", sz=str(def_sz), kern=str(def_kern))
        solidFill = _sub(defRPr, "a:solidFill")
        _sub(solidFill, "a:schemeClr", val=BULLET_DEF_RPR["scheme"])
        _sub(defRPr, "a:latin", typeface=BULLET_DEF_RPR["latin"])
        _sub(
            defRPr,
            "a:cs",
            typeface=BULLET_DEF_RPR["cs"],
            pitchFamily=BULLET_DEF_RPR["cs_pitch_family"],
            charset=BULLET_DEF_RPR["cs_charset"],
        )

    return lstStyle


def create_text_placeholder_bodyPr(anchor: str = "t") -> etree._Element:
    """Create bodyPr with zero insets and no autofit"""
    bodyPr = _el(
        "a:bodyPr", vert="horz", lIns="0", tIns="0", rIns="0", bIns="0", rtlCol="0", anchor=anchor
    )
    _sub(bodyPr, "a:noAutofit")
    return bodyPr


# ============================================================================
# MOON (RAG INDICATOR) XML
# ============================================================================


def create_moon_group_xml(
    x: int,
    y: int,
    level: str,
    size: int = MOON_SIZE_EMU,
) -> etree._Element:
    """
    Create a moon (RAG indicator) group XML element.

    The moon is a group containing:
    1. Background oval (with group fill)
    2. Arc shape (with group fill) - controls the fill amount

    The fill color is set on the group, and both shapes inherit via grpFill.
    """
    fill_pct = MOON_FILLS.get(level.lower() if level else "none", 0)
    color = MOON_COLORS_HEX.get(level.lower() if level else "none")

    # Child coordinates (fixed reference frame from template MoonLegend5 group)
    ch_off_x = int(MOON_GROUP["child_offset_x"])
    ch_off_y = int(MOON_GROUP["child_offset_y"])
    ch_size = int(MOON_GROUP["child_size"])

    # Create group
    grpSp = _el("p:grpSp")

    # Group properties
    nvGrpSpPr = _sub(grpSp, "p:nvGrpSpPr")
    _sub(nvGrpSpPr, "p:cNvPr", id="1", name="Moon")
    _sub(nvGrpSpPr, "p:cNvGrpSpPr")
    grpSpLocks = _el("a:grpSpLocks", noChangeAspect="1")
    nvGrpSpPr[-1].append(grpSpLocks)
    _sub(nvGrpSpPr, "p:nvPr")

    # Group shape properties (with fill color)
    grpSpPr = _sub(grpSp, "p:grpSpPr")
    xfrm = _sub(grpSpPr, "a:xfrm")
    _sub(xfrm, "a:off", x=str(x), y=str(y))
    _sub(xfrm, "a:ext", cx=str(size), cy=str(size))
    _sub(xfrm, "a:chOff", x=str(ch_off_x), y=str(ch_off_y))
    _sub(xfrm, "a:chExt", cx=str(ch_size), cy=str(ch_size))

    # Set fill color on group
    if color:
        solidFill = _sub(grpSpPr, "a:solidFill")
        _sub(solidFill, "a:srgbClr", val=color)

    # Create background oval
    is_full = fill_pct == 100
    oval = _create_moon_oval(ch_off_x, ch_off_y, ch_size, is_full)
    grpSp.append(oval)

    # Create arc (only if partial fill, not needed for 100% or 0%)
    arc_adj = MOON_ARC_ADJUSTMENTS.get(fill_pct)
    if arc_adj is not None and fill_pct > 0:
        arc = _create_moon_arc(ch_off_x, ch_off_y, ch_size, fill_pct)
        grpSp.append(arc)

    return grpSp


def _create_moon_oval(x: int, y: int, size: int, filled: bool) -> etree._Element:
    """Create the background oval for a moon"""
    sp = _el("p:sp")

    # Non-visual properties
    nvSpPr = _sub(sp, "p:nvSpPr")
    _sub(nvSpPr, "p:cNvPr", id="2", name="Oval")
    _sub(nvSpPr, "p:cNvSpPr")
    _sub(nvSpPr, "p:nvPr")

    # Shape properties
    spPr = _sub(sp, "p:spPr")
    xfrm = _sub(spPr, "a:xfrm")
    _sub(xfrm, "a:off", x=str(x), y=str(y))
    _sub(xfrm, "a:ext", cx=str(size), cy=str(size))

    prstGeom = _sub(spPr, "a:prstGeom", prst="ellipse")
    _sub(prstGeom, "a:avLst")

    # Fill: inherit from group if filled, otherwise no fill with outline
    line_width = str(int(MOON_GROUP["line_width_emu"]))

    if filled:
        _sub(spPr, "a:grpFill")
        ln = _sub(spPr, "a:ln", w=line_width, cap="flat", cmpd="sng", algn="ctr")
        _sub(ln, "a:noFill")
    else:
        # Light background fill
        solidFill = _sub(spPr, "a:solidFill")
        _sub(solidFill, "a:srgbClr", val=str(MOON_GROUP["bg_fill"]).lstrip("#"))
        ln = _sub(spPr, "a:ln", w=line_width, cap="flat", cmpd="sng", algn="ctr")
        lnFill = _sub(ln, "a:solidFill")
        _sub(lnFill, "a:schemeClr", val=MOON_GROUP["outline_scheme"])
        _sub(ln, "a:prstDash", val="solid")
        _sub(ln, "a:round")

    _sub(spPr, "a:effectLst")

    # Empty text body (required)
    txBody = _sub(sp, "p:txBody")
    _sub(txBody, "a:bodyPr", rtlCol="0", anchor="ctr")
    _sub(txBody, "a:lstStyle")
    p = _sub(txBody, "a:p")
    _sub(p, "a:pPr", algn="ctr")
    _sub(p, "a:endParaRPr", lang="en-US")

    return sp


def _create_moon_arc(x: int, y: int, size: int, fill_pct: int) -> etree._Element:
    """Create the arc shape for partial moon fill"""
    arc_adj = MOON_ARC_ADJUSTMENTS.get(fill_pct)
    if arc_adj is None:
        arc_adj = ("16200000", "16200000")
    adj1, adj2 = arc_adj

    sp = _el("p:sp")

    # Non-visual properties
    nvSpPr = _sub(sp, "p:nvSpPr")
    _sub(nvSpPr, "p:cNvPr", id="3", name="Arc")
    _sub(nvSpPr, "p:cNvSpPr")
    _sub(nvSpPr, "p:nvPr")

    # Shape properties
    spPr = _sub(sp, "p:spPr")
    xfrm = _sub(spPr, "a:xfrm")
    _sub(xfrm, "a:off", x=str(x), y=str(y))
    _sub(xfrm, "a:ext", cx=str(size), cy=str(size))

    prstGeom = _sub(spPr, "a:prstGeom", prst="arc")
    avLst = _sub(prstGeom, "a:avLst")
    _sub(avLst, "a:gd", name="adj1", fmla=f"val {adj1}")
    _sub(avLst, "a:gd", name="adj2", fmla=f"val {adj2}")

    # Fill from group
    _sub(spPr, "a:grpFill")

    # No line
    line_width = str(int(MOON_GROUP["line_width_emu"]))
    ln = _sub(spPr, "a:ln", w=line_width, cap="flat", cmpd="sng", algn="ctr")
    _sub(ln, "a:noFill")

    _sub(spPr, "a:effectLst")

    # Empty text body
    txBody = _sub(sp, "p:txBody")
    _sub(txBody, "a:bodyPr", rtlCol="0", anchor="ctr")
    _sub(txBody, "a:lstStyle")
    p = _sub(txBody, "a:p")
    _sub(p, "a:pPr", algn="ctr")
    _sub(p, "a:endParaRPr", lang="en-US")

    return sp


# ============================================================================
# CONNECTOR (LINE) XML
# ============================================================================


def create_line_xml(
    x: int,
    y: int,
    width: int,
    weight_emu: int = 6350,  # 0.5pt (= Pt(0.5) in EMU)
    color_scheme: str = "tx1",  # 'tx1' = Midnight
) -> etree._Element:
    """Create a horizontal line connector"""
    cxnSp = _el("p:cxnSp")

    # Non-visual properties
    nvCxnSpPr = _sub(cxnSp, "p:nvCxnSpPr")
    _sub(nvCxnSpPr, "p:cNvPr", id="1", name="Line")
    _sub(nvCxnSpPr, "p:cNvCxnSpPr")
    _sub(nvCxnSpPr, "p:nvPr")

    # Shape properties
    spPr = _sub(cxnSp, "p:spPr")
    xfrm = _sub(spPr, "a:xfrm")
    _sub(xfrm, "a:off", x=str(x), y=str(y))
    _sub(xfrm, "a:ext", cx=str(width), cy="0")

    prstGeom = _sub(spPr, "a:prstGeom", prst="line")
    _sub(prstGeom, "a:avLst")

    ln = _sub(spPr, "a:ln", w=str(weight_emu))
    solidFill = _sub(ln, "a:solidFill")
    _sub(solidFill, "a:schemeClr", val=color_scheme)

    return cxnSp


# ============================================================================
# XML MUTATION UTILITIES (idempotent, prevents duplicate elements)
# ============================================================================


def find_or_create(parent: etree._Element, tag: str) -> etree._Element:
    """
    Find existing child or create it. Tag uses short form: 'c:autoTitleDeleted'.

        atd = find_or_create(chart_el, 'c:autoTitleDeleted')
        atd.set('val', '1')
    """
    full_tag = qn(tag)
    el = parent.find(full_tag)
    if el is None:
        el = etree.SubElement(parent, full_tag)
    return el


def set_child_val(parent: etree._Element, tag: str, val: object) -> etree._Element:
    """
    Find-or-create child element and set its 'val' attribute.

        set_child_val(chart_el, 'c:autoTitleDeleted', '1')
    """
    el = find_or_create(parent, tag)
    el.set("val", str(val))
    return el


def remove_children(parent: etree._Element, *tags: str) -> None:
    """
    Remove all children matching any of the given tags.

        remove_children(spPr, 'a:noFill', 'a:solidFill', 'a:gradFill')
    """
    for tag in tags:
        full_tag = qn(tag)
        for child in parent.findall(full_tag):
            parent.remove(child)


def remove_noFill(spPr: etree._Element) -> None:
    """Remove noFill from an spPr element."""
    remove_children(spPr, "a:noFill")


def clear_children(parent: etree._Element) -> None:
    """Remove all children from an element."""
    for child in list(parent):
        parent.remove(child)


def remove_shape_style(shape: BaseShape) -> None:
    """
    Remove p:style element that python-pptx auto-adds to shapes.
    This causes default shadows/effects that don't match clean templates.
    """
    sp: etree._Element = shape.element
    style = sp.find(qn("p:style"))
    if style is not None:
        sp.remove(style)


def dump_xml(element: etree._Element, pretty: bool = True) -> str:
    """
    Return XML string for an element. Useful for debugging.

        print(dump_xml(shape._element))
    """
    return etree.tostring(element, pretty_print=pretty).decode("utf-8")
