"""
Table renderer — converts a TableLayout into slide shapes.

Every text cell uses the same text-box structure (lstStyle + noAutofit).
Paragraph level controls bullet style; run properties control font.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from typing import Any, Callable, Protocol

from lxml import etree
from pptx.oxml.shapes.groupshape import CT_GroupShape

from .constants import (
    EMU_PER_INCH,
    OOXML_FONT_SCALE,
    DefaultColors,
    Dividers,
    Fonts,
    FontSizes,
    Layout,
    TableDefaults,
)
from .content import Paragraph, make_sub_paragraph, normalize_cell
from .icons import IconSet, icon_cell_value
from .measure import column_right_pads, should_use_line_breaks, textbox_width
from .spec import Box, CellOverride, ChartRef, ContentArea, TableLayout, TableSpec
from .text_metrics import EMU_PER_PT
from .xml_helpers import (
    create_line_xml,
    create_placeholder_lstStyle_xml,
    create_text_placeholder_bodyPr,
)

# URI for the OOXML hyperlink relationship type.
_HYPERLINK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


XmlElement = Any


class _SlidePart(Protocol):
    """Minimal protocol for a slide part that can register relationships."""

    def relate_to(self, target: str, reltype: str, is_external: bool = False) -> str: ...


# Hex color pattern: optional # prefix, 6 hex digits
_HEX_COLOR_RE = re.compile(r"^#?([0-9A-Fa-f]{6})$")


def _apply_color(parent: XmlElement, color: str) -> None:
    """Add solidFill with either schemeClr (theme) or srgbClr (hex)."""
    fill = etree.SubElement(parent, f"{{{NS_A}}}solidFill")
    hex_match = _HEX_COLOR_RE.match(color)
    if hex_match:
        # Direct RGB hex color
        etree.SubElement(fill, f"{{{NS_A}}}srgbClr", val=hex_match.group(1).upper())
    else:
        # Theme color name (tx1, accent1, etc.)
        etree.SubElement(fill, f"{{{NS_A}}}schemeClr", val=color)


# Inline markdown: **bold**, *italic*, [text](url), (see Appendix)
_INLINE_RE = re.compile(
    r"(\*\*(.+?)\*\*)"
    r"|(\*(.+?)\*)"
    r"|(\[([^\]]+)\]\(([^)]*)\))"
    r"|(\(see Appendix\)|\(See Appendix\))"
)


class TableRenderer:
    """Render table cells and dividers into the slide spTree."""

    spTree: CT_GroupShape
    next_id: Callable[[], int]
    _col_right_pads: list[int]
    _slide_part: _SlidePart | None

    def __init__(
        self,
        spTree: CT_GroupShape,
        next_shape_id_fn: Callable[[], int],
        slide_part: _SlidePart | None = None,
    ) -> None:
        self.spTree = spTree
        self.next_id = next_shape_id_fn
        self._col_right_pads = []
        self._slide_part = slide_part

    def render(self, spec: TableSpec, layout: TableLayout, area: ContentArea) -> None:
        row_offset = spec.row_offset
        col_offset = spec.col_offset
        header_pt = layout.header_font_size // OOXML_FONT_SCALE
        body_pt = layout.body_font_size // OOXML_FONT_SCALE

        # Right padding per column: non-last columns get padding to prevent
        # text crowding. All cells in the same column use the same value so
        # text boxes align. Superheaders use their own (smaller) padding.
        # ~0.10" gap between columns (matches reference styling)
        # Row header column gets extra gap (~0.15") for visual separation
        total_grid_cols = col_offset + spec.num_cols
        self._col_right_pads = column_right_pads(
            total_grid_cols,
            layout.pad_top,
            spec.has_row_header,
        )

        col_hdr_color = spec.effective_col_header_color
        col_super_color = spec.effective_col_superheader_color
        row_hdr_color = spec.effective_row_header_color
        row_super_color = spec.effective_row_superheader_color

        col_header_default = Paragraph(
            text="",
            lvl=0,
            font=Fonts.HEADLINE,
            size_pt=header_pt,
            color=col_hdr_color,
            bold=True,
        )
        col_superheader_default = Paragraph(
            text="",
            lvl=0,
            font=Fonts.HEADLINE,
            size_pt=header_pt,
            color=col_super_color,
            bold=True,
        )
        body_default = Paragraph(
            text="",
            lvl=spec.body_default_lvl,
            font=Fonts.BODY,
            size_pt=body_pt,
            color=DefaultColors.BODY_TEXT,
            bold=False,
        )

        has_promoted_group = bool(spec.groups and any(g.promoted for g in spec.groups))
        # One consistent row-superheader size across grouped tables.
        superheader_pt = body_pt if has_promoted_group else header_pt
        row_superheader_default = Paragraph(
            text="",
            lvl=0,
            font=Fonts.HEADLINE,
            size_pt=superheader_pt,
            color=row_super_color,
            bold=True,
        )

        self._render_col_superheaders(spec, layout, area, col_superheader_default)
        self._render_col_headers(spec, layout, area, col_offset, col_header_default)

        row_header_default = Paragraph(
            text="",
            lvl=0,
            font=Fonts.HEADLINE,
            size_pt=header_pt,
            color=row_hdr_color,
            bold=True,
        )

        if spec.is_grouped:
            self._render_superheaders(
                spec,
                layout,
                row_offset,
                row_superheader_default,
                row_superheader_default,
            )
            self._render_body(spec, layout, row_offset, col_offset, body_default)
            self._render_grouped_dividers(spec, layout, area, row_offset, col_offset)
        else:
            self._render_row_headers(spec, layout, row_offset, row_header_default)
            self._render_body(spec, layout, row_offset, col_offset, body_default)
            self._render_flat_dividers(spec, layout, area, row_offset)

        # Legend for icon indicators
        icons = spec.icons
        if icons is not None and icons.show_legend:
            self._render_legend(icons, area)

    # -----------------------------------------------------------------------
    # Column superheaders (span multiple grid columns)
    # -----------------------------------------------------------------------

    def _render_col_superheaders(
        self, spec: TableSpec, layout: TableLayout, area: ContentArea, default: Paragraph
    ) -> None:
        col_superheaders = spec.col_superheaders
        if not spec.has_col_superheader or col_superheaders is None:
            return

        grid_col = 0
        row_h = layout.row_heights[0]
        line_y = area.y + row_h
        total_grid_cols = len(layout.col_widths)

        for csh in col_superheaders:
            x = area.x + sum(layout.col_widths[:grid_col])
            w = sum(layout.col_widths[grid_col : grid_col + csh.span])
            box: Box = (x, area.y, w, row_h)
            if csh.label:
                last_col_in_span = grid_col + csh.span - 1
                # Non-last superheaders: trim divider to match col header width
                # (subtract rightmost spanned column's right_pad)
                divider_w = w
                if last_col_in_span < total_grid_cols - 1:
                    divider_w = w - self._col_right_pads[last_col_in_span]

                paragraphs = normalize_cell(csh.label, default, parse_bullets=False)
                if csh.sub:
                    paragraphs.append(make_sub_paragraph(csh.sub, paragraphs[0]))
                self._add_cell(
                    box,
                    paragraphs,
                    anchor="b",
                    pad_top=layout.pad_top,
                    pad_bottom=layout.pad_bottom,
                    use_line_breaks=True,
                )
                self._add_line(x, line_y, divider_w, int(Dividers.HEADER), DefaultColors.DIVIDER)
            grid_col += csh.span

    # -----------------------------------------------------------------------
    # Column headers (shared)
    # -----------------------------------------------------------------------

    def _render_col_headers(
        self,
        spec: TableSpec,
        layout: TableLayout,
        area: ContentArea,
        col_offset: int,
        default: Paragraph,
    ) -> None:
        if not spec.has_col_header:
            return

        header_grid_row = 1 if spec.has_col_superheader else 0

        # Row-header column header (e.g. "Header" above superheader column)
        if spec.row_header_col_header and spec.has_row_header:
            paragraphs = normalize_cell(spec.row_header_col_header, default, parse_bullets=False)
            self._add_cell(
                layout.cells[header_grid_row][0],
                paragraphs,
                anchor="b",
                pad_top=layout.pad_top,
                pad_bottom=layout.pad_bottom,
                right_pad=self._col_right_pads[0],
                use_line_breaks=True,
            )

        col_headers = spec.col_headers

        # Body column headers
        for ci in range(spec.num_cols):
            grid_col = ci + col_offset
            text = col_headers[ci] if col_headers and ci < len(col_headers) else ""
            paragraphs = normalize_cell(text, default, parse_bullets=False)
            self._add_cell(
                layout.cells[header_grid_row][grid_col],
                paragraphs,
                anchor="b",
                pad_top=layout.pad_top,
                pad_bottom=layout.pad_bottom,
                right_pad=self._col_right_pads[grid_col],
                use_line_breaks=True,
            )

        header_bottom_y = area.y + sum(layout.row_heights[: header_grid_row + 1])
        self._add_line(
            area.x,
            header_bottom_y,
            sum(layout.col_widths),
            int(Dividers.HEADER),
            DefaultColors.DIVIDER,
        )

    # -----------------------------------------------------------------------
    # Flat row headers (non-grouped tables)
    # -----------------------------------------------------------------------

    def _render_row_headers(
        self,
        spec: TableSpec,
        layout: TableLayout,
        row_offset: int,
        default: Paragraph,
    ) -> None:
        if not spec.has_row_header:
            return

        row_headers = spec.row_headers
        for ri in range(spec.num_rows):
            text = row_headers[ri] if row_headers and ri < len(row_headers) else ""
            paragraphs = normalize_cell(text, default, parse_bullets=False)
            self._add_cell(
                layout.cells[ri + row_offset][0],
                paragraphs,
                pad_top=layout.pad_top,
                pad_bottom=layout.pad_bottom,
                right_pad=self._col_right_pads[0],
                use_line_breaks=True,
            )

    # -----------------------------------------------------------------------
    # Superheaders (grouped tables) — span multiple sub-rows
    # -----------------------------------------------------------------------

    def _render_superheaders(
        self,
        spec: TableSpec,
        layout: TableLayout,
        row_offset: int,
        default: Paragraph,
        promoted_default: Paragraph,
    ) -> None:
        groups = spec.groups
        if groups is None:
            return

        sub_row = 0
        for group in groups:
            grid_row = sub_row + row_offset
            x, y, w, _ = layout.cells[grid_row][0]
            group_h = sum(layout.row_heights[grid_row : grid_row + group.num_rows])

            # Auto-promoted singleton headers can span the first body column,
            # avoiding a visually empty superheader band.
            para_default = default
            right_pad = self._col_right_pads[0]
            if group.promoted and len(layout.col_widths) > 1:
                w += layout.col_widths[1]
                para_default = promoted_default
                right_pad = self._col_right_pads[1]

            paragraphs = normalize_cell(group.header, para_default, parse_bullets=False)
            self._add_cell(
                (x, y, w, group_h),
                paragraphs,
                pad_top=layout.pad_top,
                pad_bottom=layout.pad_bottom,
                right_pad=right_pad,
                use_line_breaks=True,
            )
            sub_row += group.num_rows

    # -----------------------------------------------------------------------
    # Body cells (shared — works for both flat and grouped)
    # -----------------------------------------------------------------------

    @staticmethod
    def _apply_override(paragraphs: list[Paragraph], ov: CellOverride) -> list[Paragraph]:
        """Return a copy of *paragraphs* with override formatting applied."""
        out: list[Paragraph] = []
        for p in paragraphs:
            out.append(
                Paragraph(
                    text=p.text,
                    lvl=p.lvl,
                    font=ov.font if ov.font is not None else p.font,
                    size_pt=ov.size if ov.size is not None else p.size_pt,
                    color=ov.color if ov.color is not None else p.color,
                    bold=ov.bold if ov.bold is not None else p.bold,
                    italic=p.italic,
                    underline=p.underline,
                )
            )
        return out

    def _render_body(
        self,
        spec: TableSpec,
        layout: TableLayout,
        row_offset: int,
        col_offset: int,
        default: Paragraph,
    ) -> None:
        icons = spec.icons
        cells = spec.cells

        for ri in range(spec.num_rows):
            row: Sequence[object]
            if cells is not None and ri < len(cells):
                row = cells[ri]
            else:
                row = []

            for ci in range(spec.num_cols):
                value: object = row[ci] if ci < len(row) else ""

                # Chart ref cell → skip text rendering; charts are rendered separately
                if isinstance(value, ChartRef):
                    continue

                # Icon cell → render colored oval instead of text box
                icon_name = icon_cell_value(value)
                if icon_name is not None and icons is not None:
                    color = icons.values.get(icon_name)
                    if color:
                        grid_col = ci + col_offset
                        box = layout.cells[ri + row_offset][grid_col]
                        # Center icon within header-row height from top of cell,
                        # so it aligns with the first line of text in adjacent columns.
                        ref_h = layout.row_heights[0] if layout.row_heights else box[3]
                        rpad = (
                            self._col_right_pads[grid_col]
                            if grid_col < len(self._col_right_pads)
                            else 0
                        )
                        self._add_icon(box, icons.size_emu, color, ref_h, rpad)
                    continue

                # Skip truly empty cells (no text box = no placeholder text)
                if value == "" or value is None:
                    continue

                paragraphs = normalize_cell(value, default, parse_bullets=spec.parse_bullets)
                grid_col = ci + col_offset

                # Merge row / column overrides (row wins over column).
                row_ov = spec.row_overrides.get(ri)
                col_ov = spec.col_overrides.get(ci)
                align = "l"
                anchor = "t"
                if col_ov is not None:
                    paragraphs = self._apply_override(paragraphs, col_ov)
                    if col_ov.align:
                        align = col_ov.align
                    if col_ov.anchor:
                        anchor = col_ov.anchor
                if row_ov is not None:
                    paragraphs = self._apply_override(paragraphs, row_ov)
                    if row_ov.align:
                        align = row_ov.align
                    if row_ov.anchor:
                        anchor = row_ov.anchor

                # Use line breaks (no inter-paragraph spacing) when all
                # paragraphs are at lvl 0 — these are header-like sub-labels,
                # not bulleted content.
                use_lb = should_use_line_breaks(paragraphs)
                self._add_cell(
                    layout.cells[ri + row_offset][grid_col],
                    paragraphs,
                    align=align,
                    anchor=anchor,
                    pad_top=layout.pad_top,
                    pad_bottom=layout.pad_bottom,
                    right_pad=self._col_right_pads[grid_col],
                    use_line_breaks=use_lb,
                )

    # -----------------------------------------------------------------------
    # Dividers — flat (all full-width between body rows)
    # -----------------------------------------------------------------------

    def _render_flat_dividers(
        self,
        spec: TableSpec,
        layout: TableLayout,
        area: ContentArea,
        row_offset: int,
    ) -> None:
        start_y = area.y + sum(layout.row_heights[:row_offset])
        for i in range(spec.num_rows - 1):
            y = start_y + sum(layout.row_heights[row_offset : row_offset + i + 1])
            self._add_line(area.x, y, sum(layout.col_widths), int(Dividers.ROW), "dk2")

    # -----------------------------------------------------------------------
    # Dividers — grouped (group = full-width, sub-row = partial)
    # -----------------------------------------------------------------------

    def _render_grouped_dividers(
        self,
        spec: TableSpec,
        layout: TableLayout,
        area: ContentArea,
        row_offset: int,
        _col_offset: int,
    ) -> None:
        groups = spec.groups
        if groups is None:
            return

        superheader_w = layout.col_widths[0]
        body_x = area.x + superheader_w
        body_w = sum(layout.col_widths[1:])

        start_y = area.y + sum(layout.row_heights[:row_offset])
        sub_row = 0

        for gi, group in enumerate(groups):
            # sub-row dividers (partial width, after superheader column)
            for si in range(group.num_rows - 1):
                y = start_y + sum(layout.row_heights[row_offset : row_offset + sub_row + si + 1])
                self._add_line(body_x, y, body_w, int(Dividers.ROW), "dk2")

            sub_row += group.num_rows

            # group divider (full width) — between groups, not after last
            if gi < len(groups) - 1:
                y = start_y + sum(layout.row_heights[row_offset : row_offset + sub_row])
                self._add_line(area.x, y, sum(layout.col_widths), int(Dividers.ROW), "dk2")

    # -----------------------------------------------------------------------
    # Shape creation
    # -----------------------------------------------------------------------

    def _add_cell(
        self,
        box: Box,
        paragraphs: list[Paragraph],
        align: str = "l",
        anchor: str = "t",
        pad_top: int = 0,
        pad_bottom: int = 0,
        right_pad: int = 0,
        use_line_breaks: bool = False,
    ) -> None:
        x, y, width, height = box
        y += pad_top
        height = max(height - pad_top - pad_bottom, 0)
        width = textbox_width(width, right_pad)

        sp = self._make_textbox_sp(x, y, width, height, anchor)
        txBody = sp.find(f"{{{NS_P}}}txBody")
        if txBody is None:
            raise ValueError("Internal error: missing p:txBody on generated textbox")

        paras = paragraphs or [Paragraph(text="", lvl=0)]

        if use_line_breaks and len(paras) > 1:
            # Render all paragraphs as a single <a:p> with <a:br> between them.
            # Avoids inter-paragraph spacing; each segment keeps its own formatting.
            self._append_paragraphs_as_br(txBody, paras, align)
        else:
            for p in paras:
                self._append_paragraph(txBody, p, align)

        self.spTree.append(sp)

    def _make_textbox_sp(self, x: int, y: int, w: int, h: int, anchor: str = "t") -> XmlElement:
        sp = etree.Element(f"{{{NS_P}}}sp")

        nv = etree.SubElement(sp, f"{{{NS_P}}}nvSpPr")
        etree.SubElement(nv, f"{{{NS_P}}}cNvPr", id=str(self.next_id()), name="TextBox")
        cnv = etree.SubElement(nv, f"{{{NS_P}}}cNvSpPr", txBox="1")
        etree.SubElement(cnv, f"{{{NS_A}}}spLocks")
        etree.SubElement(nv, f"{{{NS_P}}}nvPr")

        spPr = etree.SubElement(sp, f"{{{NS_P}}}spPr")
        xfrm = etree.SubElement(spPr, f"{{{NS_A}}}xfrm")
        etree.SubElement(xfrm, f"{{{NS_A}}}off", x=str(x), y=str(y))
        etree.SubElement(xfrm, f"{{{NS_A}}}ext", cx=str(w), cy=str(h))
        geom = etree.SubElement(spPr, f"{{{NS_A}}}prstGeom", prst="rect")
        etree.SubElement(geom, f"{{{NS_A}}}avLst")
        etree.SubElement(spPr, f"{{{NS_A}}}noFill")

        txBody = etree.SubElement(sp, f"{{{NS_P}}}txBody")
        txBody.append(create_text_placeholder_bodyPr(anchor=anchor))
        txBody.append(create_placeholder_lstStyle_xml())

        return sp

    def _append_paragraphs_as_br(
        self, txBody: XmlElement, paragraphs: list[Paragraph], align: str = "l"
    ) -> None:
        """Render multiple Paragraph objects as a single <a:p> with <a:br> between them.

        Each segment keeps its own font/size/bold/color. This produces the same
        visual as separate paragraphs but without inter-paragraph spacing (spcBef).
        """
        first = paragraphs[0]
        p = etree.SubElement(txBody, f"{{{NS_A}}}p")
        pPr: dict[str, str] = {"lvl": str(first.lvl or 0)}
        if align:
            pPr["algn"] = align
        etree.SubElement(p, f"{{{NS_A}}}pPr", attrib=pPr)

        for i, para in enumerate(paragraphs):
            if i > 0:
                # Insert line break with the NEW segment's formatting
                br = etree.SubElement(p, f"{{{NS_A}}}br")
                br_font = para.font or Fonts.BODY
                br_size = int((para.size_pt or FontSizes.DEFAULT) * OOXML_FONT_SCALE)
                br_attrs: dict[str, str] = {"lang": "en-US", "sz": str(br_size)}
                if para.bold:
                    br_attrs["b"] = "1"
                brPr = etree.SubElement(br, f"{{{NS_A}}}rPr", attrib=br_attrs)
                _apply_color(brPr, para.color or DefaultColors.BODY_TEXT)
                etree.SubElement(brPr, f"{{{NS_A}}}latin", typeface=br_font)

            font = para.font or Fonts.BODY
            size = int((para.size_pt or FontSizes.DEFAULT) * OOXML_FONT_SCALE)
            bold = bool(para.bold)
            italic = bool(para.italic)
            underline = bool(para.underline)
            color = para.color or DefaultColors.BODY_TEXT
            self._add_runs(p, para.text, font, size, bold, italic, color, underline)

    def _append_paragraph(self, txBody: XmlElement, paragraph: Paragraph, align: str = "l") -> None:
        p = etree.SubElement(txBody, f"{{{NS_A}}}p")
        pPr: dict[str, str] = {"lvl": str(paragraph.lvl or 0)}
        if align:
            pPr["algn"] = align
        etree.SubElement(p, f"{{{NS_A}}}pPr", attrib=pPr)

        font = paragraph.font or Fonts.BODY
        size = int((paragraph.size_pt or FontSizes.DEFAULT) * OOXML_FONT_SCALE)
        bold = bool(paragraph.bold)
        italic = bool(paragraph.italic)
        underline = bool(paragraph.underline)
        color = paragraph.color or DefaultColors.BODY_TEXT

        self._add_runs(p, paragraph.text, font, size, bold, italic, color, underline)

    def _add_runs(
        self,
        p: XmlElement,
        text: str,
        font: str,
        size: int,
        bold: bool,
        italic: bool,
        color: str,
        underline: bool,
    ) -> None:
        last = 0
        for m in _INLINE_RE.finditer(text):
            if m.start() > last:
                self._run(p, text[last : m.start()], font, size, bold, italic, color, underline)

            if m.group(1):
                bold_text = m.group(2)
                if bold_text is not None:
                    self._run(p, bold_text, font, size, True, italic, color, underline)
            elif m.group(3):
                italic_text = m.group(4)
                if italic_text is not None:
                    self._run(p, italic_text, font, size, bold, True, color, underline)
            elif m.group(5):
                link_text = m.group(6)
                link_url = m.group(7)
                if link_text is not None:
                    self._run(
                        p,
                        link_text,
                        font,
                        size,
                        bold,
                        italic,
                        DefaultColors.LINK,
                        True,
                        hyperlink=link_url or None,
                    )
            else:
                appendix_text = m.group(8)
                if appendix_text is not None:
                    self._run(p, appendix_text, font, size, bold, italic, DefaultColors.LINK, True)

            last = m.end()

        if last < len(text):
            self._run(p, text[last:], font, size, bold, italic, color, underline)
        elif not text:
            etree.SubElement(p, f"{{{NS_A}}}endParaRPr", lang="en-US")

    def _run(
        self,
        p: XmlElement,
        text: str,
        font: str,
        size: int,
        bold: bool,
        italic: bool,
        color: str,
        underline: bool,
        hyperlink: str | None = None,
    ) -> None:
        r = etree.SubElement(p, f"{{{NS_A}}}r")
        attrs: dict[str, str] = {"lang": "en-US", "sz": str(size)}
        if bold:
            attrs["b"] = "1"
        if italic:
            attrs["i"] = "1"
        if underline:
            attrs["u"] = "sng"
        rPr = etree.SubElement(r, f"{{{NS_A}}}rPr", attrib=attrs)
        _apply_color(rPr, color)
        etree.SubElement(rPr, f"{{{NS_A}}}latin", typeface=font)
        # hlinkClick must follow latin/ea/cs/sym per OOXML sequence
        if hyperlink and self._slide_part is not None:
            rId = self._slide_part.relate_to(hyperlink, _HYPERLINK_REL, is_external=True)
            etree.SubElement(
                rPr,
                f"{{{NS_A}}}hlinkClick",
                {f"{{{NS_R}}}id": rId},
            )
        t = etree.SubElement(r, f"{{{NS_A}}}t")
        if text.startswith(" ") or text.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text

    # -----------------------------------------------------------------------
    # Icon indicators (traffic lights)
    # -----------------------------------------------------------------------

    def _add_icon(
        self, box: Box, size_emu: int, color: str, ref_h: int = 0, right_pad: int = 0
    ) -> None:
        """Add a filled circle centered in *box*.

        *ref_h* — reference height for vertical centering.  When provided the
        icon is centered within *ref_h* from the top of the cell instead of
        the full cell height.
        *right_pad* — column gutter to exclude from horizontal centering so
        the icon aligns with the column header text.
        """
        x, y, w, h = box
        content_w = w - right_pad
        cx = x + (content_w - size_emu) // 2
        eff_h = ref_h if ref_h > 0 else h
        cy = y + (eff_h - size_emu) // 2

        sp = etree.Element(f"{{{NS_P}}}sp")
        nv = etree.SubElement(sp, f"{{{NS_P}}}nvSpPr")
        etree.SubElement(nv, f"{{{NS_P}}}cNvPr", id=str(self.next_id()), name="Icon")
        etree.SubElement(nv, f"{{{NS_P}}}cNvSpPr")
        etree.SubElement(nv, f"{{{NS_P}}}nvPr")

        spPr = etree.SubElement(sp, f"{{{NS_P}}}spPr")
        xfrm = etree.SubElement(spPr, f"{{{NS_A}}}xfrm")
        etree.SubElement(xfrm, f"{{{NS_A}}}off", x=str(cx), y=str(cy))
        etree.SubElement(xfrm, f"{{{NS_A}}}ext", cx=str(size_emu), cy=str(size_emu))
        geom = etree.SubElement(spPr, f"{{{NS_A}}}prstGeom", prst="ellipse")
        etree.SubElement(geom, f"{{{NS_A}}}avLst")

        # Fill — hex or scheme color
        _apply_color(spPr, color)

        # No outline
        ln = etree.SubElement(spPr, f"{{{NS_A}}}ln")
        etree.SubElement(ln, f"{{{NS_A}}}noFill")

        # Empty text body (required by OOXML spec)
        txBody = etree.SubElement(sp, f"{{{NS_P}}}txBody")
        etree.SubElement(txBody, f"{{{NS_A}}}bodyPr", rtlCol="0", anchor="ctr")
        etree.SubElement(txBody, f"{{{NS_A}}}lstStyle")
        p = etree.SubElement(txBody, f"{{{NS_A}}}p")
        etree.SubElement(p, f"{{{NS_A}}}endParaRPr", lang="en-US")

        self.spTree.append(sp)

    def _render_legend(self, icons: IconSet, area: ContentArea) -> None:
        """Render icon legend in the top-right of the slide.

        The legend is positioned outside the table — in the top-right
        corner of the slide between the tracker and the content area,
        right-aligned with the content area's right edge.

        Circle diameter matches font height (8pt) so text and icon
        are visually balanced.
        """
        label_font = Fonts.BODY
        label_size_pt = TableDefaults.LEGEND_LABEL_PT

        # Circle diameter = font cap height ≈ font size in EMU
        dot_size = int(label_size_pt * EMU_PER_PT)
        row_height = int(dot_size * TableDefaults.LEGEND_ROW_HEIGHT_RATIO)
        label_gap = int(dot_size * TableDefaults.LEGEND_GAP_RATIO)

        legend_entries = list(icons.values.items())  # ordered as defined

        # Estimate label width from character count
        max_label_len = max(len(name) for name, _ in legend_entries) if legend_entries else 5
        label_w = int(
            max_label_len * label_size_pt * TableDefaults.LEGEND_CHAR_WIDTH_RATIO * EMU_PER_PT
        )
        legend_w = dot_size + label_gap + label_w

        if icons.legend_x is not None:
            lx = icons.legend_x
        else:
            slide_right = area.x + area.width
            lx = slide_right - legend_w

        if icons.legend_y is not None:
            ly = icons.legend_y
        else:
            ly = int(Layout.TRACKER_Y) + int(TableDefaults.LEGEND_Y_OFFSET_INCHES * EMU_PER_INCH)

        for i, (name, color) in enumerate(legend_entries):
            entry_y = ly + i * row_height

            # Circle — sized to match font height
            self._add_icon((lx, entry_y, dot_size, row_height), dot_size, color)

            # Label text box (to the right of the circle)
            text_x = lx + dot_size + label_gap
            text_y = entry_y
            text_h = row_height

            sp = self._make_textbox_sp(text_x, text_y, label_w, text_h, anchor="ctr")
            txBody = sp.find(f"{{{NS_P}}}txBody")
            if txBody is None:
                raise ValueError("Internal error: missing p:txBody on generated legend textbox")

            para = Paragraph(
                text=name,
                lvl=0,
                font=label_font,
                size_pt=label_size_pt,
                color=DefaultColors.BODY_TEXT,
                bold=False,
            )
            self._append_paragraph(txBody, para, align="l")
            self.spTree.append(sp)

    def render_sidebar(self, paragraphs: list[Paragraph], area: ContentArea) -> None:
        """Render a list of paragraphs into a sidebar content area."""
        sp = self._make_textbox_sp(area.x, area.y, area.width, area.height)
        txBody = sp.find(f".//{{{NS_P}}}txBody")
        assert txBody is not None

        for para in paragraphs:
            self._append_paragraph(txBody, para, align="l")

        self.spTree.append(sp)

    def _add_line(self, x: int, y: int, width: int, weight: int, color: str) -> None:
        line = create_line_xml(x, y, width, weight, color)
        for el in line.iter():
            if el.get("id"):
                el.set("id", str(self.next_id()))
        self.spTree.append(line)
