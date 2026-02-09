"""
Shared measurement utilities for sizing and verification.

Both ColumnSizer and LayoutVerifier need to compute cell content heights
and available text widths. This module centralises that logic.
"""

from __future__ import annotations

from .constants import BULLET_MARGINS, Fonts, FontSizes
from .content import Paragraph
from .text_metrics import EMU_PER_PT, TextMetrics


def column_right_pads(col_count: int, pad_top: int, has_row_header: bool) -> list[int]:
    """Right-side padding per grid column.

    This value is subtracted from the cell width when creating the text box. It
    creates a consistent visual gap between columns while keeping the cell grid
    aligned.

    The values are derived from ``pad_top`` (used as the base "cell padding"
    unit in EMU):
      - normal columns: 2 * pad_top  (≈ 0.10" when pad_top = 0.05")
      - row-header col: 3 * pad_top  (≈ 0.15")
      - last column: 0
    """
    if col_count <= 0:
        return []

    col_gap = pad_top * 2
    row_header_gap = pad_top * 3

    pads: list[int] = []
    for c in range(col_count):
        if c == col_count - 1:
            pads.append(0)
        elif c == 0 and has_row_header:
            pads.append(row_header_gap)
        else:
            pads.append(col_gap)
    return pads


def textbox_width(col_width: int, right_pad: int) -> int:
    """Width available to text after subtracting right-side padding."""
    return max(col_width - right_pad, 0)


def should_use_line_breaks(paragraphs: list[Paragraph]) -> bool:
    """True when paragraphs should render as ``<a:br>`` breaks.

    Rule: when a cell contains multiple level-0 paragraphs (header-like stacked
    labels), render with line breaks so PowerPoint doesn't add inter-paragraph
    spacing (spcBef).
    """
    if len(paragraphs) <= 1:
        return False
    return all((p.lvl or 0) == 0 for p in paragraphs)


def text_width_for_level(col_width: int, lvl: int) -> int:
    """Usable text width after subtracting bullet margin for *lvl*."""
    margin = BULLET_MARGINS.get(lvl + 1, (0, 0, 0))[0]
    return max(col_width - margin, 0)


def paragraph_height(
    paragraph: Paragraph,
    col_width: int,
    metrics: TextMetrics,
    fallback_font: str = "",
    fallback_size_pt: int = 0,
    skip_spc_before: bool = False,
) -> int:
    """Height of a single paragraph (space-before + wrapped text).

    When *skip_spc_before* is True the inter-paragraph spacing is omitted,
    matching the behaviour of ``<a:br>`` line breaks in PowerPoint.
    """
    _font = fallback_font or Fonts.BODY
    _size = fallback_size_pt or FontSizes.DEFAULT

    lvl = paragraph.lvl or 0
    avail = text_width_for_level(col_width, lvl)
    if avail <= 0:
        return 0

    spc_before = (
        0 if skip_spc_before else int(BULLET_MARGINS.get(lvl + 1, (0, 0, 0))[2] * EMU_PER_PT)
    )
    text_h = metrics.text_height(
        paragraph.text,
        avail,
        paragraph.font or _font,
        paragraph.size_pt or _size,
    )
    return spc_before + text_h


def cell_content_height(
    paragraphs: list[Paragraph],
    col_width: int,
    metrics: TextMetrics,
    pad_top: int = 0,
    pad_bottom: int = 0,
    fallback_font: str = "",
    fallback_size_pt: int = 0,
    use_line_breaks: bool = False,
) -> int:
    """Total height needed for a list of paragraphs inside a cell.

    When *use_line_breaks* is True, inter-paragraph spacing (spcBef) is
    omitted for paragraphs after the first — matching ``<a:br>`` rendering.
    """
    body = 0
    for i, p in enumerate(paragraphs):
        body += paragraph_height(
            p,
            col_width,
            metrics,
            fallback_font,
            fallback_size_pt,
            skip_spc_before=(use_line_breaks and i > 0),
        )
    return body + pad_top + pad_bottom
