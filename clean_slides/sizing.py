"""
Column and row sizing algorithms.

Determines column widths (by longest word + bullet margins) and row
heights (equal body rows, header sized by content).  Warns on overflow.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from typing_extensions import TypeGuard

from .constants import BULLET_MARGINS, TableDefaults
from .content import Paragraph, header_text_and_sub, normalize_cell
from .measure import (
    cell_content_height,
    column_right_pads,
    should_use_line_breaks,
    text_width_for_level,
    textbox_width,
)
from .spec import ChartRef, TableSpec, is_icon_cell
from .text_metrics import EMU_PER_PT, TextMetrics


@dataclass
class SizingWarning:
    message: str
    details: dict[str, object]


# Safety factor on min-widths — PowerPoint text shaping wraps slightly
# earlier than our character-width estimates.
_MIN_WIDTH_SAFETY = 1.10

# ---------------------------------------------------------------------------
# Font config bundle — avoids passing 4+ font args through every method
# ---------------------------------------------------------------------------


@dataclass
class FontConfig:
    body_font: str
    body_size_pt: int
    header_font: str
    header_size_pt: int
    row_superheader_size_pt: int | None = None  # defaults to header_size_pt

    @property
    def effective_row_superheader_size_pt(self) -> int:
        return (
            self.row_superheader_size_pt
            if self.row_superheader_size_pt is not None
            else self.header_size_pt
        )


# ---------------------------------------------------------------------------
# Column sizer
# ---------------------------------------------------------------------------


class ColumnSizer:
    """Determine column widths: min-width by longest word, then distribute slack."""

    def size(
        self,
        spec: TableSpec,
        area_width: int,
        metrics: TextMetrics,
        fonts: FontConfig,
        pad_top: int = 0,
    ) -> tuple[list[int], list[SizingWarning]]:
        warnings: list[SizingWarning] = []
        col_count = spec.num_cols + (1 if spec.has_row_header else 0)

        # Right padding per column (inter-column gap). Renderer subtracts this
        # from the text-box width; we add it to min/max widths so sizing stays
        # consistent.
        col_right_pads = column_right_pads(col_count, pad_top, spec.has_row_header)

        min_widths = self._min_widths(spec, area_width, metrics, fonts, warnings)
        # Add right_pad to min widths so text fits after pad is subtracted
        for i in range(len(min_widths)):
            min_widths[i] += col_right_pads[i] if i < len(col_right_pads) else 0
        total_min = sum(min_widths)

        if total_min > area_width:
            warnings.append(
                SizingWarning(
                    "Minimum column widths exceed available width",
                    {"available": area_width, "minimum": total_min},
                )
            )
            return min_widths, warnings

        if total_min == 0:
            w = area_width // max(col_count, 1)
            widths = [w] * col_count
            widths[-1] += area_width - sum(widths)
            return widths, warnings

        extra = area_width - total_min

        # Resolve column_widths mode.
        col_widths = spec.col_widths
        explicit_weights = isinstance(col_widths, list)

        # Explicit proportions or "equal": distribute slack by weights.
        if explicit_weights or col_widths == "equal":
            weights = self._column_weights(spec, col_count)
            widths = self._distribute(min_widths, extra, weights)

            # Cap row-header to its preferred single-line width and
            # redistribute overflow to body columns.  Applies to both
            # ``column_widths: equal`` and explicit weights that include
            # a row-header entry.
            should_cap = spec.has_row_header and (
                not explicit_weights or len(col_widths) == col_count
            )
            if should_cap:
                cap = self._row_header_preferred_width(spec, metrics, fonts)
                cap += col_right_pads[0] if col_right_pads else 0
                cap = max(cap, min_widths[0])
                if widths and widths[0] > cap:
                    overflow = widths[0] - cap
                    widths[0] = cap
                    if len(widths) > 1 and overflow > 0:
                        widths[1:] = self._distribute(widths[1:], overflow, weights[1:])

            return widths, warnings

        # Default: HTML4 §B.5.2 auto-layout algorithm.
        #
        # Two widths per column:
        #   min  = longest unbreakable word  (already computed above)
        #   max  = widest cell with zero wrapping
        #
        # Three cases:
        #   sum(max) ≤ area  → use max widths (everything fits, no stretch)
        #   sum(min) ≥ area  → use min widths (already handled above)
        #   otherwise        → interpolate:
        #       width[i] = min[i] + (max[i] - min[i]) × W / D
        #       W = area - sum(min),  D = sum(max) - sum(min)

        max_widths = self._max_widths(spec, area_width, metrics, fonts)
        for i in range(len(max_widths)):
            max_widths[i] += col_right_pads[i] if i < len(col_right_pads) else 0
            max_widths[i] = max(max_widths[i], min_widths[i])

        total_max = sum(max_widths)

        if total_max <= area_width:
            # Everything fits without wrapping.  When chart columns exist
            # distribute the remaining slack so the table fills the area
            # (chart cells expand to use whatever space is available).
            has_charts = bool(spec.chart_defs)
            slack = area_width - total_max
            if has_charts and slack > 0:
                weights = self._column_weights(spec, col_count)
                result = self._distribute(max_widths, slack, weights)
                self._equalize_chart_cols(spec, result)
                return result, warnings
            return max_widths, warnings

        # Interpolate between min and max.
        D = total_max - total_min
        if D <= 0:
            weights = self._column_weights(spec, col_count)
            return self._distribute(min_widths, extra, weights), warnings

        widths = min_widths[:]
        allocated = 0
        for i in range(len(widths) - 1):
            di = max_widths[i] - min_widths[i]
            add = int(extra * (di / D)) if di > 0 else 0
            widths[i] += add
            allocated += add
        widths[-1] += extra - allocated

        # Equalize columns spanned by vertical charts (bars are equal width
        # regardless of column sizing, so spanned columns must match).
        self._equalize_chart_cols(spec, widths)

        return widths, warnings

    # -- helpers --

    @staticmethod
    def _equalize_chart_cols(spec: TableSpec, widths: list[int]) -> None:
        """Make columns spanned by the same vertical chart equal width.

        A vertical chart distributes bars equally across its span.
        If spanned columns differ in width the bars won't align with
        column boundaries.
        """
        if not spec.chart_defs or not spec.cells:
            return

        col_offset = spec.col_offset
        for chart_def in spec.chart_defs.values():
            if chart_def.dir != "vertical":
                continue
            cols_in_chart: set[int] = set()
            for row in spec.cells:
                for ci, cell in enumerate(row):
                    if isinstance(cell, ChartRef) and cell.name == chart_def.name:
                        cols_in_chart.add(ci + col_offset)
            if len(cols_in_chart) < 2:
                continue
            avg_w = sum(widths[c] for c in cols_in_chart if c < len(widths)) // len(cols_in_chart)
            for c in cols_in_chart:
                if c < len(widths):
                    widths[c] = avg_w

    def _min_widths(
        self,
        spec: TableSpec,
        area_width: int,
        metrics: TextMetrics,
        fonts: FontConfig,
        warnings: list[SizingWarning],
    ) -> list[int]:
        """Minimum width per column so the longest word fits."""
        mins: list[int] = []

        if spec.has_row_header:
            mins.append(
                self._row_header_min_width(
                    spec,
                    area_width,
                    metrics,
                    fonts,
                    warnings,
                )
            )

        body_default = _body_default(spec, fonts)
        icon_min = spec.icons.size_emu if spec.icons else 0
        for col_idx in range(spec.num_cols):
            header_w = self._header_word_width(spec, col_idx, metrics, fonts)
            body_w = self._body_word_width(spec, col_idx, metrics, fonts, body_default)
            # Icon columns need at least icon_size as minimum
            floor = icon_min if self._is_icon_column(spec, col_idx) else 1
            mins.append(int(max(header_w, body_w, floor) * _MIN_WIDTH_SAFETY))

        return mins

    def _row_header_min_width(
        self,
        spec: TableSpec,
        area_width: int,
        metrics: TextMetrics,
        fonts: FontConfig,
        warnings: list[SizingWarning],
    ) -> int:
        """Min width for the row-header / superheader column.

        Must fit every text that renders in column 0:
          - row superheaders (grouped) or row headers (flat)
          - row_header_col_header (the col-header text for this column)
          - first col_superheader label (spans this column)
        Each at its actual render font size.  A small safety margin is
        added because PowerPoint's text shaping may wrap slightly earlier
        than our metrics predict.
        """
        max_w = 0

        # 1. Row superheaders / row headers — render at superheader size (grouped) or header size
        size_pt = (
            fonts.effective_row_superheader_size_pt if spec.is_grouped else fonts.header_size_pt
        )
        for header in spec.row_headers or []:
            text, sub_text = header_text_and_sub(header)

            # Default min width is driven by the longest unbreakable token.
            w = max(_longest_word_width(text, fonts.header_font, size_pt, metrics), 1)
            if sub_text:
                w = max(w, _longest_word_width(sub_text, fonts.header_font, size_pt, metrics))

            # UX tweak: avoid clunky wraps like
            #
            #   "Matriarchal"
            #   "herds"
            #
            # For labels with 2+ words, ensure *any adjacent pair* of words fits on
            # one line. This makes the line-break algorithm much less likely to
            # produce single-word lines.
            words = [t for t in str(text).split() if t]
            target_lines = TableDefaults.ROW_HEADER_TARGET_LINES

            if len(words) >= 2:
                # Adjacent 2-word chunks treated as effectively unbreakable.
                for i in range(len(words) - 1):
                    pair = f"{words[i]} {words[i + 1]}"
                    w = max(w, metrics.text_width_no_wrap(pair, fonts.header_font, size_pt))

            # Special-case: exactly 2 words → aim for a single line when possible.
            if len(words) == 2:
                target_lines = 1

            step = int(TableDefaults.WIDTH_STEP)
            while (
                metrics.lines_needed(text, w, fonts.header_font, size_pt) > target_lines
                and w < area_width
            ):
                w += step
            if metrics.lines_needed(text, w, fonts.header_font, size_pt) > target_lines:
                warnings.append(
                    SizingWarning(
                        f"Row header exceeds {target_lines}-line target",
                        {"header": text},
                    )
                )
            max_w = max(max_w, w)

        # 2. row_header_col_header — renders at header_size_pt
        if spec.row_header_col_header:
            rhch_text, rhch_sub = header_text_and_sub(spec.row_header_col_header)
            max_w = max(
                max_w,
                _longest_word_width(rhch_text, fonts.header_font, fonts.header_size_pt, metrics),
            )
            if rhch_sub:
                max_w = max(
                    max_w,
                    _longest_word_width(
                        rhch_sub, fonts.header_font, fonts.header_size_pt, metrics
                    ),
                )

        # 3. First col superheader label (if it covers this column) — renders at header_size_pt
        if spec.col_superheaders and spec.col_superheaders[0].label:
            text = str(spec.col_superheaders[0].label)
            max_w = max(
                max_w, _longest_word_width(text, fonts.header_font, fonts.header_size_pt, metrics)
            )

        # Safety margin: PowerPoint/LibreOffice text shaping is slightly wider
        max_w = int(max_w * _MIN_WIDTH_SAFETY)
        return max(max_w, 1)

    # Bold text is ~10% wider than regular. The text metric doesn't
    # model weight, so we apply a fudge factor for bold columns
    # (row headers, superheaders).
    _BOLD_FACTOR: float = 1.10

    def _row_header_preferred_width(
        self,
        spec: TableSpec,
        metrics: TextMetrics,
        fonts: FontConfig,
    ) -> int:
        """Preferred (no-wrap) width for the row-header / superheader column.

        Returns the width of the widest single-line label that can appear in
        the row-header column (excluding the inter-column gap/right-pad).

        Used to avoid wasting slack on the row-header column when explicit
        full `column_widths` proportions are provided.
        """
        max_w = 0

        size_pt = (
            fonts.effective_row_superheader_size_pt if spec.is_grouped else fonts.header_size_pt
        )
        for header in spec.row_headers or []:
            text, sub_text = header_text_and_sub(header)
            max_w = max(
                max_w,
                int(metrics.text_width_no_wrap(text, fonts.header_font, size_pt) * self._BOLD_FACTOR),
            )
            if sub_text:
                max_w = max(
                    max_w,
                    int(metrics.text_width_no_wrap(sub_text, fonts.header_font, size_pt)),
                )

        if spec.row_header_col_header:
            rhch_text, rhch_sub = header_text_and_sub(spec.row_header_col_header)
            max_w = max(
                max_w,
                metrics.text_width_no_wrap(rhch_text, fonts.header_font, fonts.header_size_pt),
            )
            if rhch_sub:
                max_w = max(
                    max_w,
                    metrics.text_width_no_wrap(
                        rhch_sub, fonts.header_font, fonts.header_size_pt
                    ),
                )

        if spec.col_superheaders and spec.col_superheaders[0].label:
            max_w = max(
                max_w,
                metrics.text_width_no_wrap(
                    str(spec.col_superheaders[0].label),
                    fonts.header_font,
                    fonts.header_size_pt,
                ),
            )

        max_w = int(max_w * _MIN_WIDTH_SAFETY)
        return max(max_w, 1)

    def _header_word_width(
        self, spec: TableSpec, col_idx: int, metrics: TextMetrics, fonts: FontConfig
    ) -> int:
        if not spec.has_col_header or not spec.col_headers:
            return 0

        text, sub_text = header_text_and_sub(spec.col_headers[col_idx])
        w = _longest_word_width(text, fonts.header_font, fonts.header_size_pt, metrics)
        if sub_text:
            w = max(w, _longest_word_width(sub_text, fonts.header_font, fonts.header_size_pt, metrics))

        # Similar to row-header logic: avoid clunky single-word lines for
        # multi-word headers by ensuring adjacent word pairs fit.
        words = [t for t in text.split() if t]
        if len(words) >= 2:
            for i in range(len(words) - 1):
                pair = f"{words[i]} {words[i + 1]}"
                w = max(
                    w, metrics.text_width_no_wrap(pair, fonts.header_font, fonts.header_size_pt)
                )

        return w

    def _body_word_width(
        self,
        spec: TableSpec,
        col_idx: int,
        metrics: TextMetrics,
        fonts: FontConfig,
        body_default: Paragraph,
    ) -> int:
        max_w = 0
        for row in spec.cells or []:
            if col_idx >= len(row):
                continue
            cell_val = row[col_idx]
            if isinstance(cell_val, ChartRef):
                continue
            for p in normalize_cell(cell_val, body_default, parse_bullets=spec.parse_bullets):
                lvl = p.lvl or 0
                margin = BULLET_MARGINS.get(lvl + 1, (0, 0, 0))[0]
                w = _longest_word_width(
                    p.text, p.font or fonts.body_font, p.size_pt or fonts.body_size_pt, metrics
                )
                max_w = max(max_w, w + margin)
        return max_w

    @staticmethod
    def _is_icon_column(spec: TableSpec, col_idx: int) -> bool:
        """Return True if *col_idx* contains at least one icon cell."""
        if not spec.icons:
            return False
        return any(col_idx < len(row) and is_icon_cell(row[col_idx]) for row in spec.cells or [])

    @staticmethod
    def _is_chart_column(spec: TableSpec, col_idx: int) -> bool:
        """Return True if *col_idx* contains at least one chart ref cell."""
        if not spec.chart_defs:
            return False
        return any(
            col_idx < len(row) and isinstance(row[col_idx], ChartRef)
            for row in spec.cells or []
        )

    def _max_widths(
        self,
        spec: TableSpec,
        area_width: int,
        metrics: TextMetrics,
        fonts: FontConfig,
    ) -> list[int]:
        """Capped no-wrap width per column (used by the 'default' algorithm).

        The width each column would need if no text wrapping occurred at all,
        capped so that no single body column exceeds 1.5× its equal share of
        the area.  This prevents one long cell from dominating the table and
        makes it more likely that ``sum(max) ≤ area`` (triggering the
        "don't fill" path that leaves whitespace on the right).

        Row-header and icon columns are not capped (they use their natural
        no-wrap / icon width).
        """
        col_count = spec.num_cols + (1 if spec.has_row_header else 0)
        body_cap = int(area_width / max(col_count, 1) * 1.5)

        maxes: list[int] = []

        if spec.has_row_header:
            # Row header: use natural no-wrap width (usually short labels).
            maxes.append(self._row_header_preferred_width(spec, metrics, fonts))

        body_default = _body_default(spec, fonts)
        icon_min = spec.icons.size_emu if spec.icons else 0

        for col_idx in range(spec.num_cols):
            is_icon = self._is_icon_column(spec, col_idx)

            header_w = 0
            if spec.has_col_header and spec.col_headers and col_idx < len(spec.col_headers):
                hdr_text, hdr_sub = header_text_and_sub(spec.col_headers[col_idx])
                header_w = metrics.text_width_no_wrap(
                    hdr_text,
                    fonts.header_font,
                    fonts.header_size_pt,
                )
                if hdr_sub:
                    header_w = max(
                        header_w,
                        metrics.text_width_no_wrap(
                            hdr_sub, fonts.header_font, fonts.header_size_pt
                        ),
                    )

            body_w = 0
            for row in spec.cells or []:
                if col_idx >= len(row):
                    continue
                cell_val = row[col_idx]
                if isinstance(cell_val, ChartRef):
                    continue
                for p in normalize_cell(
                    cell_val,
                    body_default,
                    parse_bullets=spec.parse_bullets,
                ):
                    lvl = p.lvl or 0
                    margin = BULLET_MARGINS.get(lvl + 1, (0, 0, 0))[0]
                    w = metrics.text_width_no_wrap(
                        p.text,
                        p.font or fonts.body_font,
                        p.size_pt or fonts.body_size_pt,
                    )
                    body_w = max(body_w, w + margin)

            # Chart-only columns: the chart fills whatever space is given.
            # Use header width as the preferred size (not an equal share) so
            # that text-only columns keep enough room.  The distribution step
            # will expand chart columns with remaining slack.
            is_chart = self._is_chart_column(spec, col_idx)
            if is_chart and body_w == 0:
                body_w = header_w  # 0 is fine — min_width guarantees a floor

            floor = icon_min if is_icon else 1
            raw = int(max(header_w, body_w, floor) * _MIN_WIDTH_SAFETY)

            # Cap body columns; leave icon columns uncapped (they're tiny).
            if not is_icon:
                raw = min(raw, body_cap)

            maxes.append(raw)

        return maxes

    def _column_weights(self, spec: TableSpec, col_count: int) -> list[float]:
        """Return weights for distributing slack across columns.

        Used by ``column_widths: equal`` and ``column_widths: [...]`` modes.
        """
        cw: object = spec.col_widths
        if isinstance(cw, list):
            weights: list[float] = [float(v) for v in cw]
            if spec.has_row_header:
                if len(weights) == col_count:
                    return weights
                if len(weights) == spec.num_cols:
                    return [0.0, *weights]
            else:
                if len(weights) == col_count:
                    return weights

        # Equal: all body columns get the same weight.
        # Row header gets a small weight for breathing room.
        weights: list[float] = []
        if spec.has_row_header:
            weights.append(0.5)
        for _ in range(spec.num_cols):
            weights.append(1.0)
        return weights

    @staticmethod
    def _distribute(min_widths: list[int], extra: int, weights: list[float]) -> list[int]:
        total_weight = sum(weights) or 1.0
        widths = min_widths[:]
        allocated = 0
        for i, w in enumerate(weights[:-1]):
            add = int(extra * (w / total_weight))
            widths[i] += add
            allocated += add
        widths[-1] += extra - allocated
        return widths


# ---------------------------------------------------------------------------
# Row sizer
# ---------------------------------------------------------------------------


class RowSizer:
    """Equal body rows, header row sized by content (2–4 lines).

    After initial equal sizing, transfers height from slack rows to
    overflow rows where possible.
    """

    def size(
        self,
        spec: TableSpec,
        col_widths: list[int],
        area_height: int,
        metrics: TextMetrics,
        fonts: FontConfig,
        pad_top: int,
        pad_bottom: int,
    ) -> tuple[list[int], list[SizingWarning]]:
        warnings: list[SizingWarning] = []
        col_offset = spec.col_offset

        # Text box widths (renderer subtracts per-column right padding from the
        # cell box width to create a visible inter-column gap).
        right_pads = column_right_pads(len(col_widths), pad_top, spec.has_row_header)
        text_widths = [textbox_width(w, rp) for w, rp in zip(col_widths, right_pads)]

        # -- column superheader row --
        col_super_h = 0
        if spec.has_col_superheader:
            col_super_h = self._col_superheader_height(spec, metrics, fonts, pad_top, pad_bottom)

        # -- header row --
        header_h = 0
        if spec.has_col_header:
            header_h = self._header_height(
                spec, text_widths, col_offset, metrics, fonts, pad_top, pad_bottom
            )

        # -- body rows (proportional to content) --
        body_area = area_height - col_super_h - header_h
        if body_area < 0:
            warnings.append(
                SizingWarning(
                    "Header rows exceed available height",
                    {"headers": col_super_h + header_h, "available": area_height},
                )
            )
            body_area = 0

        heights: list[int] = []
        if spec.has_col_superheader:
            heights.append(col_super_h)
        if spec.has_col_header:
            heights.append(header_h)

        header_default = _header_default(fonts)
        body_default = _body_default(spec, fonts)

        # Compute required height for each body row
        required_body: list[int] = []
        for ri in range(spec.num_rows):
            req = self._body_row_required(
                spec,
                ri,
                text_widths,
                col_offset,
                metrics,
                header_default,
                body_default,
                pad_top,
                pad_bottom,
            )
            required_body.append(max(req, 1))

        # Distribute body_area proportionally to required heights
        total_req = sum(required_body) or 1
        min_h = int(TableDefaults.MIN_ROW_HEIGHT)
        body_heights: list[int] = []
        for req in required_body:
            h = max(int(body_area * req / total_req), min_h)
            body_heights.append(h)

        # Equalize rows spanned by horizontal charts (bars are equal height
        # regardless of text content, so the rows must match).
        self._equalize_chart_rows(spec, body_heights)

        # Fix rounding: adjust last row to consume exactly body_area
        allocated = sum(body_heights)
        if body_heights:
            body_heights[-1] += body_area - allocated

        heights.extend(body_heights)

        if spec.num_rows and body_area // max(spec.num_rows, 1) < min_h:
            warnings.append(
                SizingWarning(
                    "Average row height below minimum",
                    {"average": body_area // max(spec.num_rows, 1), "minimum": min_h},
                )
            )

        # Warn if total content exceeds available area
        if total_req > body_area:
            # Collect per-cell required heights to identify overflow culprits
            cell_heights = self._body_cell_heights(
                spec,
                text_widths,
                col_offset,
                metrics,
                header_default,
                body_default,
                pad_top,
                pad_bottom,
            )
            # For each row, find allocated vs required
            overflow_cells: list[str] = []
            for ri, (req, alloc) in enumerate(zip(required_body, body_heights)):
                if req <= alloc:
                    continue
                # Which columns drive this row's overflow?
                row_cells = cell_heights[ri] if ri < len(cell_heights) else {}
                tall_cols = [c for c, ch in sorted(row_cells.items()) if ch > alloc]
                if tall_cols:
                    cols_str = ",".join(str(c + 1) for c in tall_cols)
                    overflow_cells.append(f"row {ri + 1} col {cols_str}")

            from pptx.util import Emu

            col_widths_in = [f"{Emu(w).inches:.2f}" for w in col_widths]
            warnings.append(
                SizingWarning(
                    "Vertical overflow: content exceeds available height",
                    {
                        "required": total_req,
                        "available": body_area,
                        "overflow_pct": (
                            round(100 * (total_req - body_area) / body_area) if body_area else 0
                        ),
                        "col_widths_in": col_widths_in,
                        "overflow_cells": overflow_cells,
                    },
                )
            )

        return heights, warnings

    # -- helpers --

    def _col_superheader_height(
        self,
        spec: TableSpec,
        metrics: TextMetrics,
        fonts: FontConfig,
        pad_top: int,
        pad_bottom: int,
    ) -> int:
        """Compact row for column superheaders (1–2 lines depending on sub)."""
        line_h = int(fonts.header_size_pt * EMU_PER_PT * TableDefaults.LINE_SPACING * metrics.fudge)
        max_lines = 1
        if spec.col_superheaders:
            for csh in spec.col_superheaders:
                lines = 1
                if csh.sub:
                    lines += 1
                max_lines = max(max_lines, lines)
        # Minimal vertical padding — just enough to separate from content above
        return line_h * max_lines + pad_top

    def _header_height(
        self,
        spec: TableSpec,
        col_widths: list[int],
        col_offset: int,
        metrics: TextMetrics,
        fonts: FontConfig,
        pad_top: int,
        pad_bottom: int,
    ) -> int:
        max_lines = 1

        # Include row_header_col_header in line count if present
        if spec.row_header_col_header and spec.has_row_header:
            rhch_text, rhch_sub = header_text_and_sub(spec.row_header_col_header)
            w = text_width_for_level(col_widths[0], 0)
            lines = metrics.lines_needed(rhch_text, w, fonts.header_font, fonts.header_size_pt) or 1
            if rhch_sub:
                lines += 1
            max_lines = max(max_lines, lines)

        for col_idx in range(spec.num_cols):
            raw = (
                spec.col_headers[col_idx]
                if spec.col_headers and col_idx < len(spec.col_headers)
                else ""
            )
            text, sub_text = header_text_and_sub(raw)
            w = text_width_for_level(col_widths[col_idx + col_offset], 0)
            lines = metrics.lines_needed(text, w, fonts.header_font, fonts.header_size_pt) or 1
            if sub_text:
                lines += 1
            max_lines = max(max_lines, lines)

        # Fit header row to content: at least 1 line, capped
        lines = min(max(max_lines, 1), TableDefaults.MAX_HEADER_LINES)
        line_h = int(fonts.header_size_pt * EMU_PER_PT * TableDefaults.LINE_SPACING * metrics.fudge)
        spc = int(BULLET_MARGINS.get(1, (0, 0, 0))[2] * EMU_PER_PT)

        if spec.has_col_superheader:
            # Reduce top gap so superheader+divider sit close to header text
            return lines * line_h + pad_top // 2 + pad_bottom
        else:
            return lines * line_h + pad_top + pad_bottom + spc

    @staticmethod
    def _equalize_chart_rows(spec: TableSpec, body_heights: list[int]) -> None:
        """Make rows spanned by horizontal charts equal height.

        A horizontal chart distributes bars equally across its span.
        If the spanned rows have different heights (because of varying
        text content in non-chart columns), the bars won't align with
        the row boundaries.  The spanned rows get equal shares of their
        combined area so total height is preserved.
        """
        if not spec.chart_defs or not spec.cells:
            return

        for chart_def in spec.chart_defs.values():
            if chart_def.dir != "horizontal":
                continue
            rows_in_chart: set[int] = set()
            for ri, row in enumerate(spec.cells):
                for cell in row:
                    if isinstance(cell, ChartRef) and cell.name == chart_def.name:
                        rows_in_chart.add(ri)
                        break
            if len(rows_in_chart) < 2:
                continue
            # Equal share of the combined area (preserves total height)
            valid = [r for r in rows_in_chart if r < len(body_heights)]
            combined = sum(body_heights[r] for r in valid)
            equal_h = combined // len(valid)
            for r in valid:
                body_heights[r] = equal_h
            # Distribute rounding remainder to the last row
            remainder = combined - equal_h * len(valid)
            if remainder and valid:
                body_heights[max(valid)] += remainder

    def _body_row_required(
        self,
        spec: TableSpec,
        body_row: int,
        col_widths: list[int],
        col_offset: int,
        metrics: TextMetrics,
        hdr_def: Paragraph,
        body_def: Paragraph,
        pt: int,
        pb: int,
    ) -> int:
        h = 0
        if spec.has_row_header and not spec.is_grouped:
            # Flat row headers: each row has its own header
            raw_hdr = (
                spec.row_headers[body_row]
                if spec.row_headers and body_row < len(spec.row_headers)
                else ""
            )
            if raw_hdr:
                ps = normalize_cell(raw_hdr, hdr_def, parse_bullets=False)
                h = max(
                    h,
                    cell_content_height(
                        ps,
                        col_widths[0],
                        metrics,
                        pt,
                        pb,
                        hdr_def.font or "Arial",
                        hdr_def.size_pt or 12,
                        use_line_breaks=True,
                    ),
                )
        # Grouped: superheader spans the group, so it doesn't constrain individual sub-row height

        for col_idx in range(spec.num_cols):
            value: object = ""
            if spec.cells and body_row < len(spec.cells):
                row = spec.cells[body_row]
                if col_idx < len(row):
                    value = row[col_idx]
            # Skip empty cells, icon cells, and chart ref cells
            if value == "" or value is None or is_icon_cell(value) or isinstance(value, ChartRef):
                continue
            ps = normalize_cell(value, body_def, parse_bullets=spec.parse_bullets)
            # All-lvl0 cells use line breaks (no spcBef) — matches renderer
            use_lb = should_use_line_breaks(ps)
            h = max(
                h,
                cell_content_height(
                    ps,
                    col_widths[col_idx + col_offset],
                    metrics,
                    pt,
                    pb,
                    body_def.font or "Arial",
                    body_def.size_pt or 12,
                    use_line_breaks=use_lb,
                ),
            )
        return h

    def _body_cell_heights(
        self,
        spec: TableSpec,
        text_widths: list[int],
        col_offset: int,
        metrics: TextMetrics,
        hdr_def: Paragraph,
        body_def: Paragraph,
        pt: int,
        pb: int,
    ) -> list[dict[int, int]]:
        """Return per-cell required heights: list of {col_idx: height} per body row."""
        result: list[dict[int, int]] = []
        for ri in range(spec.num_rows):
            cell_map: dict[int, int] = {}

            # Row header (col 0 when present)
            if spec.has_row_header and not spec.is_grouped:
                raw_hdr = (
                    spec.row_headers[ri]
                    if spec.row_headers and ri < len(spec.row_headers)
                    else ""
                )
                if raw_hdr:
                    ps = normalize_cell(raw_hdr, hdr_def, parse_bullets=False)
                    cell_map[0] = cell_content_height(
                        ps,
                        text_widths[0],
                        metrics,
                        pt,
                        pb,
                        hdr_def.font or "Arial",
                        hdr_def.size_pt or 12,
                        use_line_breaks=True,
                    )

            for ci in range(spec.num_cols):
                value: object = ""
                if spec.cells and ri < len(spec.cells):
                    row = spec.cells[ri]
                    if ci < len(row):
                        value = row[ci]
                if value == "" or value is None or is_icon_cell(value) or isinstance(value, ChartRef):
                    continue
                ps = normalize_cell(value, body_def, parse_bullets=spec.parse_bullets)
                use_lb = should_use_line_breaks(ps)
                cell_map[ci + col_offset] = cell_content_height(
                    ps,
                    text_widths[ci + col_offset],
                    metrics,
                    pt,
                    pb,
                    body_def.font or "Arial",
                    body_def.size_pt or 12,
                    use_line_breaks=use_lb,
                )
            result.append(cell_map)
        return result


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------


def _is_dict(value: object) -> TypeGuard[dict[str, Any]]:
    return isinstance(value, dict)


def _is_list(value: object) -> TypeGuard[list[Any]]:
    return isinstance(value, list)


def _cell_text_length(cell_value: Any) -> int:
    """Extract total text character count from a cell value.

    Handles strings, dicts with 'text' key, and lists of those.
    Icon cells (``{icon: "X"}``) return 0 — they have no text content.

    Note: this is a heuristic used only for sizing weights, so it accepts
    dynamic YAML-derived values.
    """
    if cell_value is None:
        return 0
    if isinstance(cell_value, ChartRef):
        return 0
    if isinstance(cell_value, str):
        return len(cell_value)
    if _is_dict(cell_value):
        if is_icon_cell(cell_value):
            return 0
        return len(str(cell_value.get("text", "")))
    if _is_list(cell_value):
        return sum(_cell_text_length(item) for item in cell_value)
    return len(str(cell_value))


def _longest_word_width(text: str, font: str, size_pt: int, metrics: TextMetrics) -> int:
    word = metrics.longest_word(text or "")
    return metrics.word_width(word, font, size_pt)


def _body_default(spec: TableSpec, fonts: FontConfig) -> Paragraph:
    return Paragraph(
        text="",
        lvl=spec.body_default_lvl,
        font=fonts.body_font,
        size_pt=fonts.body_size_pt,
        color="tx1",
        bold=False,
    )


def _header_default(fonts: FontConfig) -> Paragraph:
    return Paragraph(
        text="",
        lvl=0,
        font=fonts.header_font,
        size_pt=fonts.header_size_pt,
        color="accent2",
        bold=True,
    )
