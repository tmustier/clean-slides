"""
Constraint solver for table layouts.

Orchestrates column sizing → row sizing → verification and returns
a TableLayout + LayoutReport.
"""

from dataclasses import dataclass
from typing import List, Tuple

from .constants import OOXML_FONT_SCALE, DefaultColors, Fonts, TableDefaults
from .content import Paragraph, normalize_cell
from .sizing import ColumnSizer, FontConfig, RowSizer, SizingWarning
from .spec import Box, ContentArea, TableLayout, TableSpec
from .text_metrics import TextMetrics
from .verification import LayoutIssue, LayoutReport, LayoutVerifier


@dataclass
class SolveOptions:
    """Tuneable parameters for the solver."""

    body_font_pt: int = 12
    header_font_pt: int = 12
    min_font_pt: int = 10
    max_font_pt: int = 16
    pad_top: int = int(TableDefaults.CELL_PADDING)
    pad_bottom: int = int(TableDefaults.CELL_PADDING)


class ConstraintSolver:
    """Compute column/row sizes and verify the result."""

    def __init__(self, metrics: TextMetrics):
        self.metrics = metrics
        self.columns = ColumnSizer()
        self.rows = RowSizer()

    _OVERFLOW_MSG = "Vertical overflow: content exceeds available height"

    def solve(
        self,
        spec: TableSpec,
        area: ContentArea,
        options: SolveOptions,
    ) -> Tuple[TableLayout, LayoutReport]:
        # Font reduction loop — shrink body (and header) until content fits
        # or min_font_pt is reached.
        original_body = options.body_font_pt
        original_header = options.header_font_pt

        col_widths: List[int] = []
        row_heights: List[int] = []
        col_warns: List[SizingWarning] = []
        row_warns: List[SizingWarning] = []
        fonts: FontConfig = self._resolve_fonts(spec, options)

        max_delta = max(original_body - options.min_font_pt, 0)

        for delta in range(max_delta + 1):
            trial = SolveOptions(
                body_font_pt=original_body - delta,
                header_font_pt=max(original_header - delta, options.min_font_pt),
                min_font_pt=options.min_font_pt,
                max_font_pt=options.max_font_pt,
                pad_top=options.pad_top,
                pad_bottom=options.pad_bottom,
            )
            fonts = self._resolve_fonts(spec, trial)
            col_widths, col_warns = self.columns.size(
                spec,
                area.width,
                self.metrics,
                fonts,
                trial.pad_top,
            )
            row_heights, row_warns = self.rows.size(
                spec,
                col_widths,
                area.height,
                self.metrics,
                fonts,
                trial.pad_top,
                trial.pad_bottom,
            )

            has_overflow = any(w.message == self._OVERFLOW_MSG for w in row_warns)
            if not has_overflow:
                if delta > 0:
                    row_warns.append(
                        SizingWarning(
                            f"Font reduced to fit: body {original_body}→{trial.body_font_pt}pt, "
                            f"header {original_header}→{fonts.header_size_pt}pt",
                            {
                                "original_body": original_body,
                                "body": trial.body_font_pt,
                                "original_header": original_header,
                                "header": fonts.header_size_pt,
                            },
                        )
                    )
                break

        layout = TableLayout(
            col_widths=col_widths,
            row_heights=row_heights,
            header_font_size=fonts.header_size_pt * OOXML_FONT_SCALE,
            body_font_size=fonts.body_size_pt * OOXML_FONT_SCALE,
            pad_top=options.pad_top,
            pad_bottom=options.pad_bottom,
            cells=self._build_cells(area, col_widths, row_heights),
        )

        report = LayoutReport()
        for w in col_warns + row_warns:
            report.add(LayoutIssue("warning", "sizing", w.message, w.details))

        verifier = LayoutVerifier(layout, area)
        report.issues += verifier.run_all(spec, self.metrics).issues
        return layout, report

    # -- helpers --

    def _resolve_fonts(self, spec: TableSpec, options: SolveOptions) -> FontConfig:
        """Ensure header_font_pt >= max body paragraph size."""
        max_body = options.body_font_pt
        body_default = Paragraph(
            text="",
            lvl=spec.body_default_lvl,
            font=Fonts.BODY,
            size_pt=options.body_font_pt,
            color=DefaultColors.BODY_TEXT,
            bold=False,
        )
        for row in spec.cells or []:
            for cell in row:
                for p in normalize_cell(cell, body_default, parse_bullets=spec.parse_bullets):
                    max_body = max(max_body, p.size_pt or options.body_font_pt)

        header_pt = max(options.header_font_pt, max_body)
        # Row superheaders render larger than headers
        superheader_pt = (
            header_pt + TableDefaults.SUPERHEADER_PT_BOOST if spec.is_grouped else header_pt
        )
        return FontConfig(
            body_font=Fonts.BODY,
            body_size_pt=options.body_font_pt,
            header_font=Fonts.HEADLINE,
            header_size_pt=header_pt,
            row_superheader_size_pt=superheader_pt,
        )

    @staticmethod
    def _build_cells(
        area: ContentArea,
        col_widths: List[int],
        row_heights: List[int],
    ) -> List[List[Box]]:
        cells: List[List[Box]] = []
        y = area.y
        for h in row_heights:
            row: List[Box] = []
            x = area.x
            for w in col_widths:
                row.append((x, y, w, h))
                x += w
            cells.append(row)
            y += h
        return cells
