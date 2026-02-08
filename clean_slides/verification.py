"""
Layout verification and reporting.

Checks a computed TableLayout for overlap, boundary violations, and
text-fit issues.  Returns a LayoutReport with categorised warnings.
"""

from dataclasses import dataclass, field
from typing import List

from .constants import OOXML_FONT_SCALE, DefaultColors, Fonts
from .content import Paragraph, normalize_cell
from .measure import (
    cell_content_height,
    column_right_pads,
    should_use_line_breaks,
    text_width_for_level,
    textbox_width,
)
from .spec import Box, ContentArea, TableLayout, TableSpec
from .text_metrics import TextMetrics

# ---------------------------------------------------------------------------
# Report dataclasses
# ---------------------------------------------------------------------------


@dataclass
class LayoutIssue:
    severity: str  # error | warning | info
    category: str  # overlap | boundary | fit
    message: str
    details: dict[str, object]


def _issue_list() -> List[LayoutIssue]:
    issues: List[LayoutIssue] = []
    return issues


@dataclass
class LayoutReport:
    issues: List[LayoutIssue] = field(default_factory=_issue_list)

    def add(self, issue: LayoutIssue) -> None:
        self.issues.append(issue)

    def has_errors(self) -> bool:
        return any(i.severity == "error" for i in self.issues)

    def has_warnings(self) -> bool:
        return any(i.severity == "warning" for i in self.issues)

    def summary(self) -> str:
        errors = sum(1 for i in self.issues if i.severity == "error")
        warnings = sum(1 for i in self.issues if i.severity == "warning")
        infos = sum(1 for i in self.issues if i.severity == "info")
        if not (errors or warnings or infos):
            return "OK"
        parts: List[str] = []
        if errors:
            parts.append(f"{errors} error(s)")
        if warnings:
            parts.append(f"{warnings} warning(s)")
        if infos:
            parts.append(f"{infos} info")
        return ", ".join(parts)

    def to_text(self, detail: bool = False) -> str:
        if not self.issues:
            return "OK"
        if not detail:
            return f"{self.summary()} (--detail for list)"
        return "\n".join(f"{i.severity.upper()}: {i.message} | {i.details}" for i in self.issues)

    def to_json(self) -> dict[str, object]:
        return {
            "summary": self.summary(),
            "issues": [
                {
                    "severity": i.severity,
                    "category": i.category,
                    "message": i.message,
                    "details": i.details,
                }
                for i in self.issues
            ],
        }


# ---------------------------------------------------------------------------
# Verifier
# ---------------------------------------------------------------------------


class LayoutVerifier:
    """Run checks against a computed layout."""

    def __init__(self, layout: TableLayout, area: ContentArea):
        self.layout = layout
        self.area = area

    def run_all(self, spec: TableSpec, metrics: TextMetrics) -> LayoutReport:
        report = LayoutReport()
        report.issues += self._check_overlaps()
        report.issues += self._check_boundaries()
        report.issues += self._check_text_fit(spec, metrics)
        return report

    # -- overlap --

    def _check_overlaps(self) -> List[LayoutIssue]:
        issues: List[LayoutIssue] = []
        cells = self.layout.cells
        for r1, row in enumerate(cells):
            for c1, (x1, y1, w1, h1) in enumerate(row):
                for r2 in range(r1, len(cells)):
                    for c2 in range(len(cells[r2])):
                        if (r2, c2) <= (r1, c1):
                            continue
                        x2, y2, w2, h2 = cells[r2][c2]
                        if min(x1 + w1, x2 + w2) > max(x1, x2) and min(y1 + h1, y2 + h2) > max(
                            y1, y2
                        ):
                            issues.append(
                                LayoutIssue(
                                    "error",
                                    "overlap",
                                    "Cell overlap detected",
                                    {"cell_a": (r1, c1), "cell_b": (r2, c2)},
                                )
                            )
        return issues

    # -- boundary --

    def _check_boundaries(self) -> List[LayoutIssue]:
        issues: List[LayoutIssue] = []
        a = self.area
        for r, row in enumerate(self.layout.cells):
            for c, (x, y, w, h) in enumerate(row):
                over = {
                    "left": max(a.x - x, 0),
                    "top": max(a.y - y, 0),
                    "right": max(x + w - (a.x + a.width), 0),
                    "bottom": max(y + h - (a.y + a.height), 0),
                }
                if any(over.values()):
                    issues.append(
                        LayoutIssue(
                            "error",
                            "boundary",
                            "Cell exceeds content area",
                            {"cell": (r, c), "overflow": over},
                        )
                    )
        return issues

    # -- text fit --

    def _check_text_fit(self, spec: TableSpec, metrics: TextMetrics) -> List[LayoutIssue]:
        issues: List[LayoutIssue] = []

        row_offset = spec.row_offset
        col_offset = spec.col_offset
        header_size_pt = int(self.layout.header_font_size / OOXML_FONT_SCALE)
        body_size_pt = int(self.layout.body_font_size / OOXML_FONT_SCALE)

        header_def = Paragraph(
            text="", lvl=0, font=Fonts.HEADLINE, size_pt=header_size_pt, color="accent1", bold=True
        )
        body_def = Paragraph(
            text="",
            lvl=spec.body_default_lvl,
            font=Fonts.BODY,
            size_pt=body_size_pt,
            color=DefaultColors.BODY_TEXT,
            bold=False,
        )

        right_pads = column_right_pads(
            len(self.layout.col_widths),
            self.layout.pad_top,
            spec.has_row_header,
        )

        def check(
            value: object,
            box: Box,
            default: Paragraph,
            kind: str,
            cell_id: tuple[int, int],
            right_pad: int = 0,
            use_line_breaks: bool = False,
        ) -> None:
            # Skip empty cells — renderer doesn't create text boxes for them
            if value == "" or value is None:
                return
            paragraphs = normalize_cell(value, default, parse_bullets=spec.parse_bullets)
            if not paragraphs:
                return
            _, _, w, h = box
            w = textbox_width(w, right_pad)

            # word-width check (verification-specific)
            for p in paragraphs:
                avail_w = text_width_for_level(w, p.lvl or 0)
                if avail_w <= 0:
                    issues.append(
                        LayoutIssue(
                            "warning",
                            "fit",
                            "Cell has no available width",
                            {"cell": cell_id, "kind": kind},
                        )
                    )
                    continue
                longest = metrics.longest_word(p.text)
                if longest:
                    lw = metrics.word_width(
                        longest,
                        p.font or default.font or Fonts.BODY,
                        p.size_pt or default.size_pt or body_size_pt,
                    )
                    if lw > avail_w:
                        issues.append(
                            LayoutIssue(
                                "warning",
                                "fit",
                                "Longest word exceeds column width",
                                {
                                    "cell": cell_id,
                                    "word": longest,
                                    "word_width": lw,
                                    "available": avail_w,
                                },
                            )
                        )

            # Determine if line breaks are used (matches renderer logic)
            lb = use_line_breaks or should_use_line_breaks(paragraphs)

            req = cell_content_height(
                paragraphs,
                w,
                metrics,
                self.layout.pad_top,
                self.layout.pad_bottom,
                default.font or Fonts.BODY,
                default.size_pt or body_size_pt,
                use_line_breaks=lb,
            )
            if req > h:
                issues.append(
                    LayoutIssue(
                        "warning",
                        "fit",
                        "Text exceeds cell height",
                        {"cell": cell_id, "required": req, "available": h, "overflow": req - h},
                    )
                )

        # column headers
        if spec.has_col_header:
            hdr_grid = 1 if spec.has_col_superheader else 0

            if spec.has_row_header and spec.row_header_col_header:
                check(
                    str(spec.row_header_col_header),
                    self.layout.cells[hdr_grid][0],
                    header_def,
                    "row_header_col_header",
                    (hdr_grid, 0),
                    right_pads[0] if right_pads else 0,
                    use_line_breaks=True,
                )

            for ci in range(spec.num_cols):
                text = (
                    str(spec.col_headers[ci])
                    if spec.col_headers and ci < len(spec.col_headers)
                    else ""
                )
                check(
                    text,
                    self.layout.cells[hdr_grid][ci + col_offset],
                    header_def,
                    "col_header",
                    (hdr_grid, ci),
                    right_pads[ci + col_offset],
                    use_line_breaks=True,
                )

        # row headers
        if spec.has_row_header:
            for ri in range(spec.num_rows):
                text = (
                    str(spec.row_headers[ri])
                    if spec.row_headers and ri < len(spec.row_headers)
                    else ""
                )
                check(
                    text,
                    self.layout.cells[ri + row_offset][0],
                    header_def,
                    "row_header",
                    (ri, 0),
                    right_pads[0] if right_pads else 0,
                    use_line_breaks=True,
                )

        # body cells
        for ri in range(spec.num_rows):
            row = spec.cells[ri] if spec.cells and ri < len(spec.cells) else []
            for ci in range(spec.num_cols):
                value = row[ci] if ci < len(row) else ""
                check(
                    value,
                    self.layout.cells[ri + row_offset][ci + col_offset],
                    body_def,
                    "body",
                    (ri, ci),
                    right_pads[ci + col_offset] if right_pads else 0,
                )

        return issues
