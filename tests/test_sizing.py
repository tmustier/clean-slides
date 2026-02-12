from __future__ import annotations

import unittest
from collections.abc import Callable
from typing import Any, cast

from clean_slides.constants import Fonts, TableDefaults
from clean_slides.measure import column_right_pads
from clean_slides.sizing import ColumnSizer, FontConfig, RowSizer
from clean_slides.solver import ConstraintSolver, SolveOptions
from clean_slides.spec import ChartDef, ChartRef, ColSuperHeader, ContentArea, RowGroup, TableSpec
from clean_slides.text_metrics import TextMetrics

FONTS = FontConfig(
    body_font=Fonts.BODY,
    body_size_pt=12,
    header_font=Fonts.HEADLINE,
    header_size_pt=12,
)

PAD = int(TableDefaults.CELL_PADDING)


def _protected_attr(target: object, name: str) -> Any:
    return object.__getattribute__(target, name)


class TestColumnSizer(unittest.TestCase):
    def test_widths_fit_within_area(self):
        spec = TableSpec(
            num_rows=2,
            num_cols=2,
            has_col_header=True,
            has_row_header=False,
            col_headers=["Header One", "Header Two"],
            cells=[["Alpha", "Beta"], ["Gamma", "Delta"]],
        )
        metrics = TextMetrics()
        area = ContentArea.from_layout("default")
        widths, warnings = ColumnSizer().size(spec, area.width, metrics, FONTS, pad_top=PAD)

        self.assertEqual(len(widths), 2)
        # Preferred algorithm: short content doesn't stretch to fill area
        self.assertLessEqual(sum(widths), area.width)
        self.assertGreater(sum(widths), 0)
        self.assertEqual(warnings, [])

    def test_default_algorithm_with_chart_columns_does_not_force_fill(self):
        spec = TableSpec(
            num_rows=2,
            num_cols=3,
            has_col_header=True,
            has_row_header=False,
            col_headers=["Driver", "Chart", "Note"],
            cells=[
                ["A", ChartRef(name="rev", index=1), "x"],
                ["B", ChartRef(name="rev", index=2), "y"],
            ],
            chart_defs={
                "rev": ChartDef(name="rev", type="bar", dir="horizontal", values=[10, 20]),
            },
        )
        metrics = TextMetrics()
        area = ContentArea.from_layout("default")

        widths, _warnings = ColumnSizer().size(spec, area.width, metrics, FONTS, pad_top=PAD)

        # Default auto-layout should keep no-wrap max widths and leave
        # trailing whitespace when content does not need full width.
        self.assertLess(sum(widths), area.width)
        # Chart-only columns should still keep enough intrinsic width for
        # bars + labels (must not collapse to padding-only widths).
        self.assertGreater(widths[1], 1_000_000)

    def test_single_span_col_superheader_constrains_column_width(self):
        spec = TableSpec(
            num_rows=1,
            num_cols=2,
            has_col_header=False,
            has_row_header=True,
            row_headers=["Group"],
            col_superheaders=[
                ColSuperHeader(label="Drivers", span=1),
                ColSuperHeader(label="CAGR", sub="%, FY26-33E", span=1),
                ColSuperHeader(label="Notes", span=1),
            ],
            cells=[["a", "b"]],
        )
        metrics = TextMetrics()
        area = ContentArea.from_layout("default")

        widths, _warnings = ColumnSizer().size(spec, area.width, metrics, FONTS, pad_top=PAD)

        label_w = int(
            metrics.text_width_no_wrap("CAGR", Fonts.HEADLINE, FONTS.header_size_pt)
            * float(_protected_attr(ColumnSizer, "_BOLD_FACTOR"))
        )
        sub_w = metrics.text_width_no_wrap("%, FY26-33E", Fonts.HEADLINE, FONTS.header_size_pt)
        expected_min = int(max(label_w, sub_w) * 1.10) + PAD * 4

        # Grid col 1 is the first body column (CAGR in this spec).
        self.assertGreaterEqual(widths[1], expected_min)

    def test_grouped_row_header_cap_stays_close_to_minimum(self):
        spec = TableSpec(
            num_rows=3,
            num_cols=3,
            has_col_header=False,
            has_row_header=True,
            col_widths=[1.0, 1.0, 1.0, 1.0],
            row_headers=["", "", ""],
            cells=[["", "alpha", "beta"], ["", "gamma", "delta"], ["", "epsilon", "zeta"]],
            groups=[
                RowGroup(header="Operating leverage / efficiency", num_rows=1, promoted=False),
                RowGroup(
                    header={
                        "text": "FY33E EBITDAaL at FY26 mgn.",
                        "sub": "(i.e. impact of Revenue growth)",
                    },
                    num_rows=1,
                    promoted=True,
                ),
                RowGroup(header="Price increases", num_rows=1, promoted=False),
            ],
        )
        metrics = TextMetrics()
        area = ContentArea.from_layout("default")
        sizer = ColumnSizer()

        min_widths = cast(
            Callable[..., list[int]],
            _protected_attr(sizer, "_min_widths"),
        )
        mins = min_widths(spec, area.width, metrics, FONTS, warnings=[])
        pads = column_right_pads(spec.num_cols + 1, PAD, spec.has_row_header)
        min_row_header_with_pad = mins[0] + pads[0]

        widths, _ = sizer.size(spec, area.width, metrics, FONTS, pad_top=PAD)

        # Grouped tables cap row-header growth at the minimum needed width.
        self.assertEqual(widths[0], min_row_header_with_pad)

    def test_multispan_col_superheader_does_not_inflate_row_header_cap(self):
        base_spec = TableSpec(
            num_rows=2,
            num_cols=3,
            has_col_header=False,
            has_row_header=True,
            row_headers=["", ""],
            cells=[["", "A", "B"], ["", "C", "D"]],
            groups=[
                RowGroup(header="Group A", num_rows=1, promoted=False),
                RowGroup(header="Group B", num_rows=1, promoted=False),
            ],
        )
        with_super = TableSpec(
            num_rows=base_spec.num_rows,
            num_cols=base_spec.num_cols,
            has_col_header=base_spec.has_col_header,
            has_row_header=base_spec.has_row_header,
            row_headers=base_spec.row_headers,
            cells=base_spec.cells,
            groups=base_spec.groups,
            col_superheaders=[
                ColSuperHeader(label="Very long superheader label across three columns", span=3),
                ColSuperHeader(label="Other", span=1),
            ],
        )

        metrics = TextMetrics()
        sizer = ColumnSizer()

        row_header_pref_width = cast(
            Callable[..., int],
            _protected_attr(sizer, "_row_header_preferred_width"),
        )
        pref_base = row_header_pref_width(base_spec, metrics, FONTS)
        pref_with_super = row_header_pref_width(with_super, metrics, FONTS)

        self.assertEqual(pref_with_super, pref_base)


class TestRowSizer(unittest.TestCase):
    def test_heights_sum_to_area(self):
        spec = TableSpec(
            num_rows=2,
            num_cols=2,
            has_col_header=True,
            has_row_header=False,
            col_headers=["Header One", "Header Two"],
            cells=[["Alpha", "Beta"], ["Gamma", "Delta"]],
        )
        metrics = TextMetrics()
        area = ContentArea.from_layout("default")
        widths, _ = ColumnSizer().size(spec, area.width, metrics, FONTS, pad_top=PAD)
        heights, warnings = RowSizer().size(spec, widths, area.height, metrics, FONTS, PAD, PAD)

        self.assertEqual(len(heights), 3)  # 1 header + 2 body
        self.assertEqual(sum(heights), area.height)
        self.assertEqual(warnings, [])

    def test_rebalance_shrinks_across_rows_not_just_last(self):
        # Mirrors the overflow pattern seen on long grouped/chart tables:
        # all rows initially tall, but total exceeds target by a large margin.
        body = [515730] * 8 + [515733]
        target = 4301632
        min_h = int(TableDefaults.MIN_ROW_HEIGHT)

        rebalance_body_heights = cast(
            Callable[..., None],
            _protected_attr(RowSizer, "_rebalance_body_heights"),
        )
        rebalance_body_heights(body, target, min_h)

        self.assertEqual(sum(body), target)
        self.assertGreaterEqual(min(body), min_h)
        # Regression check: last row must not collapse to a tiny leftover.
        self.assertGreater(body[-1], 400000)

    def test_rebalance_even_split_when_target_below_min_total(self):
        # If target can't satisfy min row height for all rows, split evenly
        # instead of keeping mins and collapsing the last row.
        body = [int(TableDefaults.MIN_ROW_HEIGHT)] * 3
        target = 900000
        min_h = int(TableDefaults.MIN_ROW_HEIGHT)

        rebalance_body_heights = cast(
            Callable[..., None],
            _protected_attr(RowSizer, "_rebalance_body_heights"),
        )
        rebalance_body_heights(body, target, min_h)

        self.assertEqual(sum(body), target)
        self.assertLessEqual(max(body) - min(body), 1)

    def test_grouped_header_requirement_inflates_required_rows(self):
        spec = TableSpec(
            num_rows=2,
            num_cols=2,
            has_col_header=False,
            has_row_header=True,
            row_headers=["", ""],
            cells=[["", "A"], ["", "B"]],
            groups=[
                RowGroup(
                    header={
                        "text": "FY33E EBITDAaL at FY26 mgn.",
                        "sub": "(i.e. impact of Revenue growth)",
                    },
                    num_rows=1,
                    promoted=True,
                ),
                RowGroup(header="Drivers", num_rows=1),
            ],
        )
        required = [200000, 500000]
        text_widths = [300000, 300000, 300000]

        inflate_grouped_header_requirements = cast(
            Callable[..., None],
            _protected_attr(RowSizer, "_inflate_grouped_header_requirements"),
        )
        inflate_grouped_header_requirements(
            spec,
            required,
            text_widths,
            TextMetrics(),
            FONTS,
            pt=PAD,
            pb=PAD,
        )

        self.assertGreaterEqual(required[0], 200000)
        self.assertEqual(required[1], 500000)


class TestSolverFontResolution(unittest.TestCase):
    def test_promoted_group_uses_uniform_body_size_for_all_row_superheaders(self):
        spec = TableSpec(
            num_rows=2,
            num_cols=2,
            has_col_header=False,
            has_row_header=True,
            row_headers=["", ""],
            cells=[["", "A"], ["", "B"]],
            groups=[
                RowGroup(header="Start", num_rows=1, promoted=True),
                RowGroup(header="Drivers", num_rows=1),
            ],
        )
        solver = ConstraintSolver(TextMetrics())

        resolve_fonts = cast(
            Callable[..., FontConfig],
            _protected_attr(solver, "_resolve_fonts"),
        )
        fonts = resolve_fonts(
            spec,
            SolveOptions(body_font_pt=12, header_font_pt=16),
        )

        self.assertEqual(fonts.header_size_pt, 16)
        self.assertEqual(fonts.effective_row_superheader_size_pt, 12)

    def test_non_promoted_groups_keep_header_size_for_row_superheaders(self):
        spec = TableSpec(
            num_rows=2,
            num_cols=2,
            has_col_header=False,
            has_row_header=True,
            row_headers=["", ""],
            cells=[["", "A"], ["", "B"]],
            groups=[
                RowGroup(header="Group A", num_rows=1, promoted=False),
                RowGroup(header="Group B", num_rows=1, promoted=False),
            ],
        )
        solver = ConstraintSolver(TextMetrics())

        resolve_fonts = cast(
            Callable[..., FontConfig],
            _protected_attr(solver, "_resolve_fonts"),
        )
        fonts = resolve_fonts(
            spec,
            SolveOptions(body_font_pt=12, header_font_pt=16),
        )

        self.assertEqual(fonts.effective_row_superheader_size_pt, 16)


if __name__ == "__main__":
    unittest.main()
