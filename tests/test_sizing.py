from __future__ import annotations

import unittest

from clean_slides.constants import Fonts, TableDefaults
from clean_slides.sizing import ColumnSizer, FontConfig, RowSizer
from clean_slides.spec import ContentArea, TableSpec
from clean_slides.text_metrics import TextMetrics

FONTS = FontConfig(
    body_font=Fonts.BODY,
    body_size_pt=12,
    header_font=Fonts.HEADLINE,
    header_size_pt=12,
)

PAD = int(TableDefaults.CELL_PADDING)


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


if __name__ == "__main__":
    unittest.main()
