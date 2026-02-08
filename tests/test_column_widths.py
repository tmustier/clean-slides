from clean_slides.sizing import ColumnSizer, FontConfig
from clean_slides.spec import TableSpec
from clean_slides.text_metrics import TextMetrics


def _fonts() -> FontConfig:
    return FontConfig(
        body_font="Calibri", body_size_pt=12, header_font="Calibri", header_size_pt=12
    )


def test_column_widths_full_list_accepted() -> None:
    spec = TableSpec.from_dict(
        {
            "table": {
                "rows": 3,
                "cols": 3,
                "has_col_header": True,
                "has_row_header": True,
                "row_header_col_header": "Aspect",
                "col_headers": ["A", "B"],
                "row_headers": ["R1", "R2"],
                "cells": [["x", "y"], ["xx", "yy"]],
                # full weights: row-header + 2 body cols
                "column_widths": [1.0, 2.0, 3.0],
            }
        }
    )

    sizer = ColumnSizer()
    widths, _warnings = sizer.size(
        spec, area_width=10_000_000, metrics=TextMetrics(), fonts=_fonts()
    )
    assert len(widths) == 3


def test_column_widths_body_only_list() -> None:
    spec = TableSpec.from_dict(
        {
            "table": {
                "rows": 3,
                "cols": 3,
                "has_col_header": True,
                "has_row_header": True,
                "row_header_col_header": "Aspect",
                "col_headers": ["A", "B"],
                "row_headers": ["R1", "R2"],
                "cells": [["x", "y"], ["xx", "yy"]],
                # body-only weights
                "column_widths": [2.0, 3.0],
            }
        }
    )

    sizer = ColumnSizer()
    widths, _warnings = sizer.size(
        spec, area_width=10_000_000, metrics=TextMetrics(), fonts=_fonts()
    )
    assert len(widths) == 3


def test_column_widths_full_list_caps_row_header_slack() -> None:
    spec = TableSpec.from_dict(
        {
            "table": {
                "rows": 3,
                "cols": 3,
                "has_col_header": True,
                "has_row_header": True,
                "row_header_col_header": "Aspect",
                "col_headers": ["A", "B"],
                "row_headers": ["A", "B"],
                "cells": [["x", "y"], ["xx", "yy"]],
                # Extreme preference for row-header column.
                # The sizer should cap row-header growth to its preferred width.
                "column_widths": [1000.0, 1.0, 1.0],
            }
        }
    )

    sizer = ColumnSizer()
    widths, _warnings = sizer.size(
        spec, area_width=10_000_000, metrics=TextMetrics(), fonts=_fonts()
    )

    assert len(widths) == 3
    assert sum(widths) == 10_000_000
    assert widths[0] < widths[1]


def test_column_widths_equal_mode() -> None:
    spec = TableSpec.from_dict(
        {
            "table": {
                "rows": 3,
                "cols": 3,
                "has_col_header": True,
                "has_row_header": False,
                "col_headers": ["A", "B", "C"],
                "cells": [
                    ["short", "much longer text here", "x"],
                    ["a", "another long cell value", "y"],
                ],
                "column_widths": "equal",
            }
        }
    )

    sizer = ColumnSizer()
    widths, _warnings = sizer.size(
        spec, area_width=10_000_000, metrics=TextMetrics(), fonts=_fonts(), pad_top=45_720
    )

    # All three columns should get similar widths (within 20% of each other).
    assert max(widths) < min(widths) * 1.5


def test_row_header_two_word_label_min_width_prefers_single_line() -> None:
    spec = TableSpec.from_dict(
        {
            "table": {
                "rows": 3,
                "cols": 2,
                "has_col_header": True,
                "has_row_header": True,
                "row_header_col_header": "Aspect",
                "col_headers": ["A"],
                "row_headers": ["Two Words", "Another Label"],
                "cells": [["x"], ["y"]],
            }
        }
    )

    metrics = TextMetrics()
    sizer = ColumnSizer()
    fonts = _fonts()

    pad_top = 45_720
    widths, _warnings = sizer.size(
        spec,
        area_width=10_000_000,
        metrics=metrics,
        fonts=fonts,
        pad_top=pad_top,
    )

    row_header_gap = pad_top * 3
    avail = widths[0] - row_header_gap
    assert metrics.lines_needed("Two Words", avail, fonts.header_font, fonts.header_size_pt) <= 1


def test_row_header_min_width_fits_any_adjacent_word_pair() -> None:
    spec = TableSpec.from_dict(
        {
            "table": {
                "rows": 3,
                "cols": 2,
                "has_col_header": True,
                "has_row_header": True,
                "row_header_col_header": "Aspect",
                "col_headers": ["A"],
                "row_headers": ["One Two Three", "Alpha Beta Gamma"],
                "cells": [["x"], ["y"]],
            }
        }
    )

    metrics = TextMetrics()
    sizer = ColumnSizer()
    fonts = _fonts()

    pad_top = 45_720
    widths, _warnings = sizer.size(
        spec,
        area_width=10_000_000,
        metrics=metrics,
        fonts=fonts,
        pad_top=pad_top,
    )

    row_header_gap = pad_top * 3
    avail = widths[0] - row_header_gap

    for pair in ["One Two", "Two Three", "Alpha Beta", "Beta Gamma"]:
        assert metrics.text_width_no_wrap(pair, fonts.header_font, fonts.header_size_pt) <= avail


def test_default_short_column_stays_compact() -> None:
    spec = TableSpec.from_dict(
        {
            "table": {
                "rows": 3,
                "cols": 2,
                "has_col_header": True,
                "has_row_header": False,
                "col_headers": ["Short", "Long Header Column"],
                "cells": [
                    ["ok", "This is a much longer sentence that needs wrapping"],
                    ["hi", "Another fairly long piece of text for this column"],
                    ["no", "Yet more text to demonstrate the width difference"],
                ],
            }
        }
    )

    sizer = ColumnSizer()
    widths, _warnings = sizer.size(
        spec,
        area_width=10_000_000,
        metrics=TextMetrics(),
        fonts=_fonts(),
        pad_top=45_720,
    )

    assert widths[0] < widths[1] * 0.5


def test_default_no_wrap_when_fits() -> None:
    spec = TableSpec.from_dict(
        {
            "table": {
                "rows": 2,
                "cols": 2,
                "has_col_header": True,
                "has_row_header": False,
                "col_headers": ["A", "B"],
                "cells": [
                    ["hello", "world"],
                    ["foo", "bar"],
                ],
            }
        }
    )

    sizer = ColumnSizer()
    widths, _warnings = sizer.size(
        spec,
        area_width=10_000_000,
        metrics=TextMetrics(),
        fonts=_fonts(),
        pad_top=45_720,
    )

    assert sum(widths) < 10_000_000 * 0.5
