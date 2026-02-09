from __future__ import annotations

from clean_slides.spec import TableSpec


def test_rows_cols_include_headers() -> None:
    """table.rows/table.cols are total counts including header rows/cols."""
    data = {
        "table": {
            "rows": 3,  # 1 header row + 2 body rows
            "cols": 4,  # 1 row-header col + 3 body cols
            "has_col_header": True,
            "has_row_header": True,
            "row_header_col_header": "Row",
            "col_headers": ["A", "B", "C"],
            "row_headers": ["R1", "R2"],
            "cells": [
                ["a1", "b1", "c1"],
                ["a2", "b2", "c2"],
            ],
        }
    }

    spec = TableSpec.from_dict(data)
    assert spec.num_rows == 2
    assert spec.num_cols == 3
    assert spec.has_row_header is True
    assert spec.has_col_header is True


def test_auto_extract_row_headers_from_cells() -> None:
    """has_row_header without row_headers extracts first column from cells."""
    data = {
        "table": {
            "rows": 4,  # 1 col header + 3 body
            "cols": 2,  # 1 row header + 1 body
            "has_col_header": True,
            "has_row_header": True,
            "col_headers": ["Feature", "Why it matters"],
            "cells": [
                ["Eyes", "Pinhead-sized and nearly useless"],
                ["Feet", "Fully webbed with stiff hair fringes"],
                ["Tail", "Laterally flattened rudder"],
            ],
        }
    }

    spec = TableSpec.from_dict(data)
    assert spec.num_rows == 3
    assert spec.num_cols == 1
    assert spec.row_headers == ["Eyes", "Feet", "Tail"]
    assert spec.row_header_col_header == "Feature"
    assert spec.col_headers == ["Why it matters"]
    assert spec.cells is not None
    assert spec.cells[0] == ["Pinhead-sized and nearly useless"]
    assert spec.cells[1] == ["Fully webbed with stiff hair fringes"]
    assert spec.cells[2] == ["Laterally flattened rudder"]


def test_explicit_row_headers_not_auto_extracted() -> None:
    """When row_headers is provided, cells are NOT modified."""
    data = {
        "table": {
            "rows": 4,
            "cols": 3,
            "has_col_header": True,
            "has_row_header": True,
            "col_headers": ["A", "B"],
            "row_headers": ["R1", "R2", "R3"],
            "cells": [
                ["a1", "b1"],
                ["a2", "b2"],
                ["a3", "b3"],
            ],
        }
    }

    spec = TableSpec.from_dict(data)
    assert spec.num_rows == 3
    assert spec.num_cols == 2
    assert spec.row_headers == ["R1", "R2", "R3"]
    assert spec.cells is not None
    assert spec.cells[0] == ["a1", "b1"]
