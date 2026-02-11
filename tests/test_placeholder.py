"""Tests for placeholder content generation."""

from __future__ import annotations

from clean_slides.placeholder import fill_placeholders, placeholder_body, placeholder_title
from clean_slides.spec import RowGroup, TableSpec


def _flat_spec(num_rows: int, num_cols: int, row_headers: list[str] | None = None) -> TableSpec:
    """Build a minimal flat table spec."""
    return TableSpec(
        num_rows=num_rows,
        num_cols=num_cols,
        has_col_header=True,
        has_row_header=row_headers is not None,
        row_headers=row_headers,
        cells=[],
    )


def _grouped_spec(
    groups: list[RowGroup],
    num_cols: int,
    row_headers: list[str] | None = None,
) -> TableSpec:
    """Build a minimal grouped table spec."""
    total_rows = sum(g.num_rows for g in groups)
    return TableSpec(
        num_rows=total_rows,
        num_cols=num_cols,
        has_col_header=True,
        has_row_header=True,
        row_headers=row_headers or [g.header for g in groups],
        cells=[[""] * num_cols for _ in range(total_rows)],
        groups=groups,
    )


class TestPlaceholderHelpers:
    def test_placeholder_title_default(self) -> None:
        title = placeholder_title()
        assert len(title.split()) == 4

    def test_placeholder_body_default(self) -> None:
        body = placeholder_body()
        assert len(body) > 0


class TestFillPlaceholdersFlat:
    """Flat (non-grouped) tables."""

    def test_fills_missing_col_headers(self) -> None:
        spec = TableSpec(
            num_rows=2,
            num_cols=3,
            has_col_header=True,
            col_headers=["A"],
        )
        filled = fill_placeholders(spec)
        assert filled.col_headers is not None
        assert len(filled.col_headers) == 3
        assert filled.col_headers[0] == "A"

    def test_fills_missing_row_headers(self) -> None:
        spec = _flat_spec(3, 2, row_headers=["R1"])
        filled = fill_placeholders(spec)
        assert filled.row_headers is not None
        assert len(filled.row_headers) == 3
        assert filled.row_headers[0] == "R1"

    def test_fills_missing_cells(self) -> None:
        spec = _flat_spec(2, 2)
        filled = fill_placeholders(spec)
        assert filled.cells is not None
        assert len(filled.cells) == 2
        assert len(filled.cells[0]) == 2
        assert len(filled.cells[0][0]) > 0

    def test_preserves_existing_cell_content(self) -> None:
        spec = TableSpec(
            num_rows=1,
            num_cols=2,
            has_col_header=False,
            cells=[["keep this", ""]],
        )
        filled = fill_placeholders(spec)
        assert filled.cells is not None
        assert filled.cells[0][0] == "keep this"
        assert filled.cells[0][1] != ""  # placeholder inserted


class TestFillPlaceholdersGrouped:
    """Grouped (superheader) tables — row_headers = one per group."""

    def test_does_not_add_extra_row_headers(self) -> None:
        """The bug: 3 groups with 4 total sub-rows should NOT pad
        row_headers to 4.  row_headers should stay at 3 (one per group)."""
        groups = [
            RowGroup(header="Inspect", num_rows=2),
            RowGroup(header="Generate", num_rows=1),
            RowGroup(header="Verify", num_rows=1),
        ]
        spec = _grouped_spec(groups, num_cols=3)
        assert spec.num_rows == 4  # total sub-rows
        assert len(spec.row_headers or []) == 3  # one per group

        filled = fill_placeholders(spec)
        assert filled.row_headers is not None
        assert len(filled.row_headers) == 3  # still 3, NOT 4

    def test_fills_missing_group_headers(self) -> None:
        """If groups have fewer row_headers than groups, fill to group count."""
        groups = [
            RowGroup(header="A", num_rows=1),
            RowGroup(header="B", num_rows=2),
            RowGroup(header="C", num_rows=1),
        ]
        spec = _grouped_spec(groups, num_cols=2, row_headers=["A"])
        filled = fill_placeholders(spec)
        assert filled.row_headers is not None
        assert len(filled.row_headers) == 3
        assert filled.row_headers[0] == "A"

    def test_fills_cells_to_total_sub_rows(self) -> None:
        """Cells should still be padded to num_rows (total sub-rows)."""
        groups = [
            RowGroup(header="X", num_rows=2),
            RowGroup(header="Y", num_rows=1),
        ]
        spec = _grouped_spec(groups, num_cols=2)
        filled = fill_placeholders(spec)
        assert filled.cells is not None
        assert len(filled.cells) == 3  # total sub-rows

    def test_preserves_blank_first_cell_for_promoted_group(self) -> None:
        groups = [
            RowGroup(header="Start", num_rows=1, promoted=True),
            RowGroup(header="Drivers", num_rows=1),
        ]
        spec = _grouped_spec(groups, num_cols=2)
        assert spec.cells is not None
        spec.cells[0][0] = ""
        spec.cells[0][1] = "wf-1"
        spec.cells[1][0] = "CPI"
        spec.cells[1][1] = "wf-2"

        filled = fill_placeholders(spec)
        assert filled.cells is not None
        assert filled.cells[0][0] == ""
