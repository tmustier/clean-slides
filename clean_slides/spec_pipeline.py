"""Shared YAML spec loading/parsing/validation helpers for CLI commands."""

from __future__ import annotations

from typing import cast

import yaml
from typing_extensions import TypeGuard

from .constants import FontSizes, TableDefaults
from .solver import SolveOptions
from .spec import ChartRef, ContentArea, TableSpec

YamlDict = dict[str, object]


def _is_object_list(value: object) -> TypeGuard[list[object]]:
    return isinstance(value, list)


def _is_str_object_dict(value: object) -> TypeGuard[dict[str, object]]:
    return isinstance(value, dict)


def _cells_can_provide_row_headers(raw_cells: object) -> bool:
    """True when *raw_cells* is a non-empty list of lists each with ≥ 2 elements."""
    if not isinstance(raw_cells, list):
        return False
    rows = cast(list[object], raw_cells)
    if len(rows) == 0:
        return False
    return all(not (not isinstance(row, list) or len(cast(list[object], row)) < 2) for row in rows)


def load_yaml(path: str) -> YamlDict:
    """Load YAML from disk."""
    with open(path, encoding="utf-8") as handle:
        parsed: object = yaml.safe_load(handle)

    # YAML root must be a mapping for our spec format. If it's not, treat it as
    # empty so downstream validation can surface a friendly error.
    if _is_str_object_dict(parsed):
        return parsed

    return {}


def _to_int(value: object) -> int:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(value)
    return int(str(value))


def _parse_content_area(data: YamlDict, layout_override: str | None = None) -> ContentArea:
    if "content_area" in data:
        area_obj = data.get("content_area")
        area: dict[str, object] = area_obj if _is_str_object_dict(area_obj) else {}

        return ContentArea(
            x=_to_int(area.get("x", 0)),
            y=_to_int(area.get("y", 0)),
            width=_to_int(area.get("width", 0)),
            height=_to_int(area.get("height", 0)),
        )

    # Content-area layout (not slide-master layout)
    layout_obj = data.get("content_layout")
    if layout_obj is None:
        layout_obj = data.get("layout")

    if isinstance(layout_obj, str) and layout_obj:
        layout = layout_obj
    elif layout_override is not None:
        layout = layout_override
    else:
        layout = "default"

    return ContentArea.from_layout(layout)


def _parse_options(table: dict[str, object]) -> SolveOptions:
    fonts_obj = table.get("fonts")
    fonts: dict[str, object] = fonts_obj if _is_str_object_dict(fonts_obj) else {}

    padding_obj = table.get("padding")
    padding: dict[str, object] = padding_obj if _is_str_object_dict(padding_obj) else {}

    # Default to template config font size (16pt) instead of hardcoded 12pt
    body_raw = fonts.get("body", FontSizes.TABLE_BODY)
    header_raw = fonts.get("header", FontSizes.TABLE_HEADER)

    default_pad = int(TableDefaults.CELL_PADDING)

    return SolveOptions(
        body_font_pt=_to_int(body_raw),
        header_font_pt=_to_int(header_raw),
        min_font_pt=_to_int(fonts.get("min", 10)),
        max_font_pt=_to_int(fonts.get("max", 16)),
        pad_top=_to_int(padding.get("top", default_pad)),
        pad_bottom=_to_int(padding.get("bottom", default_pad)),
    )


def parse_spec(
    data: YamlDict,
    layout_override: str | None = None,
) -> tuple[TableSpec, ContentArea, SolveOptions, bool]:
    """Parse YAML dict into structured objects."""

    table_obj = data.get("table")
    table: dict[str, object] = table_obj if _is_str_object_dict(table_obj) else {}

    spec = TableSpec.from_dict(data)
    area = _parse_content_area(data, layout_override=layout_override)
    options = _parse_options(table)

    placeholders_raw = table.get("placeholders", True)
    placeholders = bool(placeholders_raw)

    return spec, area, options, placeholders


def _validate_chart_refs(spec: TableSpec) -> list[str]:
    """Validate all chart cell references against chart definitions.

    Returns a list of error messages (empty if valid).
    """
    if not spec.chart_defs or spec.cells is None:
        return []

    errors: list[str] = []
    chart_defs = spec.chart_defs

    # Collect all refs by chart name → list of (row, col, index)
    refs_by_chart: dict[str, list[tuple[int, int, int]]] = {}
    for ri, row in enumerate(spec.cells):
        for ci, value in enumerate(row):
            if not isinstance(value, ChartRef):
                continue

            # 1. Unknown chart name
            if value.name not in chart_defs:
                errors.append(
                    f"Cell ({ri + 1}, {ci + 1}): chart '{value.name}' is not defined in charts:"
                )
                continue

            chart = chart_defs[value.name]

            # 2. Index out of range
            if value.index < 1 or value.index > len(chart.values):
                errors.append(
                    f"Cell ({ri + 1}, {ci + 1}): {value.name}-{value.index} "
                    f"but {value.name} has only {len(chart.values)} values"
                )
                continue

            refs_by_chart.setdefault(value.name, []).append((ri, ci, value.index))

    # Per-chart structural validation
    for chart_name, ref_list in refs_by_chart.items():
        chart = chart_defs[chart_name]
        num_values = len(chart.values)

        # 3. Duplicate indices (same chart, same index, different cell)
        seen_indices: dict[int, tuple[int, int]] = {}
        for ri, ci, idx in ref_list:
            if idx in seen_indices:
                prev_r, prev_c = seen_indices[idx]
                errors.append(
                    f"{chart_name}-{idx} appears in cells ({prev_r + 1}, {prev_c + 1}) "
                    f"and ({ri + 1}, {ci + 1})"
                )
            else:
                seen_indices[idx] = (ri, ci)

        # 4. Index gaps
        used_indices = sorted(seen_indices.keys())
        expected = list(range(1, num_values + 1))
        if used_indices and used_indices != expected[: len(used_indices)]:
            missing = [i for i in range(1, max(used_indices) + 1) if i not in seen_indices]
            if missing:
                errors.append(
                    f"{chart_name} has indices {used_indices} — missing "
                    f"{'index' if len(missing) == 1 else 'indices'} {missing}"
                )

        # 5. Direction mismatch — horizontal charts should span rows (same column),
        #    vertical charts should span columns (same row)
        rows_used = sorted({r for r, _, _ in ref_list})
        cols_used = sorted({c for _, c, _ in ref_list})

        if chart.dir == "horizontal" and len(cols_used) > 1:
            errors.append(
                f"{chart_name} (horizontal) has refs in columns {[c + 1 for c in cols_used]} "
                f"— expected all in the same column"
            )
        elif chart.dir == "vertical" and len(rows_used) > 1:
            errors.append(
                f"{chart_name} (vertical) has refs in rows {[r + 1 for r in rows_used]} "
                f"— expected all in the same row"
            )

        # 6. Non-contiguous cells
        if chart.dir == "horizontal" and len(rows_used) > 1:
            for i in range(len(rows_used) - 1):
                if rows_used[i + 1] - rows_used[i] != 1:
                    errors.append(
                        f"{chart_name} refs in column {cols_used[0] + 1} at rows "
                        f"{[r + 1 for r in rows_used]} — rows are not contiguous"
                    )
                    break
        elif chart.dir == "vertical" and len(cols_used) > 1:
            for i in range(len(cols_used) - 1):
                if cols_used[i + 1] - cols_used[i] != 1:
                    errors.append(
                        f"{chart_name} refs in row {rows_used[0] + 1} at columns "
                        f"{[c + 1 for c in cols_used]} — columns are not contiguous"
                    )
                    break

    return errors


def validate_spec(data: YamlDict) -> tuple[list[str], list[str]]:
    """Return (errors, warnings)."""
    errors: list[str] = []
    warnings: list[str] = []

    if not data.get("title"):
        warnings.append("Missing 'title' — slide will have no title text")

    table_obj = data.get("table")
    if not _is_str_object_dict(table_obj) or not table_obj:
        # No table section — metadata-only slide (title, section, quote, etc.).
        # This is valid; generate will create the slide with the correct layout
        # and fill placeholders without rendering a table.
        return errors, warnings

    table = table_obj

    raw_groups_raw: object = table.get("row_groups")
    rows_obj: object = table.get("rows")
    cols_obj: object = table.get("cols")

    if raw_groups_raw:
        if cols_obj is None:
            errors.append("table.cols is required")

        if not _is_object_list(raw_groups_raw) or not raw_groups_raw:
            errors.append("table.row_groups must be a non-empty list")
    else:
        if rows_obj is None:
            errors.append("table.rows is required")
        if cols_obj is None:
            errors.append("table.cols is required")

    # Basic numeric checks (best-effort; actual dimension validation happens in TableSpec.from_dict)
    if isinstance(rows_obj, int) and rows_obj <= 0:
        errors.append("table.rows must be > 0")
    if isinstance(cols_obj, int) and cols_obj <= 0:
        errors.append("table.cols must be > 0")

    if "has_col_header" not in table:
        warnings.append("table.has_col_header not specified (defaulting to true)")
    if not raw_groups_raw and "has_row_header" not in table:
        warnings.append("table.has_row_header not specified (defaulting to false)")

    # Parse normalized spec (body-only dims) so validation stays aligned with renderer/solver.
    try:
        spec = TableSpec.from_dict(data)  # handles row_groups + header-inclusive rows/cols
    except ValueError as e:
        errors.append(str(e))
        return errors, warnings

    if not spec.has_col_header:
        warnings.append("Column headers disabled (table.has_col_header = false)")
    if not spec.has_row_header and not raw_groups_raw:
        warnings.append("Row headers disabled (table.has_row_header = false)")

    body_rows = spec.num_rows
    body_cols = spec.num_cols

    # --- headers ---

    raw_col_headers = table.get("col_headers")
    raw_row_headers = table.get("row_headers")
    has_row_header_col = spec.has_row_header

    if spec.has_col_header and not raw_col_headers:
        warnings.append("Column headers missing; placeholders will be generated")
    # Auto-extract: when has_row_header but no row_headers list, the parser
    # extracts them from the first column of cells.  Only warn if cells are
    # also missing (truly no source for row headers).
    raw_cells: object = table.get("cells")
    if has_row_header_col and not raw_row_headers and not raw_groups_raw:
        cells_can_provide = _cells_can_provide_row_headers(raw_cells)
        if not cells_can_provide:
            warnings.append("Row headers missing; placeholders will be generated")

    # col_headers should match *body* columns.
    # Convenience: if has_row_header and row_header_col_header not provided, allow one extra.
    if _is_object_list(raw_col_headers):
        expected = body_cols
        allow = expected + (
            1 if (has_row_header_col and not table.get("row_header_col_header")) else 0
        )
        if len(raw_col_headers) not in {expected, allow}:
            warnings.append(
                f"col_headers length ({len(raw_col_headers)}) does not match expected body columns ({expected})"
            )

    # row_headers should match *body* rows (except grouped mode where row headers are groups).
    if (
        not spec.is_grouped
        and _is_object_list(raw_row_headers)
        and len(raw_row_headers) != body_rows
    ):
        warnings.append(
            f"row_headers length ({len(raw_row_headers)}) does not match expected body rows ({body_rows})"
        )

    # --- column widths ---

    cw: object = table.get("column_widths")

    if isinstance(cw, str) and cw.lower() not in {"equal"}:
        warnings.append(
            f"column_widths '{cw}' is not recognized (use 'equal' or a list of numbers)"
        )
    elif _is_object_list(cw):
        total_cols = body_cols + (1 if has_row_header_col else 0)
        if has_row_header_col:
            if len(cw) not in {body_cols, total_cols}:
                warnings.append(
                    f"column_widths length ({len(cw)}) does not match expected columns ({body_cols} body, or {total_cols} including row header)"
                )
        else:
            if len(cw) != body_cols:
                warnings.append(
                    f"column_widths length ({len(cw)}) does not match expected body columns ({body_cols})"
                )

    # --- cells grid ---

    cells_obj = table.get("cells")
    # When auto-extracting row headers from cells, the first column is consumed
    # by the parser — account for that in width checks.
    auto_extract_row_hdr = (
        has_row_header_col
        and not raw_row_headers
        and not raw_groups_raw
        and _cells_can_provide_row_headers(raw_cells)
    )
    effective_body_cols = body_cols + (1 if auto_extract_row_hdr else 0)

    if _is_object_list(cells_obj):
        if len(cells_obj) > body_rows:
            warnings.append("cells has more rows than expected body rows")

        for idx, row_obj in enumerate(cells_obj):
            if not _is_object_list(row_obj):
                warnings.append(f"cells row {idx + 1} is not a list")
                continue
            if len(row_obj) > effective_body_cols:
                warnings.append(f"cells row {idx + 1} has more columns than expected body columns")

    # --- layout keys ---

    # Validate content_layout by trying to construct it
    cl_raw = data.get("content_layout") or data.get("layout")
    if cl_raw is not None and isinstance(cl_raw, str) and cl_raw:
        try:
            ContentArea.from_layout(cl_raw)
        except ValueError:
            warnings.append(f"content_layout '{cl_raw}' is not recognized")

    content_area_obj = data.get("content_area")
    if content_area_obj:
        if not _is_str_object_dict(content_area_obj):
            errors.append("content_area must be a mapping when set")
        else:
            for key in ("x", "y", "width", "height"):
                if key not in content_area_obj:
                    errors.append(f"content_area.{key} is required when content_area is set")

    # --- fonts / bullets ---

    fonts_obj = table.get("fonts")
    fonts: dict[str, object] = fonts_obj if _is_str_object_dict(fonts_obj) else {}

    body = _to_int(fonts.get("body", FontSizes.TABLE_BODY))
    header = _to_int(fonts.get("header", body))
    min_font = _to_int(fonts.get("min", 10))
    max_font = _to_int(fonts.get("max", 16))

    if header < body:
        warnings.append("header font is smaller than body font")
    if body < min_font or body > max_font:
        warnings.append("body font outside min/max bounds")

    body_default_lvl = table.get("body_default_lvl", 0)
    if not isinstance(body_default_lvl, int) or body_default_lvl < 0 or body_default_lvl > 8:
        warnings.append("body_default_lvl should be between 0 and 8")

    parse_bullets = table.get("parse_bullets", True)
    if not isinstance(parse_bullets, bool):
        warnings.append("parse_bullets should be true or false")

    # --- icons legend ---

    if spec.icons is not None and spec.icons.show_legend:
        legend_items = list(spec.icons.values.items())
        if len(legend_items) > 5:
            warnings.append(
                "icons.legend has more than 5 items; consider limiting legend to 5 for readability"
            )
        colors = [c for _, c in legend_items]
        if len(set(colors)) != len(colors):
            warnings.append(
                "icons.legend contains duplicate colors; legend should map 1:1 with meaning"
            )

    # --- chart cell references ---

    chart_errors = _validate_chart_refs(spec)
    errors.extend(chart_errors)

    return errors, warnings


def preview_spec(data: YamlDict) -> str:
    """Generate a text preview of table structure."""

    table_obj = data.get("table")
    table: dict[str, object] = table_obj if _is_str_object_dict(table_obj) else {}

    rows = table.get("rows")
    cols = table.get("cols")
    has_col_header = bool(table.get("has_col_header", True))
    has_row_header = bool(table.get("has_row_header", False))

    lines: list[str] = ["=" * 60, "TABLE PREVIEW", "=" * 60]
    lines.append(f"Rows: {rows} | Cols: {cols}")
    lines.append(f"Column headers: {has_col_header} | Row headers: {has_row_header}")

    col_headers_obj = table.get("col_headers")
    col_headers: list[str] = (
        [str(v) for v in col_headers_obj] if _is_object_list(col_headers_obj) else []
    )
    if has_col_header and col_headers:
        lines.append("")
        lines.append("Column headers:")
        lines.append(" | ".join(col_headers))

    row_headers_obj = table.get("row_headers")
    row_headers: list[str] = (
        [str(v) for v in row_headers_obj] if _is_object_list(row_headers_obj) else []
    )
    if has_row_header and row_headers:
        lines.append("")
        lines.append("Row headers:")
        lines.append(", ".join(row_headers[:5]))
        if len(row_headers) > 5:
            lines.append(f"... and {len(row_headers) - 5} more")

    cells_obj = table.get("cells")
    cells: list[object] = cells_obj if _is_object_list(cells_obj) else []

    if cells:
        lines.append("")
        lines.append("Sample rows:")
        for row_obj in cells[:3]:
            if not _is_object_list(row_obj):
                continue

            preview_row: list[str] = []
            for cell in row_obj[:5]:
                cell_text = str(cell)
                if len(cell_text) > 24:
                    cell_text = cell_text[:21] + "..."
                preview_row.append(cell_text)

            lines.append(" | ".join(preview_row))

        if len(cells) > 3:
            lines.append(f"... and {len(cells) - 3} more rows")

    lines.append("")
    lines.append("=" * 60)
    return "\n".join(lines)
