"""
Placeholder content generation.
"""

from __future__ import annotations

from dataclasses import replace
from typing import Any

from .spec import ChartRef, TableSpec

TITLE_WORDS: list[str] = [
    "Placeholder",
    "Title",
    "Goes",
    "Here",
    "Now",
    "Example",
    "Topic",
    "Summary",
]

BODY_SENTENCES: list[str] = [
    "This is placeholder text describing the point.",
    "Add more detail here when final content is ready.",
    "Use this space for supporting context and evidence.",
]


def _cycle_words(words: list[str], count: int) -> list[str]:
    if count <= 0:
        return []
    result: list[str] = []
    idx = 0
    while len(result) < count:
        result.append(words[idx % len(words)])
        idx += 1
    return result


def placeholder_title(words: int = 4) -> str:
    """Return a short placeholder title (3-5 words)."""
    words = max(3, min(words, 5))
    return " ".join(_cycle_words(TITLE_WORDS, words))


def placeholder_body(sentences: int = 2) -> str:
    """Return a placeholder body (default two sentences)."""
    sentences = max(1, sentences)
    return " ".join(_cycle_words(BODY_SENTENCES, sentences))


def fill_placeholders(spec: TableSpec) -> TableSpec:
    """Populate missing headers/cells with placeholder text.

    For flat tables, ``row_headers`` has one entry per body row.
    For grouped tables, ``row_headers`` has one entry per *group*
    (superheader), not per body row — so we pad against the group
    count instead.
    """
    col_headers: list[str] = list(spec.col_headers or [])
    row_headers: list[str] = list(spec.row_headers or [])
    cells: list[list[Any]] = list(spec.cells or [])

    if spec.has_col_header:
        while len(col_headers) < spec.num_cols:
            col_headers.append(placeholder_title())

    if spec.has_row_header:
        # Grouped tables: one row_header per group (not per body row).
        expected_row_headers = (
            len(spec.groups) if spec.is_grouped and spec.groups else spec.num_rows
        )
        while len(row_headers) < expected_row_headers:
            row_headers.append(placeholder_title())

    promoted_group_rows: set[int] = set()
    if spec.is_grouped and spec.groups:
        sub_row = 0
        for group in spec.groups:
            if group.promoted:
                promoted_group_rows.add(sub_row)
            sub_row += group.num_rows

    filled_cells: list[list[Any]] = []
    for r in range(spec.num_rows):
        row: list[Any] = cells[r] if r < len(cells) else []
        new_row: list[Any] = []
        for c in range(spec.num_cols):
            value = row[c] if c < len(row) else ""
            # Chart references are not text — leave them as-is
            if isinstance(value, ChartRef):
                new_row.append(value)
            # Promoted singleton-group first cell is intentionally blank;
            # keep it blank instead of injecting placeholder text.
            elif c == 0 and r in promoted_group_rows and value == "":
                new_row.append("")
            elif not value:
                new_row.append(placeholder_body())
            else:
                new_row.append(value)
        filled_cells.append(new_row)

    return replace(
        spec,
        col_headers=col_headers if spec.has_col_header else spec.col_headers,
        row_headers=row_headers if spec.has_row_header else spec.row_headers,
        cells=filled_cells,
    )
