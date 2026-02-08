"""
Placeholder content generation.
"""

from dataclasses import replace
from typing import Any, List

from .spec import TableSpec

TITLE_WORDS: List[str] = [
    "Placeholder",
    "Title",
    "Goes",
    "Here",
    "Now",
    "Example",
    "Topic",
    "Summary",
]

BODY_SENTENCES: List[str] = [
    "This is placeholder text describing the point.",
    "Add more detail here when final content is ready.",
    "Use this space for supporting context and evidence.",
]


def _cycle_words(words: List[str], count: int) -> List[str]:
    if count <= 0:
        return []
    result: List[str] = []
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
    """Populate missing headers/cells with placeholder text."""
    col_headers: List[str] = list(spec.col_headers or [])
    row_headers: List[str] = list(spec.row_headers or [])
    cells: List[List[Any]] = list(spec.cells or [])

    if spec.has_col_header:
        while len(col_headers) < spec.num_cols:
            col_headers.append(placeholder_title())

    if spec.has_row_header:
        while len(row_headers) < spec.num_rows:
            row_headers.append(placeholder_title())

    filled_cells: List[List[Any]] = []
    for r in range(spec.num_rows):
        row: List[Any] = cells[r] if r < len(cells) else []
        new_row: List[Any] = []
        for c in range(spec.num_cols):
            value = row[c] if c < len(row) else ""
            if not value:
                value = placeholder_body()
            new_row.append(value)
        filled_cells.append(new_row)

    return replace(
        spec,
        col_headers=col_headers if spec.has_col_header else spec.col_headers,
        row_headers=row_headers if spec.has_row_header else spec.row_headers,
        cells=filled_cells,
    )
