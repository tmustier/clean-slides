"""Spec normalization helpers shared by chart payload builders."""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
from typing import cast


def str_key_dict(value: object) -> dict[str, object]:
    """Return a ``dict[str, object]`` view of a mapping-like payload."""
    if not isinstance(value, dict):
        return {}

    result: dict[str, object] = {}
    typed_value = cast(dict[object, object], value)
    for key, item in typed_value.items():
        if isinstance(key, str):
            result[key] = item
    return result


def object_list(value: object) -> list[object]:
    """Coerce list/tuple payloads into ``list[object]``."""
    if isinstance(value, list):
        return cast(list[object], value)
    if isinstance(value, tuple):
        return list(cast(tuple[object, ...], value))
    return []


def optional_str_list(value: object) -> list[str | None]:
    """Coerce unknown list payload to ``list[str | None]``."""
    result: list[str | None] = []
    for item in object_list(value):
        if item is None or isinstance(item, str):
            result.append(item)
    return result


def normalize_list(value: object) -> list[object]:
    """Normalize scalar/None values into list form."""
    if value is None:
        return []
    if isinstance(value, list):
        return cast(list[object], value)
    return [value]


def safe_value(values: Sequence[object], idx: int) -> object | None:
    """Safely access a sequence value by index with blank-string handling."""
    if idx < 0 or idx >= len(values):
        return None
    value = values[idx]
    if isinstance(value, str) and not value.strip():
        return None
    return value


def numeric_value(value: object) -> float | None:
    """Parse numeric-like scalar to float."""
    if isinstance(value, (int, float)):
        return float(value)
    return None


def sum_numeric(values: Iterable[object]) -> float:
    """Sum values, skipping non-numeric entries."""
    total = 0.0
    for value in values:
        number = numeric_value(value)
        if number is not None:
            total += number
    return total


def normalize_category_set(categories: list[object], raw: object) -> set[int]:
    """Resolve category selectors to a unique set of category indices."""
    indices: set[int] = set()
    for item in normalize_list(raw):
        if isinstance(item, int):
            indices.add(item)
        elif isinstance(item, str) and item in categories:
            indices.add(categories.index(item))
    return indices


def normalize_category_indices(categories: list[object], raw: object) -> list[int]:
    """Resolve category selectors to ordered unique category indices."""
    indices: list[int] = []
    for item in normalize_list(raw):
        if isinstance(item, int):
            indices.append(item)
        elif isinstance(item, str) and item in categories:
            indices.append(categories.index(item))
    return list(dict.fromkeys(indices))


def normalize_total_series(raw: object) -> set[str]:
    """Normalize series-name selectors for total-series lookup."""
    names: set[str] = set()
    for item in normalize_list(raw):
        if isinstance(item, str):
            names.add(item)
    return names


def is_total_series(entry: Mapping[str, object], total_series_names: set[str]) -> bool:
    """Return whether a series entry should be treated as totals."""
    if entry.get("role") == "total":
        return True
    name = str(entry.get("name", ""))
    if total_series_names and name in total_series_names:
        return True
    return name.lower() in {"total", "totals"}


def normalize_series_set(entries: list[Mapping[str, object]], raw: object) -> set[str]:
    """Resolve series selectors to names (supports indices and names)."""
    names = [str(entry.get("name", "Series")) for entry in entries]
    resolved: set[str] = set()
    for item in normalize_list(raw):
        if isinstance(item, int):
            if 0 <= item < len(names):
                resolved.add(names[item])
        elif isinstance(item, str) and item in names:
            resolved.add(item)
    return resolved


def is_range_series(entry: Mapping[str, object], range_series_names: set[str]) -> bool:
    """Return whether series contributes only to visual range/band, not totals."""
    role = str(entry.get("role", "")).lower()
    if role in {"range", "band"}:
        return True
    name = str(entry.get("name", ""))
    return bool(range_series_names and name in range_series_names)


def format_label(value: float | None, decimals: int = 0) -> str | None:
    """Format numeric labels with optional fixed decimal places."""
    if value is None:
        return None
    if decimals <= 0:
        return str(round(value))
    return f"{value:.{decimals}f}"


def infer_label_decimals(values: Iterable[float | None], default: int = 0) -> int:
    """Infer one decimal place when non-integer values are present."""
    for value in values:
        if value is None:
            continue
        if abs(value - round(value)) > 1e-6:
            return 1
    return default


def infer_total_categories(
    categories: list[object],
    segment_entries: list[Mapping[str, object]],
    total_entry: Mapping[str, object] | None,
) -> set[int]:
    """Infer total categories where total-series value exists without segment values."""
    if total_entry is None:
        return set()

    indices: set[int] = set()
    total_values_obj = total_entry.get("values")
    total_values = (
        cast(list[object], total_values_obj) if isinstance(total_values_obj, list) else []
    )

    for idx in range(len(categories)):
        total_value = numeric_value(safe_value(total_values, idx))
        if total_value is None:
            continue

        segment_has_value = False
        for entry in segment_entries:
            values_obj = entry.get("values")
            values = cast(list[object], values_obj) if isinstance(values_obj, list) else []
            if numeric_value(safe_value(values, idx)) is not None:
                segment_has_value = True
                break

        if not segment_has_value:
            indices.add(idx)

    return indices
