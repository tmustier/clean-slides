"""Unit and path coercion helpers for chart rendering."""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

from pptx.util import Emu, Inches


def coerce_emu(value: Any) -> int | None:
    """Coerce common scalar inputs to EMU.

    Behavior matches historical chart-generator semantics:
    - ``Emu`` values are passed through
    - large ints/floats (>|10000|) are treated as already-EMU
    - smaller numeric values are interpreted as inches
    """
    if value is None:
        return None
    if isinstance(value, Emu):
        return int(value)
    if isinstance(value, (int, float)):
        if abs(value) > 10000:
            return int(value)
        return int(Inches(float(value)))
    return None


def emu_or_default(value: Any, default: int) -> int:
    """Return EMU coercion result or a fallback default."""
    coerced = coerce_emu(value)
    return default if coerced is None else coerced


def coerce_line_width(value: Any) -> int | None:
    """Coerce line width preserving historical narrow-value behavior."""
    if value is None:
        return None
    if isinstance(value, Emu):
        return int(value)
    if isinstance(value, (int, float)) and abs(value) <= 10000:
        return int(value)
    return coerce_emu(value)


def coerce_offset_value(value: Any) -> int:
    """Coerce offset values used in overlay adjustment matrices."""
    if value is None:
        return 0
    if isinstance(value, Emu):
        return int(value)
    if isinstance(value, (int, float)):
        return int(value)
    coerced = coerce_emu(value)
    return coerced if coerced is not None else 0


def normalize_offset_matrix(value: object) -> list[list[int]]:
    """Normalize scalar/list inputs to a matrix of integer offsets."""

    def _normalize_list(raw: object) -> list[object]:
        if raw is None:
            return []
        if isinstance(raw, list):
            return cast(list[object], raw)
        return [raw]

    matrix: list[list[int]] = []
    for row in _normalize_list(value):
        if isinstance(row, (list, tuple)):
            values = cast(list[object] | tuple[object, ...], row)
            matrix.append([coerce_offset_value(item) for item in values])
        else:
            matrix.append([coerce_offset_value(row)])
    return matrix


def resolve_path(value: str, base_dir: str | Path | None = None) -> Path:
    """Resolve a possibly-relative path against an optional base directory."""
    path = Path(value)
    if path.is_absolute():
        return path

    if isinstance(base_dir, str):
        root = Path(base_dir)
    elif isinstance(base_dir, Path):
        root = base_dir
    else:
        root = Path.cwd()

    return (root / path).resolve()
