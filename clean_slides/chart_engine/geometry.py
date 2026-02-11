"""Geometry helpers for chart overlays and manual data-label layout."""

from __future__ import annotations

ORIENTATION_MAP = {
    "vertical": "vertical",
    "v": "vertical",
    "column": "vertical",
    "col": "vertical",
    "columns": "vertical",
    "horizontal": "horizontal",
    "h": "horizontal",
    "bar": "horizontal",
    "row": "horizontal",
    "rows": "horizontal",
}


def normalize_orientation(value: str | None) -> str:
    """Normalize chart orientation aliases to ``vertical``/``horizontal``."""
    if not value:
        return "vertical"
    key = str(value).strip().lower()
    return ORIENTATION_MAP.get(key, "vertical")


def value_to_x(
    value: float,
    axis_min: float,
    axis_max: float,
    plot_left: float,
    plot_width: float,
) -> float:
    """Map a chart value to X-coordinate in plot space."""
    if axis_max == axis_min:
        return plot_left + plot_width / 2
    return plot_left + (value - axis_min) / (axis_max - axis_min) * plot_width


def value_to_y(
    value: float, axis_min: float, axis_max: float, plot_top: float, plot_height: float
) -> float:
    """Map a chart value to Y-coordinate in plot space."""
    if axis_max == axis_min:
        return plot_top + plot_height / 2
    return plot_top + (axis_max - value) / (axis_max - axis_min) * plot_height


def compute_category_geometry(
    chart_box: tuple[int, int, int, int],
    plot_layout: dict[str, float],
    categories: list[object],
    gap_width: int,
    orientation: str = "vertical",
) -> dict[str, object]:
    """Compute plot bounds and per-category bar geometry."""
    x, y, cx, cy = chart_box
    plot_left = x + plot_layout.get("x", 0) * cx
    plot_top = y + plot_layout.get("y", 0) * cy
    plot_width = plot_layout.get("w", 1) * cx
    plot_height = plot_layout.get("h", 1) * cy

    count = max(1, len(categories))
    gap_ratio = gap_width / 100.0
    orientation_value = normalize_orientation(orientation)

    if orientation_value == "horizontal":
        slot_height = plot_height / count
        bar_height = slot_height / (1 + gap_ratio)
        gap = slot_height - bar_height

        bar_tops = [plot_top + slot_height * idx + gap / 2 for idx in range(count)]
        bar_bottoms = [top + bar_height for top in bar_tops]
        bar_centers = [top + bar_height / 2 for top in bar_tops]

        return {
            "plot_left": plot_left,
            "plot_top": plot_top,
            "plot_width": plot_width,
            "plot_height": plot_height,
            "slot_height": slot_height,
            "bar_height": bar_height,
            "gap": gap,
            "bar_tops": bar_tops,
            "bar_bottoms": bar_bottoms,
            "bar_centers": bar_centers,
            "orientation": orientation_value,
        }

    slot_width = plot_width / count
    bar_width = slot_width / (1 + gap_ratio)
    gap = slot_width - bar_width

    bar_lefts = [plot_left + slot_width * idx + gap / 2 for idx in range(count)]
    bar_rights = [left + bar_width for left in bar_lefts]
    bar_centers = [left + bar_width / 2 for left in bar_lefts]

    return {
        "plot_left": plot_left,
        "plot_top": plot_top,
        "plot_width": plot_width,
        "plot_height": plot_height,
        "slot_width": slot_width,
        "bar_width": bar_width,
        "gap": gap,
        "bar_lefts": bar_lefts,
        "bar_rights": bar_rights,
        "bar_centers": bar_centers,
        "orientation": orientation_value,
    }


def resolve_label_collisions(
    labels: list[dict[str, float | int | str]],
    axis: str = "y",
    min_gap: int = 0,
    direction: int = 1,
) -> None:
    """Resolve label overlap in-place along one axis."""
    if axis not in {"x", "y"}:
        return
    key = "y" if axis == "y" else "x"
    size_key = "height" if axis == "y" else "width"

    if direction < 0:
        ordered = sorted(labels, key=lambda item: float(item.get(key, 0)), reverse=True)
        for idx in range(1, len(ordered)):
            prev = ordered[idx - 1]
            curr = ordered[idx]
            prev_start = float(prev.get(key, 0))
            curr_end = float(curr.get(key, 0)) + float(curr.get(size_key, 0))
            if curr_end + min_gap > prev_start:
                curr[key] = float(curr.get(key, 0)) - (curr_end + min_gap - prev_start)
    else:
        ordered = sorted(labels, key=lambda item: float(item.get(key, 0)))
        for idx in range(1, len(ordered)):
            prev = ordered[idx - 1]
            curr = ordered[idx]
            prev_end = float(prev.get(key, 0)) + float(prev.get(size_key, 0))
            curr_start = float(curr.get(key, 0))
            if curr_start < prev_end + min_gap:
                curr[key] = curr_start + (prev_end + min_gap - curr_start)
