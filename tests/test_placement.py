"""Tests for placement-based spatial verification."""

from __future__ import annotations

from clean_slides.placement import (
    Placement,
    check_all,
    check_bounds,
    check_group_alignment,
    check_overlaps,
    check_uniform_spacing,
)
from clean_slides.spec import ContentArea

# EMU helpers
_IN = 914400  # 1 inch in EMU


def _box(x_in: float, y_in: float, w_in: float, h_in: float) -> tuple[int, int, int, int]:
    """Build a Box from inches."""
    return (int(x_in * _IN), int(y_in * _IN), int(w_in * _IN), int(h_in * _IN))


def _area(x_in: float, y_in: float, w_in: float, h_in: float) -> ContentArea:
    return ContentArea(
        x=int(x_in * _IN),
        y=int(y_in * _IN),
        width=int(w_in * _IN),
        height=int(h_in * _IN),
    )


# ── check_overlaps ──────────────────────────────────────────────────


class TestCheckOverlaps:
    def test_no_overlaps(self) -> None:
        placements = [
            Placement(name="a", role="bar", box=_box(0, 0, 1, 1)),
            Placement(name="b", role="bar", box=_box(2, 0, 1, 1)),
        ]
        assert check_overlaps(placements) == []

    def test_overlap_detected(self) -> None:
        placements = [
            Placement(name="a", role="bar", box=_box(0, 0, 2, 1)),
            Placement(name="b", role="bar", box=_box(1, 0, 2, 1)),
        ]
        issues = check_overlaps(placements)
        assert len(issues) == 1
        assert issues[0].severity == "error"
        assert issues[0].category == "overlap"
        assert "a" in issues[0].message
        assert "b" in issues[0].message

    def test_adjacent_no_overlap(self) -> None:
        """Touching edges are not overlaps."""
        placements = [
            Placement(name="a", role="bar", box=_box(0, 0, 1, 1)),
            Placement(name="b", role="bar", box=_box(1, 0, 1, 1)),
        ]
        assert check_overlaps(placements) == []

    def test_dividers_skipped(self) -> None:
        """Dividers never participate in overlap checks."""
        placements = [
            Placement(name="a", role="bar", box=_box(0, 0, 2, 1)),
            Placement(name="d", role="divider", box=_box(0, 0, 3, 0.01)),
        ]
        assert check_overlaps(placements) == []

    def test_background_skipped(self) -> None:
        placements = [
            Placement(name="a", role="bar", box=_box(0, 0, 2, 1)),
            Placement(name="bg", role="background", box=_box(0, 0, 10, 10)),
        ]
        assert check_overlaps(placements) == []

    def test_connectors_skipped(self) -> None:
        placements = [
            Placement(name="a", role="bar", box=_box(0, 0, 2, 1)),
            Placement(name="c", role="connector", box=_box(0, 0, 2, 1)),
        ]
        assert check_overlaps(placements) == []


# ── check_bounds ─────────────────────────────────────────────────────


class TestCheckBounds:
    def test_within_bounds(self) -> None:
        area = _area(1, 1, 8, 5)
        placements = [
            Placement(name="a", role="bar", box=_box(1, 1, 2, 2)),
        ]
        assert check_bounds(placements, area) == []

    def test_exceeds_right(self) -> None:
        area = _area(1, 1, 3, 5)
        placements = [
            Placement(name="a", role="bar", box=_box(1, 1, 4, 2)),
        ]
        issues = check_bounds(placements, area)
        assert len(issues) == 1
        assert issues[0].severity == "error"
        assert issues[0].category == "boundary"
        assert "right" in issues[0].message

    def test_exceeds_left(self) -> None:
        area = _area(2, 1, 3, 5)
        placements = [
            Placement(name="a", role="bar", box=_box(1, 1, 1, 1)),
        ]
        issues = check_bounds(placements, area)
        assert len(issues) == 1
        assert "left" in issues[0].message

    def test_dividers_skipped(self) -> None:
        area = _area(1, 1, 3, 5)
        placements = [
            Placement(name="d", role="divider", box=_box(0, 0, 20, 0.01)),
        ]
        assert check_bounds(placements, area) == []


# ── check_group_alignment ────────────────────────────────────────────


class TestCheckGroupAlignment:
    def test_aligned(self) -> None:
        """Label centered over bar → no issues."""
        placements = [
            Placement(name="bar_Rev", role="bar", group="Revenue", box=_box(2, 2, 1, 3)),
            Placement(name="label_Rev", role="label", group="Revenue", box=_box(2, 1.5, 1, 0.3)),
        ]
        assert check_group_alignment(placements) == []

    def test_misaligned(self) -> None:
        """Label offset from bar center → warning."""
        placements = [
            Placement(name="bar_Rev", role="bar", group="Revenue", box=_box(2, 2, 1, 3)),
            Placement(name="label_Rev", role="label", group="Revenue", box=_box(3, 1.5, 1, 0.3)),
        ]
        issues = check_group_alignment(placements)
        assert len(issues) == 1
        assert issues[0].severity == "warning"
        assert issues[0].category == "alignment"

    def test_no_group_ignored(self) -> None:
        """Shapes without group are not checked."""
        placements = [
            Placement(name="bar_Rev", role="bar", box=_box(2, 2, 1, 3)),
            Placement(name="label_Rev", role="label", box=_box(5, 1.5, 1, 0.3)),
        ]
        assert check_group_alignment(placements) == []

    def test_different_groups_independent(self) -> None:
        """Misalignment within one group, aligned in another."""
        placements = [
            # Revenue: aligned
            Placement(name="bar_Rev", role="bar", group="Revenue", box=_box(2, 2, 1, 3)),
            Placement(name="label_Rev", role="label", group="Revenue", box=_box(2, 1.5, 1, 0.3)),
            # Cost: misaligned
            Placement(name="bar_Cost", role="bar", group="Cost", box=_box(4, 2, 1, 3)),
            Placement(name="label_Cost", role="label", group="Cost", box=_box(6, 1.5, 1, 0.3)),
        ]
        issues = check_group_alignment(placements)
        assert len(issues) == 1
        assert issues[0].details.get("group") == "Cost"

    def test_custom_roles(self) -> None:
        """Can check alignment between custom role pairs."""
        placements = [
            Placement(name="dot_A", role="dot", group="A", box=_box(2, 2, 0.2, 0.2)),
            Placement(name="caption_A", role="caption", group="A", box=_box(2, 2.5, 0.2, 0.3)),
        ]
        assert check_group_alignment(placements, label_role="caption", anchor_role="dot") == []


# ── check_uniform_spacing ────────────────────────────────────────────


class TestCheckUniformSpacing:
    def test_uniform(self) -> None:
        """Three evenly-spaced bars → no issues."""
        placements = [
            Placement(name="bar_A", role="bar", box=_box(1, 1, 1, 3)),
            Placement(name="bar_B", role="bar", box=_box(2.5, 1, 1, 3)),
            Placement(name="bar_C", role="bar", box=_box(4, 1, 1, 3)),
        ]
        assert check_uniform_spacing(placements, "bar") == []

    def test_uneven(self) -> None:
        """Gaps differ beyond tolerance → warning."""
        placements = [
            Placement(name="bar_A", role="bar", box=_box(1, 1, 1, 3)),
            Placement(name="bar_B", role="bar", box=_box(2.1, 1, 1, 3)),  # gap 0.1
            Placement(name="bar_C", role="bar", box=_box(4, 1, 1, 3)),  # gap 0.9
        ]
        issues = check_uniform_spacing(placements, "bar")
        assert len(issues) == 1
        assert issues[0].severity == "warning"
        assert issues[0].category == "spacing"

    def test_two_items_skipped(self) -> None:
        """Need at least 3 items to compare gaps."""
        placements = [
            Placement(name="bar_A", role="bar", box=_box(1, 1, 1, 3)),
            Placement(name="bar_B", role="bar", box=_box(3, 1, 1, 3)),
        ]
        assert check_uniform_spacing(placements, "bar") == []

    def test_vertical_axis(self) -> None:
        """Uniform vertical spacing → no issues."""
        placements = [
            Placement(name="row_A", role="row", box=_box(1, 1, 5, 1)),
            Placement(name="row_B", role="row", box=_box(1, 2.5, 5, 1)),
            Placement(name="row_C", role="row", box=_box(1, 4, 5, 1)),
        ]
        assert check_uniform_spacing(placements, "row", axis="vertical") == []

    def test_filters_by_role(self) -> None:
        """Only checks the specified role."""
        placements = [
            Placement(name="bar_A", role="bar", box=_box(1, 1, 1, 3)),
            Placement(name="bar_B", role="bar", box=_box(2.5, 1, 1, 3)),
            Placement(name="bar_C", role="bar", box=_box(4, 1, 1, 3)),
            Placement(name="label_X", role="label", box=_box(0, 0, 0.5, 0.3)),
        ]
        assert check_uniform_spacing(placements, "bar") == []


# ── check_all ────────────────────────────────────────────────────────


class TestCheckAll:
    def test_clean_layout(self) -> None:
        area = _area(0, 0, 10, 8)
        placements = [
            Placement(name="bar_A", role="bar", group="A", box=_box(1, 2, 1, 3)),
            Placement(name="label_A", role="label", group="A", box=_box(1, 1.5, 1, 0.3)),
            Placement(name="bar_B", role="bar", group="B", box=_box(3, 2, 1, 3)),
            Placement(name="label_B", role="label", group="B", box=_box(3, 1.5, 1, 0.3)),
            Placement(name="bar_C", role="bar", group="C", box=_box(5, 2, 1, 3)),
            Placement(name="label_C", role="label", group="C", box=_box(5, 1.5, 1, 0.3)),
        ]
        issues = check_all(placements, area, spacing_roles=["bar"])
        assert issues == []

    def test_multiple_issues(self) -> None:
        area = _area(2, 2, 4, 4)
        placements = [
            # Overlap: a and b share space
            Placement(name="a", role="bar", box=_box(2, 2, 2, 2)),
            Placement(name="b", role="bar", box=_box(3, 2, 2, 2)),
            # Out of bounds
            Placement(name="c", role="bar", box=_box(0, 0, 1, 1)),
        ]
        issues = check_all(placements, area)
        categories = {i.category for i in issues}
        assert "overlap" in categories
        assert "boundary" in categories
