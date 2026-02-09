"""
Placement-based spatial verification.

A Placement is a shape the renderer placed, annotated with its role and
group.  Standalone check functions operate on ``list[Placement]`` and
return ``list[LayoutIssue]`` — no classes, no inheritance.

For the design rationale, see docs/VISUAL-VERIFICATION-RESEARCH.md §4.
"""

from __future__ import annotations

from dataclasses import dataclass

from .spec import Box, ContentArea

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------


@dataclass
class Placement:
    """A shape placed by the renderer, with its intended role.

    Attributes:
        name:     Shape name in the PPTX (e.g. "bar_Revenue").
        role:     Semantic role — "bar", "label", "connector", "cell",
                  "header", "icon", "divider", "background", etc.
        box:      (x, y, w, h) in EMU.
        group:    Shared key linking related elements.  A label and its
                  bar both get group="Revenue" so alignment checks can
                  pair them without heuristic proximity matching.
        text:     Text content (for text-fit checks).
        font:     Font name.
        size_pt:  Font size in points.
    """

    name: str
    role: str
    box: Box
    group: str = ""
    text: str = ""
    font: str = ""
    size_pt: int = 0


# ---------------------------------------------------------------------------
# Geometry helpers
# ---------------------------------------------------------------------------


def _boxes_overlap(a: Box, b: Box) -> bool:
    """True if two EMU bounding boxes intersect."""
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return min(ax + aw, bx + bw) > max(ax, bx) and min(ay + ah, by + bh) > max(ay, by)


def _box_center_x(box: Box) -> int:
    """Horizontal center of a box in EMU."""
    return box[0] + box[2] // 2


def _emu_to_in(emu: int) -> float:
    """EMU → inches, rounded to 2 decimal places."""
    return round(emu / 914400, 2)


def _box_desc(p: Placement) -> str:
    """Short human-readable box description in inches."""
    x, y, w, h = p.box
    return f"{p.name} ({_emu_to_in(x)}, {_emu_to_in(y)}) " f"{_emu_to_in(w)}×{_emu_to_in(h)}in"


# ---------------------------------------------------------------------------
# Issue dataclass (re-export from verification to keep a single type)
# ---------------------------------------------------------------------------

# Import here to avoid circular deps — placement.py is a leaf module
# that only depends on spec.py.  verification.py imports from us, not
# the other way round.  We define a lightweight issue type that is
# structurally identical so callers can mix them freely.


@dataclass
class PlacementIssue:
    """A spatial issue found by a placement check.

    Fields mirror ``LayoutIssue`` from verification.py so they can be
    merged into the same report.
    """

    severity: str  # "error" | "warning" | "info"
    category: str  # "overlap" | "boundary" | "alignment" | "spacing"
    message: str
    details: dict[str, object]


# ---------------------------------------------------------------------------
# Check functions
# ---------------------------------------------------------------------------

# Roles that are never flagged for overlap (decorative / structural).
_OVERLAP_SKIP_ROLES = frozenset({"background", "divider", "connector"})


def check_overlaps(placements: list[Placement]) -> list[PlacementIssue]:
    """Flag any two content shapes whose bounding boxes intersect."""
    issues: list[PlacementIssue] = []
    content = [p for p in placements if p.role not in _OVERLAP_SKIP_ROLES]
    for i, a in enumerate(content):
        for b in content[i + 1 :]:
            if _boxes_overlap(a.box, b.box):
                issues.append(
                    PlacementIssue(
                        severity="error",
                        category="overlap",
                        message=f'"{a.name}" overlaps "{b.name}"',
                        details={
                            "shape_a": _box_desc(a),
                            "shape_b": _box_desc(b),
                        },
                    )
                )
    return issues


def check_bounds(
    placements: list[Placement],
    area: ContentArea,
) -> list[PlacementIssue]:
    """Flag shapes that exceed the content area."""
    issues: list[PlacementIssue] = []
    for p in placements:
        if p.role in ("divider",):
            continue
        x, y, w, h = p.box
        overflow = {
            "left": max(area.x - x, 0),
            "top": max(area.y - y, 0),
            "right": max((x + w) - (area.x + area.width), 0),
            "bottom": max((y + h) - (area.y + area.height), 0),
        }
        if any(overflow.values()):
            over_desc = ", ".join(
                f"{side} {_emu_to_in(v)}in" for side, v in overflow.items() if v > 0
            )
            issues.append(
                PlacementIssue(
                    severity="error",
                    category="boundary",
                    message=f'"{p.name}" exceeds content area ({over_desc})',
                    details={"shape": _box_desc(p), "overflow": overflow},
                )
            )
    return issues


def check_group_alignment(
    placements: list[Placement],
    label_role: str = "label",
    anchor_role: str = "bar",
    tolerance_emu: int = 45720,  # 0.05 inches
) -> list[PlacementIssue]:
    """Within each group, check that labels are horizontally centered over anchors.

    Only examines placements that have a non-empty ``group`` and matching
    ``label_role`` / ``anchor_role``.
    """
    issues: list[PlacementIssue] = []

    # Build {group: {role: [placements]}}
    groups: dict[str, dict[str, list[Placement]]] = {}
    for p in placements:
        if not p.group:
            continue
        groups.setdefault(p.group, {}).setdefault(p.role, []).append(p)

    for group_key, roles in groups.items():
        labels = roles.get(label_role, [])
        anchors = roles.get(anchor_role, [])
        if not labels or not anchors:
            continue

        for label in labels:
            for anchor in anchors:
                label_cx = _box_center_x(label.box)
                anchor_cx = _box_center_x(anchor.box)
                delta = abs(label_cx - anchor_cx)
                if delta > tolerance_emu:
                    issues.append(
                        PlacementIssue(
                            severity="warning",
                            category="alignment",
                            message=(
                                f'"{label.name}" is not centered over '
                                f'"{anchor.name}" (off by {_emu_to_in(delta)}in)'
                            ),
                            details={
                                "group": group_key,
                                "label": _box_desc(label),
                                "anchor": _box_desc(anchor),
                                "offset_emu": delta,
                            },
                        )
                    )
    return issues


def check_uniform_spacing(
    placements: list[Placement],
    role: str,
    axis: str = "horizontal",
    tolerance_emu: int = 27432,  # 0.03 inches
) -> list[PlacementIssue]:
    """Check that shapes of a given role are evenly spaced.

    Computes pairwise gaps along the specified axis (sorted by position)
    and flags if max_gap - min_gap exceeds tolerance.
    """
    items = sorted(
        [p for p in placements if p.role == role],
        key=lambda p: p.box[0] if axis == "horizontal" else p.box[1],
    )
    if len(items) < 3:
        # Need at least 3 items to compare gaps meaningfully
        return []

    gaps: list[int] = []
    gap_labels: list[str] = []
    for i in range(len(items) - 1):
        a, b = items[i], items[i + 1]
        if axis == "horizontal":
            gap = b.box[0] - (a.box[0] + a.box[2])
        else:
            gap = b.box[1] - (a.box[1] + a.box[3])
        gaps.append(gap)
        gap_labels.append(f"{a.name}→{b.name}")

    if not gaps:
        return []

    min_gap = min(gaps)
    max_gap = max(gaps)
    spread = max_gap - min_gap

    if spread <= tolerance_emu:
        return []

    gap_detail = ", ".join(f"{label}: {_emu_to_in(g)}in" for label, g in zip(gap_labels, gaps))
    return [
        PlacementIssue(
            severity="warning",
            category="spacing",
            message=f"{role}s have uneven {axis} gaps (spread {_emu_to_in(spread)}in)",
            details={
                "role": role,
                "axis": axis,
                "gaps": gap_detail,
                "min_gap_in": _emu_to_in(min_gap),
                "max_gap_in": _emu_to_in(max_gap),
            },
        )
    ]


# ---------------------------------------------------------------------------
# Runner — convenience to call all checks
# ---------------------------------------------------------------------------


def check_all(
    placements: list[Placement],
    area: ContentArea | None = None,
    spacing_roles: list[str] | None = None,
) -> list[PlacementIssue]:
    """Run all standard checks and return combined issues.

    Args:
        placements:     Shapes to verify.
        area:           Content area for boundary checks (skipped if None).
        spacing_roles:  Roles to check for uniform spacing (e.g. ["bar"]).
    """
    issues: list[PlacementIssue] = []
    issues += check_overlaps(placements)
    if area is not None:
        issues += check_bounds(placements, area)
    issues += check_group_alignment(placements)
    for role in spacing_roles or []:
        issues += check_uniform_spacing(placements, role)
    return issues
