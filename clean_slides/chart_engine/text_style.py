"""Normalization helpers for text/chart label placement enums."""

from __future__ import annotations

from pptx.enum.chart import XL_LABEL_POSITION
from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN

LABEL_POSITION_MAP = {
    "center": XL_LABEL_POSITION.CENTER,
    "ctr": XL_LABEL_POSITION.CENTER,
    "inside_end": XL_LABEL_POSITION.INSIDE_END,
    "inside_base": XL_LABEL_POSITION.INSIDE_BASE,
    "outside_end": XL_LABEL_POSITION.OUTSIDE_END,
}

ALIGNMENT_MAP = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "ctr": PP_ALIGN.CENTER,
}

VERTICAL_ANCHOR_MAP = {
    "top": MSO_VERTICAL_ANCHOR.TOP,
    "center": MSO_VERTICAL_ANCHOR.MIDDLE,
    "middle": MSO_VERTICAL_ANCHOR.MIDDLE,
    "ctr": MSO_VERTICAL_ANCHOR.MIDDLE,
    "bottom": MSO_VERTICAL_ANCHOR.BOTTOM,
    "b": MSO_VERTICAL_ANCHOR.BOTTOM,
    "t": MSO_VERTICAL_ANCHOR.TOP,
}


def normalize_label_position(value: str | None) -> XL_LABEL_POSITION | None:
    """Resolve user-facing data label position aliases."""
    if not value:
        return None
    key = value.strip().lower()
    return LABEL_POSITION_MAP.get(key)


def normalize_alignment(value: str | None) -> PP_ALIGN | None:
    """Resolve text alignment aliases."""
    if not value:
        return None
    key = value.strip().lower()
    return ALIGNMENT_MAP.get(key)


def normalize_vertical_anchor(value: str | None) -> MSO_VERTICAL_ANCHOR | None:
    """Resolve text vertical-anchor aliases."""
    if not value:
        return None
    key = value.strip().lower()
    return VERTICAL_ANCHOR_MAP.get(key)
