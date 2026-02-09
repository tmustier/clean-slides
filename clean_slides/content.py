"""
Cell content normalization helpers.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from typing_extensions import TypeGuard

BULLET_RE = re.compile(r"^(?P<indent>\s*)([-•])\s+(?P<text>.+)")


def _is_dict(value: object) -> TypeGuard[dict[Any, Any]]:
    return isinstance(value, dict)


def _is_list(value: object) -> TypeGuard[list[Any]]:
    return isinstance(value, list)


@dataclass
class Paragraph:
    text: str
    lvl: int | None = None
    font: str | None = None
    size_pt: int | None = None
    color: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None


def normalize_cell(
    value: Any,
    default: Paragraph,
    parse_bullets: bool = True,
) -> list[Paragraph]:
    """Normalize a cell value into a list of Paragraphs."""
    if value is None:
        return []

    if isinstance(value, Paragraph):
        return [_merge_paragraph(value, default)]

    if _is_dict(value):
        if "icon" in value:
            return []  # icon cells render as shapes, not text
        data: dict[str, Any] = {str(k): v for k, v in value.items()}
        return [_merge_paragraph(_paragraph_from_dict(data), default)]

    if _is_list(value):
        paragraphs: list[Paragraph] = []
        for item in value:
            paragraphs.extend(normalize_cell(item, default, parse_bullets=parse_bullets))
        return paragraphs

    text = str(value)
    if parse_bullets:
        return _parse_text_lines(text, default)

    return [_merge_paragraph(Paragraph(text=text), default)]


def _parse_text_lines(text: str, default: Paragraph) -> list[Paragraph]:
    lines = text.splitlines() or [""]
    paragraphs: list[Paragraph] = []
    for line in lines:
        match = BULLET_RE.match(line)
        if match:
            indent = match.group("indent")
            lvl = 1 if len(indent) < 2 else 2
            paragraphs.append(
                _merge_paragraph(
                    Paragraph(text=match.group("text"), lvl=lvl),
                    default,
                )
            )
        else:
            paragraphs.append(_merge_paragraph(Paragraph(text=line, lvl=None), default))
    return paragraphs


def _paragraph_from_dict(data: dict[str, Any]) -> Paragraph:
    lvl_raw = data.get("lvl", data.get("level"))
    lvl = int(lvl_raw) if lvl_raw is not None else None

    raw_text = data.get("text")
    text = "" if raw_text is None else str(raw_text)

    raw_font = data.get("font")
    font = str(raw_font) if raw_font is not None else None

    raw_size = data.get("size", data.get("size_pt"))
    size_pt = int(raw_size) if raw_size is not None else None

    raw_color = data.get("color")
    color = str(raw_color) if raw_color is not None else None

    return Paragraph(
        text=text,
        lvl=lvl,
        font=font,
        size_pt=size_pt,
        color=color,
        bold=_optional_bool(data.get("bold")),
        italic=_optional_bool(data.get("italic")),
        underline=_optional_bool(data.get("underline")),
    )


def _merge_paragraph(paragraph: Paragraph, default: Paragraph) -> Paragraph:
    lvl = paragraph.lvl if paragraph.lvl is not None else default.lvl
    return Paragraph(
        text=paragraph.text,
        lvl=lvl,
        font=paragraph.font if paragraph.font is not None else default.font,
        size_pt=paragraph.size_pt if paragraph.size_pt is not None else default.size_pt,
        color=paragraph.color if paragraph.color is not None else default.color,
        bold=paragraph.bold if paragraph.bold is not None else default.bold,
        italic=paragraph.italic if paragraph.italic is not None else default.italic,
        underline=paragraph.underline if paragraph.underline is not None else default.underline,
    )


def _optional_bool(value: Any) -> bool | None:
    if value is None:
        return None
    return bool(value)
