"""Text body template cache and application helpers."""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Any, Union, cast

from pptx import Presentation
from pptx.oxml.ns import qn

TxBodyTemplate = object
PathOrNone = Union[Path, None]

_TEXT_STYLE_CACHE: dict[tuple[str, str], TxBodyTemplate] = {}


def load_txbody_template(template_path: Path, sample_text: str) -> TxBodyTemplate | None:
    key = (str(template_path), sample_text)
    cached = _TEXT_STYLE_CACHE.get(key)
    if cached is not None:
        return cached

    prs = Presentation(str(template_path))
    for slide in prs.slides:
        for shape in slide.shapes:
            if not bool(getattr(shape, "has_text_frame", False)):
                continue

            text_value = getattr(shape, "text", None)
            if not isinstance(text_value, str):
                continue
            if text_value.strip() != sample_text:
                continue

            shape_element = getattr(shape, "_element", None)
            if shape_element is None:
                continue

            tx_body = shape_element.find(qn("p:txBody"))
            if tx_body is None:
                continue

            cached = copy.deepcopy(tx_body)
            _TEXT_STYLE_CACHE[key] = cached
            return cached

    return None


def apply_txbody_template(box: Any, template: TxBodyTemplate | None, text: str | None) -> None:
    if template is None:
        return

    tx_body = copy.deepcopy(template)
    if text is not None:
        tx_body_element = cast(Any, tx_body)
        for t_elem in tx_body_element.iter(qn("a:t")):
            t_elem.text = text

    box_element = getattr(box, "_element", None)
    if box_element is None:
        return

    existing = box_element.find(qn("p:txBody"))
    if existing is not None:
        parent = existing.getparent()
        if parent is not None:
            parent.replace(existing, tx_body)
    else:
        box_element.append(tx_body)


def resolve_txbody_template(
    template_path: PathOrNone,
    text: str,
    fallback: TxBodyTemplate | None,
) -> TxBodyTemplate | None:
    if template_path is None:
        return fallback

    template = load_txbody_template(template_path, text)
    return template if template is not None else fallback
