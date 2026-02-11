"""Text body template cache and application helpers."""

# pyright: reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false
# pyright: reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false
# pyright: reportMissingParameterType=false
# pyright: reportMissingTypeArgument=false
# pyright: reportGeneralTypeIssues=false
# pyright: reportArgumentType=false
# pyright: reportCallIssue=false
# pyright: reportAttributeAccessIssue=false

from __future__ import annotations

import copy
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement

_TEXT_STYLE_CACHE: dict[tuple[str, str], OxmlElement] = {}


def load_txbody_template(template_path: Path, sample_text: str) -> OxmlElement | None:
    key = (str(template_path), sample_text)
    cached = _TEXT_STYLE_CACHE.get(key)
    if cached is not None:
        return cached

    prs = Presentation(str(template_path))
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            if shape.text.strip() != sample_text:
                continue
            tx_body = shape._element.find(qn("p:txBody"))
            if tx_body is None:
                continue
            cached = copy.deepcopy(tx_body)
            _TEXT_STYLE_CACHE[key] = cached
            return cached
    return None


def apply_txbody_template(box, template: OxmlElement | None, text: str | None) -> None:
    if template is None:
        return
    tx_body = copy.deepcopy(template)
    if text is not None:
        for t_elem in tx_body.iter(qn("a:t")):
            t_elem.text = text
    existing = box._element.find(qn("p:txBody"))
    if existing is not None:
        existing.getparent().replace(existing, tx_body)
    else:
        box._element.append(tx_body)


def resolve_txbody_template(
    template_path: Path | None,
    text: str,
    fallback: OxmlElement | None,
) -> OxmlElement | None:
    if template_path is None:
        return fallback
    template = load_txbody_template(template_path, text)
    return template if template is not None else fallback
