"""Typed compatibility helpers for python-pptx dynamic objects.

python-pptx exposes many attributes via runtime descriptors without precise
static typing. These helpers centralize guarded access so callers can avoid
repeating ``getattr(...)`` probes and local casts.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Protocol, cast


def _iter_objects(value: object) -> list[object]:
    if value is None:
        return []
    if isinstance(value, list):
        return cast(list[object], value)
    if isinstance(value, tuple):
        return list(cast(tuple[object, ...], value))
    if isinstance(value, set):
        return list(cast(set[object], value))
    if isinstance(value, (str, bytes, bytearray)):
        return []
    if not isinstance(value, Iterable):
        return []

    iterable = cast(Iterable[object], value)
    return list(iterable)


def iter_slides(presentation: object) -> list[object]:
    """Return presentation slides as plain objects."""
    return _iter_objects(getattr(presentation, "slides", None))


def iter_shapes(container: object) -> list[object]:
    """Return shape collection as plain objects."""
    return _iter_objects(getattr(container, "shapes", None))


def shape_has_chart(shape: object) -> bool:
    """Return whether ``shape`` exposes a chart."""
    return bool(getattr(shape, "has_chart", False))


def shape_chart(shape: object) -> object | None:
    """Return chart object from ``shape`` when present."""
    if not shape_has_chart(shape):
        return None
    return getattr(shape, "chart", None)


def slide_charts(slide: object) -> list[object]:
    """Collect chart objects on a slide."""
    charts: list[object] = []
    for shape in iter_shapes(slide):
        chart = shape_chart(shape)
        if chart is not None:
            charts.append(chart)
    return charts


def chart_type_value(chart: object) -> int | None:
    """Return integer chart type when available."""
    chart_type = getattr(chart, "chart_type", None)
    if chart_type is None:
        return None
    try:
        return int(chart_type)
    except (TypeError, ValueError):
        return None


def shape_chart_type(shape: object) -> int | None:
    """Return integer chart type for a shape when available."""
    chart = shape_chart(shape)
    if chart is None:
        return None
    return chart_type_value(chart)


def presentation_chart_types(presentation: object) -> list[int]:
    """Collect chart type values across all slides in a presentation."""
    chart_types: list[int] = []
    for slide in iter_slides(presentation):
        for shape in iter_shapes(slide):
            chart_type = shape_chart_type(shape)
            if chart_type is not None:
                chart_types.append(chart_type)
    return chart_types


def shape_has_text_frame(shape: object) -> bool:
    """Return whether ``shape`` has a text frame."""
    return bool(getattr(shape, "has_text_frame", False))


def shape_text_frame(shape: object) -> object | None:
    """Return text frame object when present."""
    if not shape_has_text_frame(shape):
        return None
    return getattr(shape, "text_frame", None)


class _MutableTextFrame(Protocol):
    text: str


def text_frame_text(text_frame: object) -> str | None:
    """Return text from a text frame object when present."""
    text = getattr(text_frame, "text", None)
    return text if isinstance(text, str) else None


def set_text_frame_text(text_frame: object, text: str) -> None:
    """Set text on a text frame object."""
    mutable_text_frame = cast(_MutableTextFrame, text_frame)
    mutable_text_frame.text = text


def shape_text_frame_text(shape: object) -> str | None:
    """Return text-frame text for a shape when present."""
    text_frame = shape_text_frame(shape)
    if text_frame is None:
        return None
    return text_frame_text(text_frame)


def text_frame_paragraphs(text_frame: object) -> list[object]:
    """Return paragraph objects from a text frame."""
    return _iter_objects(getattr(text_frame, "paragraphs", None))


def shape_text(shape: object) -> str | None:
    """Return ``shape.text`` when available."""
    text = getattr(shape, "text", None)
    return text if isinstance(text, str) else None


def shape_is_placeholder(shape: object) -> bool:
    """Return whether ``shape`` is a placeholder."""
    return bool(getattr(shape, "is_placeholder", False))


def shape_xml_element(shape: object) -> object | None:
    """Return underlying OOXML element for a shape when available."""
    return getattr(shape, "_element", None)


def shape_has_connector_endpoints(shape: object) -> bool:
    """Return whether shape exposes connector endpoints."""
    return getattr(shape, "begin_x", None) is not None and getattr(shape, "end_x", None) is not None


def chart_series_names(chart: object) -> list[str]:
    """Return series names for a chart object."""
    names: list[str] = []
    for series in _iter_objects(getattr(chart, "series", None)):
        names.append(str(getattr(series, "name", "")))
    return names
