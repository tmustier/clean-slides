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


def paragraph_xml_element(paragraph: object) -> object | None:
    """Return underlying OOXML paragraph element when available."""
    return getattr(paragraph, "_element", None)


def text_frame_xml_element(text_frame: object) -> object | None:
    """Return underlying OOXML text-frame element when available."""
    return getattr(text_frame, "_element", None)


def chart_xml_space(chart: object) -> object | None:
    """Return underlying OOXML chart-space element when available."""
    return getattr(chart, "_chartSpace", None)


def chart_xml_element(chart: object) -> object | None:
    """Return underlying OOXML chart element when available."""
    return getattr(chart, "_element", None)


def chart_part_name(chart: object) -> str | None:
    """Return chart partname (e.g. '/ppt/charts/chart1.xml') when available."""
    chart_part = getattr(chart, "part", None)
    partname = getattr(chart_part, "partname", None)
    return str(partname) if partname is not None else None


class _AddChartCallable(Protocol):
    def __call__(
        self,
        chart_type: object,
        x: object,
        y: object,
        cx: object,
        cy: object,
        chart_data: object,
    ) -> object: ...


def slide_add_chart(
    slide: object,
    chart_type: object,
    x: object,
    y: object,
    cx: object,
    cy: object,
    chart_data: object,
) -> object | None:
    """Add a chart to a slide when supported and return the chart frame."""
    shapes = getattr(slide, "shapes", None)
    add_chart = getattr(shapes, "add_chart", None)
    if not callable(add_chart):
        return None
    add_chart_fn = cast(_AddChartCallable, add_chart)
    return add_chart_fn(chart_type, x, y, cx, cy, chart_data)


class _MutableChartLegend(Protocol):
    has_legend: bool


def set_chart_has_legend(chart: object, has_legend: bool) -> None:
    """Set chart legend visibility."""
    mutable_chart = cast(_MutableChartLegend, chart)
    mutable_chart.has_legend = has_legend


def chart_series(chart: object) -> list[object]:
    """Return chart series objects."""
    return _iter_objects(getattr(chart, "series", None))


def chart_plots(chart: object) -> list[object]:
    """Return chart plot objects."""
    return _iter_objects(getattr(chart, "plots", None))


def chart_first_plot(chart: object) -> object | None:
    """Return chart first plot when present."""
    plots = chart_plots(chart)
    return plots[0] if plots else None


class _MutablePlotDataLabels(Protocol):
    has_data_labels: bool


class _PlotDataLabels(Protocol):
    @property
    def data_labels(self) -> object: ...


def set_plot_has_data_labels(plot: object, has_data_labels: bool) -> None:
    """Set data-label visibility for a chart plot."""
    mutable_plot = cast(_MutablePlotDataLabels, plot)
    mutable_plot.has_data_labels = has_data_labels


def plot_data_labels(plot: object) -> object | None:
    """Return plot data-label collection when available."""
    return getattr(cast(_PlotDataLabels, plot), "data_labels", None)


def series_points(series: object) -> list[object]:
    """Return points from a chart series."""
    return _iter_objects(getattr(series, "points", None))


def point_fill_solid(point: object) -> None:
    """Convert point fill to solid when supported."""
    point_format = getattr(point, "format", None)
    fill = getattr(point_format, "fill", None)
    solid = getattr(fill, "solid", None)
    if callable(solid):
        solid()


def point_fill_fore_color(point: object) -> object | None:
    """Return point fill foreground color when available."""
    point_format = getattr(point, "format", None)
    fill = getattr(point_format, "fill", None)
    return getattr(fill, "fore_color", None)


def point_line_fill_background(point: object) -> None:
    """Set point line fill background when supported."""
    point_format = getattr(point, "format", None)
    line = getattr(point_format, "line", None)
    fill = getattr(line, "fill", None)
    background = getattr(fill, "background", None)
    if callable(background):
        background()


def slide_size_emu(slide: object) -> tuple[int, int] | None:
    """Return presentation slide size for a slide in EMU."""
    part = getattr(slide, "part", None)
    package = getattr(part, "package", None)
    presentation_part = getattr(package, "presentation_part", None)
    presentation = getattr(presentation_part, "presentation", None)
    width = getattr(presentation, "slide_width", None)
    height = getattr(presentation, "slide_height", None)
    if width is None or height is None:
        return None
    try:
        return (int(width), int(height))
    except (TypeError, ValueError):
        return None


def paragraph_font(paragraph: object) -> object | None:
    """Return paragraph font object when available."""
    return getattr(paragraph, "font", None)


class _MutableFont(Protocol):
    size: object


def set_font_size(font: object, size: object) -> None:
    """Set font size for a paragraph font object."""
    mutable_font = cast(_MutableFont, font)
    mutable_font.size = size


def shape_has_connector_endpoints(shape: object) -> bool:
    """Return whether shape exposes connector endpoints."""
    return getattr(shape, "begin_x", None) is not None and getattr(shape, "end_x", None) is not None


def chart_series_names(chart: object) -> list[str]:
    """Return series names for a chart object."""
    names: list[str] = []
    for series in _iter_objects(getattr(chart, "series", None)):
        names.append(str(getattr(series, "name", "")))
    return names
