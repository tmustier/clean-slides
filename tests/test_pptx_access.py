from __future__ import annotations

from dataclasses import dataclass, field

from clean_slides.pptx_access import (
    chart_first_plot,
    chart_plots,
    chart_series,
    chart_series_names,
    chart_type_value,
    chart_xml_element,
    chart_xml_space,
    iter_shapes,
    iter_slides,
    paragraph_font,
    paragraph_xml_element,
    plot_data_labels,
    point_fill_fore_color,
    point_fill_solid,
    point_line_fill_background,
    presentation_chart_types,
    series_points,
    set_chart_has_legend,
    set_font_size,
    set_plot_has_data_labels,
    set_text_frame_text,
    shape_chart,
    shape_chart_type,
    shape_has_chart,
    shape_has_connector_endpoints,
    shape_has_text_frame,
    shape_is_placeholder,
    shape_text,
    shape_text_frame,
    shape_text_frame_text,
    shape_xml_element,
    slide_add_chart,
    slide_size_emu,
    text_frame_paragraphs,
    text_frame_text,
    text_frame_xml_element,
)


def _object_list() -> list[object]:
    return []


def _chart_add_calls() -> list[tuple[object, object, object, object, object, object]]:
    return []


@dataclass
class _Series:
    name: str


@dataclass
class _Chart:
    chart_type: int
    series: list[object]
    _chartSpace: object | None = None
    _element: object | None = None
    plots: list[object] = field(default_factory=_object_list)
    has_legend: bool = True


@dataclass
class _TextFrame:
    text: str
    paragraphs: list[object]
    _element: object | None = None


@dataclass
class _Paragraph:
    _element: object


@dataclass
class _Shape:
    has_chart: bool = False
    chart: _Chart | None = None
    has_text_frame: bool = False
    text_frame: _TextFrame | None = None
    text: str | None = None
    is_placeholder: bool = False
    begin_x: int | None = None
    end_x: int | None = None
    _element: object | None = None


@dataclass
class _Slide:
    shapes: list[_Shape]


@dataclass
class _Presentation:
    slides: list[_Slide]


@dataclass
class _Fill:
    fore_color: object
    solid_calls: int = 0

    def solid(self) -> None:
        self.solid_calls += 1


@dataclass
class _LineFill:
    background_calls: int = 0

    def background(self) -> None:
        self.background_calls += 1


@dataclass
class _Line:
    fill: _LineFill


@dataclass
class _PointFormat:
    fill: _Fill
    line: _Line


@dataclass
class _Point:
    format: _PointFormat


@dataclass
class _SeriesWithPoints:
    points: list[_Point]


@dataclass
class _Plot:
    has_data_labels: bool = False
    data_labels: object | None = None


@dataclass
class _ChartFrame:
    has_chart: bool
    chart: _Chart


@dataclass
class _ShapesWithAddChart:
    chart: _Chart
    calls: list[tuple[object, object, object, object, object, object]] = field(
        default_factory=_chart_add_calls
    )

    def add_chart(
        self,
        chart_type: object,
        x: object,
        y: object,
        cx: object,
        cy: object,
        chart_data: object,
    ) -> _ChartFrame:
        self.calls.append((chart_type, x, y, cx, cy, chart_data))
        return _ChartFrame(has_chart=True, chart=self.chart)


@dataclass
class _PresentationRoot:
    slide_width: int
    slide_height: int


@dataclass
class _PresentationPart:
    presentation: _PresentationRoot


@dataclass
class _Package:
    presentation_part: _PresentationPart


@dataclass
class _Part:
    package: _Package


@dataclass
class _SlideWithPart:
    shapes: _ShapesWithAddChart
    part: _Part


@dataclass
class _Font:
    size: object | None = None


@dataclass
class _ParagraphWithFont:
    font: _Font


def test_presentation_chart_types_collects_chart_values() -> None:
    prs = _Presentation(
        slides=[
            _Slide(
                shapes=[
                    _Shape(has_chart=True, chart=_Chart(chart_type=57, series=[])),
                    _Shape(has_chart=False),
                    _Shape(has_chart=True, chart=_Chart(chart_type=58, series=[])),
                ]
            )
        ]
    )

    assert len(iter_slides(prs)) == 1
    assert len(iter_shapes(prs.slides[0])) == 3
    assert presentation_chart_types(prs) == [57, 58]


def test_shape_and_text_frame_helpers_cover_text_placeholder_connector() -> None:
    text_frame = _TextFrame(text="  Hello  ", paragraphs=[object()], _element={"tag": "txBody"})
    shape = _Shape(
        has_chart=True,
        chart=_Chart(chart_type=57, series=[_Series(name="Revenue")], _chartSpace={"tag": "cs"}),
        has_text_frame=True,
        text_frame=text_frame,
        text="inline",
        is_placeholder=True,
        begin_x=10,
        end_x=20,
        _element={"tag": "shape"},
    )

    assert shape_has_chart(shape) is True
    chart = shape_chart(shape)
    assert chart is not None
    assert chart_type_value(chart) == 57
    assert shape_chart_type(shape) == 57

    assert shape_has_text_frame(shape) is True
    frame = shape_text_frame(shape)
    assert frame is not None
    assert text_frame_text(frame) == "  Hello  "
    assert shape_text_frame_text(shape) == "  Hello  "
    assert text_frame_xml_element(frame) == {"tag": "txBody"}
    assert len(text_frame_paragraphs(frame)) == 1

    assert shape_text(shape) == "inline"
    assert shape_is_placeholder(shape) is True
    assert shape_has_connector_endpoints(shape) is True
    assert shape_xml_element(shape) == {"tag": "shape"}

    assert chart_xml_space(chart) == {"tag": "cs"}

    paragraph = _Paragraph(_element={"tag": "p"})
    assert paragraph_xml_element(paragraph) == {"tag": "p"}

    set_text_frame_text(frame, "Updated")
    assert text_frame.text == "Updated"

    assert chart_series_names(chart) == ["Revenue"]


def test_chart_mutation_and_slide_helpers_cover_dynamic_graphic_access() -> None:
    point = _Point(
        format=_PointFormat(
            fill=_Fill(fore_color={"kind": "accent"}),
            line=_Line(fill=_LineFill()),
        )
    )
    series = _SeriesWithPoints(points=[point])
    plot = _Plot(data_labels={"k": "v"})
    chart = _Chart(
        chart_type=57,
        series=[series],
        _element={"tag": "chart"},
        plots=[plot],
    )

    shapes = _ShapesWithAddChart(chart=chart)
    slide = _SlideWithPart(
        shapes=shapes,
        part=_Part(
            package=_Package(
                presentation_part=_PresentationPart(
                    presentation=_PresentationRoot(slide_width=1280, slide_height=720)
                )
            )
        ),
    )

    frame = slide_add_chart(slide, chart_type=57, x=1, y=2, cx=3, cy=4, chart_data={"d": 1})
    assert frame is not None
    assert shape_chart(frame) is chart
    assert len(shapes.calls) == 1

    assert chart_xml_element(chart) == {"tag": "chart"}
    set_chart_has_legend(chart, False)
    assert chart.has_legend is False

    assert chart_series(chart) == [series]
    assert series_points(series) == [point]

    point_fill_solid(point)
    assert point.format.fill.solid_calls == 1
    assert point_fill_fore_color(point) == {"kind": "accent"}

    point_line_fill_background(point)
    assert point.format.line.fill.background_calls == 1

    assert chart_plots(chart) == [plot]
    assert chart_first_plot(chart) is plot

    set_plot_has_data_labels(plot, True)
    assert plot.has_data_labels is True
    assert plot_data_labels(plot) == {"k": "v"}

    assert slide_size_emu(slide) == (1280, 720)

    paragraph = _ParagraphWithFont(font=_Font())
    font = paragraph_font(paragraph)
    assert font is not None
    set_font_size(font, 11)
    assert paragraph.font.size == 11
