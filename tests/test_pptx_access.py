from __future__ import annotations

from dataclasses import dataclass

from clean_slides.pptx_access import (
    chart_series_names,
    chart_type_value,
    chart_xml_space,
    iter_shapes,
    iter_slides,
    paragraph_xml_element,
    presentation_chart_types,
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
    text_frame_paragraphs,
    text_frame_text,
)


@dataclass
class _Series:
    name: str


@dataclass
class _Chart:
    chart_type: int
    series: list[_Series]
    _chartSpace: object | None = None


@dataclass
class _TextFrame:
    text: str
    paragraphs: list[object]


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
    text_frame = _TextFrame(text="  Hello  ", paragraphs=[object()])
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
