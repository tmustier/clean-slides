from __future__ import annotations

from clean_slides.content import SUB_COLOR, Paragraph, normalize_cell


def _default_header() -> Paragraph:
    return Paragraph(
        text="",
        lvl=0,
        font="Aptos Display",
        size_pt=14,
        color="tx1",
        bold=True,
    )


def test_normalize_cell_adds_sub_paragraph_with_expected_style() -> None:
    paragraphs = normalize_cell(
        {
            "text": "CAGR",
            "sub": "%, FY26-33E",
        },
        _default_header(),
        parse_bullets=False,
    )

    assert len(paragraphs) == 2
    main = paragraphs[0]
    sub = paragraphs[1]

    assert main.text == "CAGR"
    assert sub.text == "%, FY26-33E"

    assert sub.lvl == main.lvl
    assert sub.font == main.font
    assert sub.size_pt == main.size_pt
    assert sub.color == SUB_COLOR
    assert sub.bold is False


def test_normalize_cell_sub_paragraph_keeps_explicit_typography_but_body_color() -> None:
    paragraphs = normalize_cell(
        {
            "text": "Impact",
            "sub": "at constant prices",
            "lvl": 1,
            "font": "Arial",
            "size": 11,
            "color": "accent1",
            "bold": True,
        },
        _default_header(),
        parse_bullets=False,
    )

    assert len(paragraphs) == 2
    main = paragraphs[0]
    sub = paragraphs[1]

    assert main.lvl == 1
    assert main.font == "Arial"
    assert main.size_pt == 11
    assert main.color == "accent1"
    assert main.bold is True

    assert sub.lvl == 1
    assert sub.font == "Arial"
    assert sub.size_pt == 11
    assert sub.color == SUB_COLOR
    assert sub.bold is False
