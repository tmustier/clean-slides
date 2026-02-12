from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE

from clean_slides.chart_engine.builder import build_chart


def _chart_types(path: Path) -> list[int]:
    prs = Presentation(str(path))
    result: list[int] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not bool(getattr(shape, "has_chart", False)):
                continue
            chart = getattr(shape, "chart", None)
            if chart is None:
                continue
            chart_type = getattr(chart, "chart_type", None)
            if chart_type is None:
                continue
            result.append(int(chart_type))
    return result


def _shape_texts(path: Path) -> list[str]:
    prs = Presentation(str(path))
    texts: list[str] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not bool(getattr(shape, "has_text_frame", False)):
                continue
            text_frame = getattr(shape, "text_frame", None)
            if text_frame is None:
                continue
            text_value = getattr(text_frame, "text", None)
            if not isinstance(text_value, str):
                continue
            text = text_value.strip()
            if text:
                texts.append(text)
    return texts


def test_build_chart_bar_with_overlays_smoke(tmp_path: Path) -> None:
    output_path = tmp_path / "bar-smoke.pptx"

    spec: dict[str, object] = {
        "type": "clustered",
        "categories": ["A", "B"],
        "series": [{"name": "S1", "values": [10, 20], "color": "accent1"}],
        "show_data_labels": True,
        "add_overlay_labels": True,
        "bar": {
            "orientation": "horizontal",
            "show_totals": True,
            "segment_labels": [{"show": True, "series_indices": [0]}],
        },
    }

    prs = Presentation()
    replacements = build_chart(prs, spec, output_path)

    assert replacements == []
    assert output_path.exists()
    assert _chart_types(output_path) == [int(XL_CHART_TYPE.BAR_CLUSTERED)]

    texts = _shape_texts(output_path)
    assert "A" in texts
    assert "B" in texts


def test_build_chart_waterfall_with_overlays_smoke(tmp_path: Path) -> None:
    output_path = tmp_path / "waterfall-smoke.pptx"

    spec: dict[str, object] = {
        "type": "waterfall",
        "categories": ["Start", "Growth", "Costs", "End"],
        "series": [{"name": "Values", "values": [100, 40, -20, 120], "color": "accent1"}],
        "show_data_labels": True,
        "add_overlay_labels": True,
        "waterfall": {
            "total_categories": ["Start", "End"],
            "decrease_categories": ["Costs"],
            "connector_style": "gap",
        },
    }

    prs = Presentation()
    replacements = build_chart(prs, spec, output_path)

    assert replacements == []
    assert output_path.exists()
    assert _chart_types(output_path) == [int(XL_CHART_TYPE.COLUMN_STACKED)]

    texts = _shape_texts(output_path)
    assert "Start" in texts
    assert "End" in texts


def test_build_chart_defers_chart_template_copy_replacement(tmp_path: Path) -> None:
    template_path = tmp_path / "template-without-charts.pptx"
    Presentation().save(str(template_path))

    output_path = tmp_path / "deferred-template-copy.pptx"
    spec: dict[str, object] = {
        "type": "clustered",
        "categories": ["A"],
        "series": [{"name": "S1", "values": [1], "color": "accent1"}],
        "bar": {
            "chart_template": str(template_path),
            "chart_template_copy": True,
        },
    }

    prs = Presentation()
    replacements = build_chart(
        prs,
        spec,
        output_path,
        save=False,
        defer_template_copy=True,
    )

    assert len(replacements) == 1
    replacement = replacements[0]
    assert replacement.template_path == template_path
    assert replacement.chart_part.startswith("ppt/charts/chart")
