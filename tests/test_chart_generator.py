from __future__ import annotations

from pathlib import Path

import pytest

from clean_slides.chart_engine.defaults import EXPECTED_TEMPLATE_ALIASES
from clean_slides.chart_generator import normalize_chart_specs, resolve_expected_template


def _chart_spec(value: int) -> dict[str, object]:
    return {
        "categories": ["A"],
        "series": [{"name": "S1", "values": [value]}],
    }


def test_normalize_chart_specs_single_spec() -> None:
    charts, deck_meta = normalize_chart_specs(_chart_spec(1))

    assert len(charts) == 1
    assert deck_meta == {}
    assert charts[0]["series"] == [{"name": "S1", "values": [1]}]


def test_normalize_chart_specs_multi_spec_with_deck_meta() -> None:
    raw = {
        "template": "templates/base.pptx",
        "layout": "Default",
        "charts": [_chart_spec(1), _chart_spec(2)],
    }

    charts, deck_meta = normalize_chart_specs(raw)

    assert len(charts) == 2
    assert deck_meta == {"template": "templates/base.pptx", "layout": "Default"}


def test_normalize_chart_specs_rejects_non_list_charts() -> None:
    raw = {
        "charts": {"categories": ["A"], "series": [{"name": "S1", "values": [1]}]},
    }

    with pytest.raises(ValueError, match=r"'charts' must be a JSON array"):
        normalize_chart_specs(raw)


def test_resolve_expected_template_alias() -> None:
    resolved = resolve_expected_template({}, Path("/tmp/spec.json"), "clean-slides")

    assert resolved == EXPECTED_TEMPLATE_ALIASES["clean-slides"]


def test_resolve_expected_template_relative_from_spec() -> None:
    spec_path = Path("/tmp/example/specs/charts.json")
    resolved = resolve_expected_template(
        {"expected_template": "templates/consulting-template.pptx"},
        spec_path,
        None,
    )

    expected = (spec_path.parent / "templates" / "consulting-template.pptx").resolve()
    assert resolved == expected
