#!/usr/bin/env python3
"""Chart generator: generate bar/column or waterfall charts from JSON.

Layers:
1) tiny hidden shape named "chart data - do not delete"
2) PowerPoint chart (clustered/stacked/waterfall)
3) optional overlay labels as text boxes
"""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import cast

from pptx import Presentation

from clean_slides.chart_engine.builder import apply_data_label_style, build_chart
from clean_slides.chart_engine.colors import apply_color
from clean_slides.chart_engine.defaults import EXPECTED_TEMPLATE_ALIASES
from clean_slides.chart_engine.overlays import add_waterfall_overlays
from clean_slides.chart_engine.payloads import build_bar_payload, build_waterfall_payload
from clean_slides.chart_engine.spec_utils import object_list, str_key_dict
from clean_slides.chart_engine.style import (
    apply_bar_chart_style,
    apply_series_colors,
    apply_waterfall_chart_style,
    apply_waterfall_style,
)
from clean_slides.chart_engine.template_ops import (
    ChartTemplateReplacement,
    apply_chart_template_replacements,
    slide_master_signature,
    theme_name,
)

ChartSpec = dict[str, object]
DeckMeta = dict[str, object]


@dataclass(frozen=True)
class ChartCliArgs:
    input: Path
    output: Path
    template: Path | None
    layout: str | None
    expected_template: str | None


__all__ = [
    "Presentation",
    "add_waterfall_overlays",
    "apply_bar_chart_style",
    "apply_chart_template_replacements",
    "apply_color",
    "apply_data_label_style",
    "apply_series_colors",
    "apply_waterfall_chart_style",
    "apply_waterfall_style",
    "build_bar_payload",
    "build_chart",
    "build_waterfall_payload",
    "ensure_expected_template",
    "load_spec",
    "main",
    "normalize_chart_specs",
    "resolve_expected_template",
]


def parse_cli_args() -> ChartCliArgs:
    parser = argparse.ArgumentParser(description="Generate a bar/column chart PPTX from JSON")
    parser.add_argument("input", type=Path, help="Path to JSON spec")
    parser.add_argument("output", type=Path, help="Path to output .pptx")
    parser.add_argument("--template", type=Path, help="Template PPTX to use as base")
    parser.add_argument("--layout", type=str, help="Layout name when using a template")
    parser.add_argument(
        "--expected-template",
        type=str,
        help="Expected template path or alias (e.g. clean-slides) to verify",
    )

    namespace = parser.parse_args()
    input_path = getattr(namespace, "input", None)
    output_path = getattr(namespace, "output", None)
    template_path = getattr(namespace, "template", None)
    layout_name = getattr(namespace, "layout", None)
    expected_template = getattr(namespace, "expected_template", None)

    if not isinstance(input_path, Path):
        raise TypeError("input must resolve to a pathlib.Path")
    if not isinstance(output_path, Path):
        raise TypeError("output must resolve to a pathlib.Path")
    if template_path is not None and not isinstance(template_path, Path):
        raise TypeError("template must resolve to a pathlib.Path")
    if layout_name is not None and not isinstance(layout_name, str):
        raise TypeError("layout must be a string when provided")
    if expected_template is not None and not isinstance(expected_template, str):
        raise TypeError("expected_template must be a string when provided")

    return ChartCliArgs(
        input=input_path,
        output=output_path,
        template=template_path,
        layout=layout_name,
        expected_template=expected_template,
    )


def resolve_expected_template(
    spec: ChartSpec, spec_path: Path, expected_template: str | Path | None
) -> Path | None:
    raw = expected_template if expected_template is not None else spec.get("expected_template")
    if raw is None:
        return None

    if isinstance(raw, Path):
        key = raw.as_posix()
    elif isinstance(raw, str):
        key = raw.strip()
    else:
        raise ValueError("expected_template must be a string path or alias")

    if not key:
        return None

    alias = EXPECTED_TEMPLATE_ALIASES.get(key.lower())
    if alias is not None:
        return alias

    path = Path(key)
    if not path.is_absolute():
        path = (spec_path.parent / path).resolve()
    return path


def ensure_expected_template(expected: Path, actual: Path | None) -> None:
    if actual is None:
        raise ValueError(f"Expected template {expected}, but no template was provided")
    if not expected.exists():
        raise FileNotFoundError(f"Expected template not found: {expected}")
    if not actual.exists():
        raise FileNotFoundError(f"Template not found: {actual}")

    expected_sig = slide_master_signature(expected)
    actual_sig = slide_master_signature(actual)
    if expected_sig != actual_sig:
        expected_theme = theme_name(expected)
        actual_theme = theme_name(actual)
        raise ValueError(
            f"Template mismatch. expected {expected} (theme={expected_theme}, master={expected_sig}) "
            f"but got {actual} (theme={actual_theme}, master={actual_sig})."
        )


def load_spec(path: Path) -> object:
    return json.loads(path.read_text(encoding="utf-8"))


def validate_chart_spec(spec: ChartSpec) -> None:
    if "categories" not in spec or "series" not in spec:
        raise ValueError("Input JSON must contain 'categories' and 'series'.")


def normalize_chart_specs(raw: object) -> tuple[list[ChartSpec], DeckMeta]:
    charts_raw: list[object]
    deck_meta: DeckMeta = {}

    if isinstance(raw, list):
        charts_raw = object_list(cast(object, raw))
    elif isinstance(raw, dict):
        raw_dict = str_key_dict(cast(object, raw))
        if "charts" in raw_dict:
            charts_value = raw_dict.get("charts")
            if charts_value is None:
                charts_raw = []
            elif isinstance(charts_value, list):
                charts_raw = object_list(cast(object, charts_value))
            else:
                raise ValueError("'charts' must be a JSON array when provided.")
            deck_meta = {key: value for key, value in raw_dict.items() if key != "charts"}
        else:
            charts_raw = [raw_dict]
    else:
        charts_raw = [raw]

    if not charts_raw:
        raise ValueError("No charts found in input JSON.")

    normalized: list[ChartSpec] = []
    for spec in charts_raw:
        if not isinstance(spec, dict):
            raise ValueError("Each chart spec must be a JSON object.")
        normalized_spec = str_key_dict(cast(object, spec))
        validate_chart_spec(normalized_spec)
        normalized.append(normalized_spec)

    return normalized, deck_meta


def main() -> None:
    args = parse_cli_args()

    raw_spec = load_spec(args.input)
    chart_specs, deck_meta = normalize_chart_specs(raw_spec)

    expected_spec = deck_meta if deck_meta else chart_specs[0]
    expected_template = resolve_expected_template(expected_spec, args.input, args.expected_template)

    template_path = args.template
    deck_template = deck_meta.get("template")
    if template_path is None and isinstance(deck_template, str) and deck_template:
        template_path = Path(deck_template)
        if not template_path.is_absolute():
            template_path = (args.input.parent / template_path).resolve()

    if expected_template and template_path is None:
        template_path = expected_template
    if expected_template:
        ensure_expected_template(expected_template, template_path)

    deck_layout = deck_meta.get("layout")
    layout_name = (
        args.layout or (deck_layout if isinstance(deck_layout, str) else None) or "Default"
    )

    if template_path:
        prs = Presentation(str(template_path))
    else:
        prs = Presentation()

    template_replacements: list[ChartTemplateReplacement] = []
    for idx, spec in enumerate(chart_specs):
        if idx > 0 and "append_slide" not in spec:
            spec["append_slide"] = True
        if "_base_dir" not in spec:
            spec["_base_dir"] = str(args.input.parent.resolve())

        template_replacements.extend(
            build_chart(
                prs,
                spec,
                args.output,
                template_path=template_path,
                layout_name=layout_name,
                save=False,
                defer_template_copy=True,
            )
        )

    prs.save(str(args.output))
    if template_replacements:
        apply_chart_template_replacements(args.output, template_replacements)


if __name__ == "__main__":
    main()
