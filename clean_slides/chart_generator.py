#!/usr/bin/env python3
# pyright: reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false
# pyright: reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false
# pyright: reportMissingParameterType=false
# pyright: reportMissingTypeArgument=false
# pyright: reportGeneralTypeIssues=false
# pyright: reportUnnecessaryIsInstance=false
"""Chart generator: generate bar/column or waterfall charts from JSON.

Layers:
1) tiny hidden shape named "chart data - do not delete"
2) PowerPoint chart (clustered/stacked/waterfall)
3) optional overlay labels as text boxes
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from pptx import Presentation

from clean_slides.chart_engine.builder import apply_data_label_style, build_chart
from clean_slides.chart_engine.colors import apply_color
from clean_slides.chart_engine.defaults import EXPECTED_TEMPLATE_ALIASES
from clean_slides.chart_engine.overlays import add_waterfall_overlays
from clean_slides.chart_engine.payloads import build_bar_payload, build_waterfall_payload
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


def resolve_expected_template(
    spec: dict, spec_path: Path, expected_template: str | Path | None
) -> Path | None:
    raw = expected_template if expected_template is not None else spec.get("expected_template")
    if not raw:
        return None
    if isinstance(raw, Path):
        key = raw.as_posix()
    elif isinstance(raw, str):
        key = raw.strip()
    else:
        raise ValueError("expected_template must be a string path or alias")
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
        raise ValueError(
            "Template mismatch. expected {expected} (theme={expected_theme}, master={expected_sig}) "
            "but got {actual} (theme={actual_theme}, master={actual_sig}).".format(
                expected=expected,
                actual=actual,
                expected_theme=theme_name(expected),
                actual_theme=theme_name(actual),
                expected_sig=expected_sig,
                actual_sig=actual_sig,
            )
        )


def load_spec(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def validate_chart_spec(spec: dict) -> None:
    if "categories" not in spec or "series" not in spec:
        raise ValueError("Input JSON must contain 'categories' and 'series'.")


def normalize_chart_specs(raw: Any) -> tuple[list[dict], dict]:
    if isinstance(raw, list):
        charts = raw
        deck_meta: dict[str, Any] = {}
    elif isinstance(raw, dict) and "charts" in raw:
        charts = raw.get("charts") or []
        deck_meta = {key: value for key, value in raw.items() if key != "charts"}
    else:
        charts = [raw]
        deck_meta = {}

    if not charts:
        raise ValueError("No charts found in input JSON.")

    normalized: list[dict] = []
    for spec in charts:
        if not isinstance(spec, dict):
            raise ValueError("Each chart spec must be a JSON object.")
        validate_chart_spec(spec)
        normalized.append(spec)

    return normalized, deck_meta


def main() -> None:
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
    args = parser.parse_args()

    raw_spec = load_spec(args.input)
    chart_specs, deck_meta = normalize_chart_specs(raw_spec)

    expected_spec = deck_meta if deck_meta else chart_specs[0]
    expected_template = resolve_expected_template(expected_spec, args.input, args.expected_template)

    template_path = args.template
    deck_template = deck_meta.get("template") if isinstance(deck_meta, dict) else None
    if template_path is None and deck_template:
        template_path = Path(deck_template)
        if not template_path.is_absolute():
            template_path = (args.input.parent / template_path).resolve()

    if expected_template and template_path is None:
        template_path = expected_template
    if expected_template:
        ensure_expected_template(expected_template, template_path)

    layout_name = (
        args.layout
        or (deck_meta.get("layout") if isinstance(deck_meta, dict) else None)
        or "Default"
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

    prs.save(args.output)
    if template_replacements:
        apply_chart_template_replacements(args.output, template_replacements)


if __name__ == "__main__":
    main()
