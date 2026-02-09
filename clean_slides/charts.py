"""Chart generation wrapper for the clean-slides CLI.

Loads the JSON chart generator module and runs it without exposing
internal implementation details in the CLI layer.
"""

from __future__ import annotations

import importlib.util
import os
from pathlib import Path
from types import ModuleType
from typing import Any

_CHARTS_ENV_VAR = "CLEAN_SLIDES_CHARTS_PATH"


def _to_str_dict(value: object) -> dict[str, Any]:
    """Safely coerce an unknown mapping to ``dict[str, Any]``."""
    if not isinstance(value, dict):
        return {}
    result: dict[str, Any] = {}
    items: list[tuple[str, Any]] = list(value.items())  # type: ignore[arg-type]
    for k, v in items:
        result[k] = v
    return result
_MODULE_CACHE: dict[Path, ModuleType] = {}


def _coerce_module_path(raw: str | Path) -> Path:
    path = Path(raw).expanduser()
    if path.is_dir():
        path = path / "generate_bar_chart.py"
    return path.resolve()


def resolve_charts_module_path(input_path: Path, module_path: str | None) -> Path:
    if module_path:
        candidate = _coerce_module_path(module_path)
    else:
        env_value = os.getenv(_CHARTS_ENV_VAR)
        if env_value:
            candidate = _coerce_module_path(env_value)
        else:
            for base in (input_path.parent, Path.cwd()):
                candidate = base / "generate_bar_chart.py"
                if candidate.is_file():
                    return candidate.resolve()
            raise FileNotFoundError(
                "Charts module not found. Pass --module-path or set "
                f"{_CHARTS_ENV_VAR} to the chart generator script."
            )

    if not candidate.is_file():
        raise FileNotFoundError(
            "Charts module not found. Pass --module-path or set "
            f"{_CHARTS_ENV_VAR} to the chart generator script."
        )
    return candidate


def load_charts_module(module_path: Path) -> ModuleType:
    module_path = module_path.resolve()
    cached = _MODULE_CACHE.get(module_path)
    if cached is not None:
        return cached

    spec = importlib.util.spec_from_file_location("clean_slides._charts_ext", module_path)
    if spec is None or spec.loader is None:
        raise ImportError("Unable to load chart generator module")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    _MODULE_CACHE[module_path] = module
    return module


def generate_charts_from_json(
    input_path: Path,
    output_path: Path,
    *,
    template: Path | None = None,
    layout: str | None = None,
    expected_template: str | None = None,
    module_path: str | None = None,
) -> None:
    charts_module_path = resolve_charts_module_path(input_path, module_path)
    charts = load_charts_module(charts_module_path)

    raw_spec = charts.load_spec(input_path)
    chart_specs, deck_meta = charts.normalize_chart_specs(raw_spec)

    expected_spec: dict[str, Any] = deck_meta if deck_meta else chart_specs[0]
    expected = charts.resolve_expected_template(expected_spec, input_path, expected_template)

    deck_meta_dict: dict[str, Any] = _to_str_dict(deck_meta)

    template_path: Path | None = template
    deck_template: str | None = deck_meta_dict.get("template")
    if template_path is None and deck_template:
        template_path = Path(deck_template)
        if not template_path.is_absolute():
            template_path = (input_path.parent / template_path).resolve()

    if expected and template_path is None:
        template_path = expected
    if expected:
        charts.ensure_expected_template(expected, template_path)

    layout_name: str = layout or deck_meta_dict.get("layout", "Default") or "Default"

    if template_path:
        prs = charts.Presentation(str(template_path))
    else:
        prs = charts.Presentation()

    template_replacements: list[Any] = []
    for idx, spec in enumerate(chart_specs):
        if idx > 0 and "append_slide" not in spec:
            spec["append_slide"] = True
        template_replacements.extend(
            charts.build_chart(
                prs,
                spec,
                output_path,
                template_path=template_path,
                layout_name=layout_name,
                save=False,
                defer_template_copy=True,
            )
        )

    prs.save(output_path)
    if template_replacements:
        charts.apply_chart_template_replacements(output_path, template_replacements)
