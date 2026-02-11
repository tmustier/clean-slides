"""Typed facade for chart generation.

Uses the bundled chart engine implementation shipped inside ``clean_slides``.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Protocol, cast

from . import chart_generator


class ChartEngine(Protocol):
    """Structural interface implemented by the bundled chart engine module."""

    Presentation: Any

    def load_spec(self, path: Path) -> Any: ...

    def normalize_chart_specs(self, raw: Any) -> tuple[list[dict[str, Any]], dict[str, Any]]: ...

    def resolve_expected_template(
        self,
        spec: dict[str, Any],
        spec_path: Path,
        expected_template: str | Path | None,
    ) -> Path | None: ...

    def ensure_expected_template(self, expected: Path, actual: Path | None) -> None: ...

    def build_chart(
        self,
        prs: Any,
        spec: dict[str, Any],
        output_path: Path,
        template_path: Path | None = None,
        layout_name: str | None = None,
        save: bool = True,
        defer_template_copy: bool = False,
    ) -> list[Any]: ...

    def apply_chart_template_replacements(
        self, output_path: Path, replacements: list[Any]
    ) -> None: ...

    def build_bar_payload(self, spec: dict[str, Any]) -> tuple[Any, Any, dict[str, Any]]: ...

    def build_waterfall_payload(self, spec: dict[str, Any]) -> tuple[Any, Any, dict[str, Any]]: ...

    def apply_series_colors(self, chart: Any, colors: list[str | None]) -> None: ...

    def apply_bar_chart_style(self, chart: Any, meta: dict[str, Any]) -> None: ...

    def apply_data_label_style(self, labels: Any, data_cfg: dict[str, Any]) -> None: ...

    def apply_waterfall_style(self, chart: Any, meta: dict[str, Any]) -> None: ...

    def apply_waterfall_chart_style(self, chart: Any, meta: dict[str, Any]) -> None: ...

    def add_waterfall_overlays(
        self,
        slide: Any,
        chart_box: tuple[int, int, int, int],
        meta: dict[str, Any],
        slide_size: tuple[int, int] | None = None,
    ) -> None: ...

    def apply_color(self, target: Any, value: Any) -> bool: ...


def _to_str_dict(value: object) -> dict[str, Any]:
    """Safely coerce an unknown mapping to ``dict[str, Any]``."""
    if not isinstance(value, dict):
        return {}

    result: dict[str, Any] = {}
    items: list[tuple[str, Any]] = list(value.items())  # type: ignore[arg-type]
    for k, v in items:
        result[k] = v
    return result


def load_chart_engine() -> ChartEngine:
    """Return the bundled chart engine implementation."""
    return cast(ChartEngine, chart_generator)


def load_charts_module() -> ChartEngine:
    """Backward-compatible alias for callers still using the old name."""
    return load_chart_engine()


def resolve_charts_module_path() -> Path:
    """Return the filesystem path to the bundled chart engine module."""
    module_file = chart_generator.__file__
    return Path(module_file).resolve()


def generate_charts_from_json(
    input_path: Path,
    output_path: Path,
    *,
    template: Path | None = None,
    layout: str | None = None,
    expected_template: str | None = None,
) -> None:
    charts = load_chart_engine()

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
        if "_base_dir" not in spec:
            spec["_base_dir"] = str(input_path.parent.resolve())
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
