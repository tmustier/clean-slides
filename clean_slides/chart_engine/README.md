# `clean_slides.chart_engine`

Internal chart-generation modules used by `clean_slides/chart_generator.py`.

This package is **not** a public API yet. It exists to keep chart logic modular, typed, and easier to evolve.

## Why this package exists

Historically, most chart behavior lived in one large file (`chart_generator.py`).
The code is now split so payload building, styling, overlays, and template operations are isolated.

## Module map

### Core utilities
- `defaults.py` — shared constants/default values
- `colors.py` — color parsing + theme/rgb application
- `units.py` — EMU coercion, line width coercion, path resolution
- `spec_utils.py` — spec normalization helpers (series/categories/values)
- `geometry.py` — orientation + chart coordinate helpers
- `text_style.py` — label/alignment/anchor normalization
- `text_templates.py` — cached txBody template load/apply/resolve
- `template_ops.py` — chart XML replacement + template verification helpers

### Chart construction
- `builder.py` — slide/chart orchestration (`build_chart`) and chart-template data-label copy helper
- `payloads.py` — builds bar/waterfall chart data + metadata payloads
- `style.py` — applies chart/axis/series/data-label style to `python-pptx` chart objects

### Overlays and annotations
- `annotations.py` — textbox/shape/line annotation primitives + waterfall title helper
- `overlay_bar.py` — bar overlay orchestration (annotations + helper wiring)
- `overlay_bar_totals_categories.py` — bar total and category label rendering paths
- `overlay_bar_legend.py` — bar legend labels and marker rendering
- `overlay_bar_segments.py` — bar segment-label placement logic
- `overlay_waterfall.py` — waterfall overlay orchestration (wires geometry/config into specialized helpers)
- `overlay_waterfall_connectors.py` — waterfall connector segment rendering
- `overlay_waterfall_labels.py` — waterfall value/category/series label layout + rendering
- `overlay_waterfall_data_labels.py` — manual per-point waterfall data-label layout in chart XML
- `overlays.py` — compatibility facade that re-exports overlay entrypoints

## Main entrypoints currently consumed by `chart_generator.py`

- Builder:
  - `build_chart`
  - `apply_data_label_style`
- Payloads:
  - `build_bar_payload`
  - `build_waterfall_payload`
- Styling:
  - `apply_series_colors`
  - `apply_bar_chart_style`
  - `apply_waterfall_style`
  - `apply_waterfall_chart_style`
  - `apply_waterfall_data_labels`
- Overlays:
  - `add_bar_overlays`
  - `add_waterfall_overlays`
  - `apply_waterfall_data_label_layout`
- Color helper used by chart-cells:
  - `apply_color`

## Conventions

- Prefer behavior-preserving refactors first; style/geometry changes should be explicit.
- Keep EMU/unit coercion centralized in `units.py`.
- Keep spec normalization in `spec_utils.py` (avoid ad-hoc parsing in overlay/style code).
- If you split a module, keep a small facade/re-export when useful to avoid broad call-site churn.

## Validation checklist

From repo root:

```bash
.venv/bin/ruff check clean_slides/chart_engine clean_slides/chart_generator.py
.venv/bin/pyright
.venv/bin/pytest -q
```

## Next decomposition targets

- Keep adding dedicated chart-engine smoke tests at the module boundary (builder + overlays).
- Continue simplifying internal coercion helpers now that per-file pyright suppressions are removed.
- Keep trimming compatibility glue in `chart_generator.py` as typed wrappers absorb more behavior.
