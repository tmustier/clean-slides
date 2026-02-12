# Release Checklist

Use this checklist before cutting a release or tagging a stable snapshot.

## 1) Local quality gate

Run from repo root:

```bash
.venv/bin/ruff check clean_slides tests
.venv/bin/pyright
.venv/bin/pytest -q
.venv/bin/pre-commit run --all-files
```

Expected:
- no lint/type errors
- all tests green
- no formatting drift

## 2) Canary behavior gate

Run the dedicated canary suite:

```bash
.venv/bin/pytest -q tests/test_canary_decks.py
```

This verifies representative end-to-end behavior:
- mixed YAML deck generation
- waterfall chart-cell generation with connectors
- multi-chart JSON deck generation
- chart template copy roundtrip (`chart_template_copy`)

## 3) CI gate

Ensure GitHub Actions `ci` is green:
- `test (3.9)`
- `test (3.12)`
- `canary`
- `lint`

## 4) Manual openability check

Generate at least one deck from YAML and one from JSON charts, then open in PowerPoint.

Confirm:
- no repair dialog
- expected chart counts and labels
- expected connector overlays on waterfall canaries

## 5) Docs sanity

Verify docs are consistent with shipped behavior:
- `docs/INPUT-SCHEMA.md`
- `docs/CHART-CELLS.md`
- `docs/CHARTS.md`
