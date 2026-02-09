# Charts (JSON) — alpha

> **Alpha**: the chart generator works but the JSON schema and CLI flags may change.

The `pptx charts` command generates bar/stacked/waterfall charts from JSON specs using
python-pptx plus optional overlay labels.

## Usage

```bash
pptx charts spec.json output.pptx
```

Optional template flags:

```bash
pptx charts spec.json output.pptx \
  --template /path/to/template.pptx \
  --layout "Default" \
  --expected-template clean-slides
```

If the chart generator module is not found automatically, pass `--module-path` or set
`CLEAN_SLIDES_CHARTS_PATH` to the generator script path.

## Base schema

```json
{
  "title": "Chart title",
  "type": "clustered", // "stacked" or "waterfall"
  "categories": ["A", "B", "C"],
  "series": [
    {"name": "BU1", "values": [1, 2, 3], "color": "#4472C4"},
    {"name": "BU2", "values": [4, 5, 6], "color": "#ED7D31"}
  ],
  "show_data_labels": true,
  "add_overlay_labels": true
}
```

## Waterfall config

```json
"waterfall": {
  "decrease_categories": ["Costs"],
  "total_categories": ["Net"],
  "total_series": ["Totals"],
  "range_series": ["Range"],
  "reuse_start_base": true,
  "label_gap": 25600,
  "connector_style": "gap",
  "connector_value": "totals",
  "connector_overlap": 6000,
  "connector_inset": 10000,
  "total_override": false
}
```

Notes:
- `range_series` marks series that render visually but do **not** affect running totals.
- Set `connector_value: "tops"` for legacy connector anchoring.
- `total_override: true` restores legacy behavior where any series value in a total category overrides
  computed totals.

## Horizontal charts

```json
"bar": { "orientation": "horizontal" }
"waterfall": { "orientation": "horizontal" }
```

## Multi‑chart decks

Use a top-level `charts` list to generate multiple charts in one deck:

```json
{
  "template": "template.pptx",
  "layout": "Default",
  "expected_template": "clean-slides",
  "charts": [
    {"type": "clustered", "categories": ["A"], "series": [{"name": "S1", "values": [1]}]},
    {"type": "waterfall", "categories": ["Start", "End"], "series": [{"name": "S1", "values": [10, 10]}]}
  ]
}
```
