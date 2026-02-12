# Charts (JSON)

The `pptx charts` command generates native PowerPoint charts from JSON specs.

Supported chart families:
- clustered bars/columns
- stacked bars/columns
- waterfall (overlay-driven)

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

The chart engine is bundled with `clean-slides`; no external module path is required.

## Base schema

```json
{
  "title": "Chart title",
  "type": "clustered", // "stacked" or "waterfall"
  "categories": ["A", "B", "C"],
  "series": [
    {"name": "BU1", "values": [1, 2, 3], "color": "#4472C4"},
    {"name": "BU2", "values": [4, 5, 6], "color": "accent2"}
  ],
  "show_data_labels": true,
  "add_overlay_labels": true
}
```

## Bar options

```json
"bar": {
  "orientation": "horizontal",
  "chart_template": "templates/chart-style.pptx",
  "chart_template_copy": true,
  "chart_template_slide": 1,
  "chart_template_chart_index": 0
}
```

Notes:
- `orientation` can be `horizontal` or `vertical`.
- `chart_template_copy: true` applies an OPC-level chart XML/relationship replacement,
  preserving template internals.
- Template paths are resolved from the spec base directory when relative.

## Waterfall options

```json
"waterfall": {
  "orientation": "horizontal",
  "decrease_categories": ["Costs"],
  "total_categories": ["Net"],
  "total_series": ["Totals"],
  "range_series": ["Range"],
  "reuse_start_base": true,
  "label_gap": 25600,
  "connector_style": "gap",        // "gap" | "step"
  "connector_dash_style": "long_dash", // "solid" | "long_dash" | "dot"
  "connector_value": "totals",     // "totals" | "tops"
  "connector_overlap": 6000,
  "connector_inset": 10000,
  "total_override": false
}
```

Notes:
- `range_series` renders visually but does **not** affect running totals.
- `connector_value: "tops"` restores legacy connector anchoring.
- `total_override: true` restores legacy behavior where explicit values in total
  categories override computed totals.

## Multi-chart decks

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
