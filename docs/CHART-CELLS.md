# Chart Cells — Design Document

> **Status**: Proposed. Not yet implemented.

## Summary

Chart cells embed native PowerPoint chart shapes inside table cells. Instead of
rendering text, the cell contains a bar whose size is proportional to a numeric
value. The table structure becomes the chart's scaffold — row headers label the
categories, column headers label the series, and dividers provide the grid lines.

This extends clean-slides' core philosophy: constrained structure produces
consistent output. Rows with chart cells get equal height. Columns with chart
cells get equal width. The tool enforces the constraint; the author focuses on
the data.

---

## Motivating example

A leverage sensitivity table where some columns show text and others show bars:

```
┌────────────┬──────────┬──────────────────┬──────────────────┬─────────┐
│            │ Interest │ ND / EBITDAaL    │ UFCF / Interest  │ GIP IRR │
│ Net debt   │ paid     ├────────┬─────────┼────────┬─────────┤ %       │
│ €m         │ €m       │ FY26E  │  FY33E  │ FY26E  │  FY33E  │         │
├────────────┼──────────┼────────┴─────────┼────────┴─────────┼─────────┤
│ €3,771m    │ ██ €156m │  ██      █       │  ██       ████   │  8.3%   │
│ €5,000m    │ ███€206m │  ███     ██      │  ██       ███    │  8.9%   │
│ €6,000m    │ ████248m │  ████    ███     │  █        ██     │  9.6%   │
│ €7,000m    │ █████289 │  █████   ████    │  █        ██     │ 10.4%   │
│ €8,000m    │ ██████330│  ██████  █████   │  █        █      │ 11.5%   │
└────────────┴──────────┴──────────────────┴──────────────────┴─────────┘
```

Two chart patterns on this slide:

1. **Column chart** — "Interest paid" column: one chart shape spanning all 5 rows.
   Each row gets one horizontal bar. The chart is a single native chart object.

2. **Row charts** — ratio columns: one chart shape per row, each spanning 4 columns.
   Each chart has 4 vertical bars (FY26E and FY33E for two metrics).

---

## YAML schema

### Chart definitions

Charts are defined at the top level of the YAML spec, outside the table. Each
chart has a name, direction, data values, and optional formatting.

```yaml
charts:
  bar1:
    dir: horizontal            # horizontal | vertical
    values: [156, 206, 248, 289, 330]
    format: "€{}m"             # label format (Python str.format style)
    color: accent1             # optional: theme color or hex

  bar2: &ratios
    dir: vertical
    values: [4.0, 3.1, 4.7, 7.3]
    format: "{}x"
    color: accent1

  bar3:
    <<: *ratios
    values: [5.2, 4.1, 3.5, 5.5]

  bar4:
    <<: *ratios
    values: [6.3, 5.0, 3.0, 4.6]

  bar5:
    <<: *ratios
    values: [7.3, 5.8, 2.5, 4.0]

  bar6:
    <<: *ratios
    values: [8.4, 6.6, 2.2, 3.5]
```

YAML anchors (`&ratios` / `<<: *ratios`) reduce repetition when multiple charts
share the same type, direction, and formatting — only `values` differs.

### Chart properties

| Key        | Type     | Required | Description |
|------------|----------|----------|-------------|
| `dir`      | string   | yes      | `horizontal` or `vertical` |
| `values`   | list     | yes      | Numeric values, one per bar |
| `format`   | string   | no       | Label format string. `{}` is replaced by the value. Default: `"{}"` |
| `color`    | string   | no       | Bar fill color. Theme name (`accent1`, `dk2`) or hex (`#4472C4`). Default: template accent. |
| `label_position` | string | no  | `above` (default for vertical), `right` (default for horizontal), `on`, `none` |
| `scale_max`| number   | no       | Override automatic scale maximum |

### Cell references

Cells reference charts using the pattern `chartname-N` where `N` is the 1-based
index into the chart's `values` list.

```yaml
table:
  cells:
    - [bar1-1, bar2-1, bar2-2, bar2-3, bar2-4, "8.3%"]
    - [bar1-2, bar3-1, bar3-2, bar3-3, bar3-4, "8.9%"]
    - [bar1-3, bar4-1, bar4-2, bar4-3, bar4-4, "9.6%"]
    - [bar1-4, bar5-1, bar5-2, bar5-3, bar5-4, "10.4%"]
    - [bar1-5, bar6-1, bar6-2, bar6-3, bar6-4, "11.5%"]
```

**Merge rule**: adjacent cells referencing the same chart name (e.g. `bar1-1`,
`bar1-2`, `bar1-3` in a column, or `bar2-1`, `bar2-2`, `bar2-3`, `bar2-4` in a
row) are merged into a single native chart shape spanning those cells. The `-N`
suffix determines which bar maps to which position in the chart.

**Direction determines merge axis**:
- `dir: horizontal` → bars stack vertically (one per row). References should
  appear in the same **column** across consecutive rows.
- `dir: vertical` → bars sit side-by-side (one per column). References should
  appear in the same **row** across consecutive columns.

---

## Layout constraints

Chart cells impose two layout constraints beyond normal text cells:

1. **Equal row heights** — rows containing horizontal-bar chart cells get equal
   height, so bars are visually comparable.

2. **Equal column widths** — columns containing vertical-bar chart cells get
   equal width, so bars are visually comparable.

These constraints are applied during the sizing pass, alongside existing
`column_widths: equal` logic. Text-only rows and columns are unaffected.

---

## Scale sharing

Bars across cells must share a common axis scale so their lengths/heights are
comparable. Scale groups are determined by chart identity:

- All cells referencing the same chart name share one scale (the max of that
  chart's values).
- Separate chart names have independent scales, even if adjacent.
- `scale_max` on a chart definition overrides the automatic maximum (useful when
  you want two charts on the same scale without merging them).

For the motivating example:
- `bar1` (Interest paid): scale max = 330
- `bar2`–`bar6` (ratios): each has its own scale. To share a scale across all
  ratio charts, set the same `scale_max` on each (or use the YAML anchor).

---

## Superheader interaction

### Single-span columns merge upward

When a table has `col_superheaders`, columns not covered by any multi-column
superheader get their `col_header` merged vertically across both the superheader
row and the header row. This is automatic — no flag needed.

```
Before (wasted space):             After (merged):
┌──────┬──────┬─────────────┐     ┌──────────────┬─────────────┐
│      │      │ ND/EBITDAaL │     │              │ ND/EBITDAaL │
├──────┼──────┼──────┬──────┤     │  Net debt    ├──────┬──────┤
│Net   │Int.  │FY26E │FY33E │     │  €m          │FY26E │FY33E │
│debt  │paid  │      │      │     │              │      │      │
├──────┼──────┼──────┼──────┤     ├──────────────┼──────┼──────┤
```

**Rule**: if a column's superheader span is 1 (or the column is not covered by
any superheader), the col_header cell spans both header rows.

---

## Native chart rendering

Each merged chart region renders as a python-pptx `chart` shape:

- **Chart type**: `COLUMN_CLUSTERED` for vertical, `BAR_CLUSTERED` for horizontal
- **No chrome**: axes hidden, gridlines off, legend off, plot area fills the shape
- **Data labels**: positioned per `label_position`, formatted per `format`
- **Fill**: solid color from `color` property (resolved against template theme)
- **Gap width**: tight (e.g. 50%) to maximise bar area within cell bounds

The chart is sized to the bounding box of the merged cells, minus cell padding.

---

## Validation

All validation happens at parse time, before rendering. Errors reference the
cell position (row, column) and the problematic reference.

| Check | Error message |
|-------|---------------|
| Unknown chart name | `Cell (2, 3): chart "bar7" is not defined in charts:` |
| Index out of range | `Cell (1, 1): bar1-6 but bar1 has only 5 values` |
| Index gap | `bar2 has indices [1, 2, 4] — missing index 3` |
| Duplicate index | `bar1-2 appears in cells (2, 1) and (3, 1)` |
| Direction mismatch | `bar1 (horizontal) has refs in the same row — expected same column` |
| Non-contiguous | `bar2 refs in row 1 at columns [2, 4] — column 3 is missing` |
| Mixed content | `Cell (1, 1) has chart ref "bar1-1" but also text content` |

`pptx validate` catches these alongside existing schema checks.

---

## Full example — leverage sensitivity

```yaml
title: "Leverage"
subtitle: "Significant further debt capacity required to reach target IRR"
slide_layout: "2/3"

sidebar:
  - { text: "Illustrative analysis assumptions", bold: true, color: accent2 }
  - "4.1% blended annual cash interest (fixed)"
  - "No amortisation"
  - "FY33 exit without refinancing"
  - "Base case operating assumptions"
  - ""
  - { text: "Commentary", bold: true, color: accent2 }
  - "Likely room for c.2.0x additional EBITDAaL leverage given healthy ratios"
  - "To refine and test lender appetite"

charts:
  interest:
    dir: horizontal
    values: [156, 206, 248, 289, 330]
    format: "€{}m"
    color: accent1

  r1: &ratios
    dir: vertical
    format: "{}x"
    color: accent1
    scale_max: 9.0
    values: [4.0, 3.1, 4.7, 7.3]

  r2:
    <<: *ratios
    values: [5.2, 4.1, 3.5, 5.5]

  r3:
    <<: *ratios
    values: [6.3, 5.0, 3.0, 4.6]

  r4:
    <<: *ratios
    values: [7.3, 5.8, 2.5, 4.0]

  r5:
    <<: *ratios
    values: [8.4, 6.6, 2.2, 3.5]

table:
  rows: 5
  cols: 7
  has_row_header: true
  has_col_header: true

  col_headers:
    - "Net debt\n€m"
    - "Interest paid\n€m"
    - "FY26E"
    - "FY33E"
    - "FY26E"
    - "FY33E"
    - "GIP IRR\n%"

  col_superheaders:
    - header: "Net Debt / EBITDAaL\nRatio"
      span: [3, 4]
    - header: "UFCF / Interest\nRatio"
      span: [5, 6]

  row_headers:
    - ["€3,771m", { text: "FY26E consensus", bold: false }]
    - "€5,000m"
    - ["€6,000m", { text: "Base Case", bold: false }]
    - "€7,000m"
    - "€8,000m"

  cells:
    - [interest-1, r1-1, r1-2, r1-3, r1-4, "8.3%"]
    - [interest-2, r2-1, r2-2, r2-3, r2-4, "8.9%"]
    - [interest-3, r3-1, r3-2, r3-3, r3-4, "9.6%"]
    - [interest-4, r4-1, r4-2, r4-3, r4-4, "10.4%"]
    - [interest-5, r5-1, r5-2, r5-3, r5-4, "11.5%"]
```

---

## Implementation plan

### Phase 1: Parsing and validation

1. **`spec.py`** — Parse `charts:` top-level key into `ChartDef` dataclass.
   Detect `chartname-N` references in cells during `TableSpec.from_dict()`.
   Store as `ChartRef(name, index)` in the cell grid.

2. **`placeholder.py`** — Skip chart-ref cells when filling placeholders.

3. **`validate.py`** — Add all checks from the validation table above.

### Phase 2: Sizing

4. **`sizing.py`** — When chart refs are present in a column, that column
   participates in equal-width sizing. When chart refs span rows, those rows
   participate in equal-height sizing.

5. **`measure.py`** — Chart cells have zero text width (no text to measure).
   Min width comes from the label format string at the configured font size.

### Phase 3: Rendering

6. **`renderer.py`** — After placing all text boxes, iterate chart groups.
   For each group of merged chart refs:
   - Compute bounding box from the cell positions
   - Create a python-pptx chart shape (`slide.shapes.add_chart`)
   - Configure: hide axes, set gap width, add data labels, apply fill color
   - Position and size to the bounding box

### Phase 4: Integration

7. **`cli.py`** — No changes needed; `pptx generate` and `pptx validate`
   pick up chart cells automatically through the existing pipeline.

8. **Tests** — Unit tests for parsing, validation, sizing, and rendering.
   Integration test generating a full chart-cell slide and verifying shape
   count and positions.

---

## Out of scope (for now)

- **Waterfall charts in cells** — waterfall bar positions must align with row
  boundaries, requiring coordination between the chart's internal layout and
  the table's row heights. Complex; defer to a later phase.

- **Stacked bars in cells** — each cell currently maps to one bar. Stacked
  bars (multiple values per cell) would need a different cell reference syntax.

- **Chart-only slides** — this design is for charts embedded in tables. The
  existing `pptx charts` command handles full-slide charts from JSON.
