# Input Schema (v1 – Text Tables + Icons)

This schema defines table slides with text content and optional icon indicators.

Key goals:
- Clean table structure (rows/cols + optional headers)
- Content fits within the layout bounds
- Consistent bullet/indent behavior via a single lstStyle

---

## Minimal Example

```yaml
title: "Pain Points"         # Required. Fills the title placeholder (ph0)
subtitle: "Key findings"     # Optional. Fills the subtitle placeholder (ph1)
tracker: "Section 3"         # Optional. Fills the on-page tracker placeholder

# Slide layout (template slide master) vs content-area layout are separate concerns.
slide_layout: Default        # template layout name (e.g. "Default", "2/3")
content_layout: default      # default | content | body | full | 1/1
                             # default/content/body start below header divider (safe with title/subtitle)
                             # full/1/1 start near top (tracker zone)

table:
  # rows/cols describe the *rendered grid*, including enabled header rows/cols
  # - rows includes the column header row when has_col_header=true
  # - cols includes the row header column when has_row_header=true
  rows: 6
  cols: 4
  has_col_header: true
  has_row_header: false

  col_headers:
    - "Category"
    - "Details"
    - "Owner"
    - "Next step"

  # cells is body-only: (rows - header_rows) × (cols - header_cols)
  cells:
    - ["Alpha", "Two sentences of body text here.", "Team A", "Define scope"]
    - ["Beta", "Another short description.", "Team B", "Prototype"]
    - ["Gamma", "Placeholder content used here.", "Team C", "Review"]
    - ["Delta", "Short summary.", "Team D", "Decide"]
    - ["Epsilon", "Wrap up.", "Team E", "Finalize"]
```

---

## Full Schema

```yaml
title: "Slide Title"           # Required. Fills title placeholder (ph0)
subtitle: "Subtitle text"      # Optional. Fills subtitle placeholder (ph1)
tracker: "Section name"        # Optional. Fills on-page tracker. Supports \n for line breaks.

slide_layout: Default          # Optional. Template layout name (e.g. "Default", "2/3")
content_layout: default        # Optional. Content-area policy
                              # default|content|body = start below header divider (safe with title/subtitle)
                              # full|1/1 = start near top (tracker zone)

content_area:                  # Optional override for layout bounds (EMU)
  x: 561090
  y: 1710000
  width: 11076170
  height: 4740000

table:
  rows: 6                       # Required (total rows, including header rows)
  cols: 5                       # Required (total cols, including row header col)

  has_col_header: true          # Optional (default true)
  has_row_header: true          # Optional (default false)

  row_header_col_header: "Row"  # Optional label for the row-header column

  col_headers:                  # Optional, used if has_col_header=true (body cols only)
    - "Header 1"
    - { text: "Header 2", sub: "(units)" }  # optional second line, non-bold
    - "Header 3"
    - "Header 4"

  row_headers:                  # Optional. If omitted and has_row_header=true,
    - "Row A"                   # the first column of cells is auto-extracted as row headers.
    - "Row B"                   # When provided explicitly, cells should NOT include the
    - "Row C"                   # row-header column data.
    - "Row D"
    - "Row E"

  cells:                        # (rows - header_rows) × (cols - header_cols)
    - ["A1", "B1", "C1", "D1"] # If row_headers omitted: first element per row becomes the row header
    - ["A2", "B2", "C2", "D2"]

  column_widths: equal            # Optional. Controls column sizing:
                                   #   (omit)       → auto: content-aware, wider columns for cells that need more room
                                   #   equal        → equal body-column widths (best when columns are items being compared)
                                   #   [1, 2, 1, 1] → manual proportions

  fonts:                        # Optional font sizing overrides
    body: 16                    # Default comes from template-config (typically 16pt)
    header: 16                  # Header >= body (warn if smaller)
    min: 10                     # Min size for fit checks
    max: 16                     # Max size for fit checks

  body_default_lvl: 0           # Default bullet level for body text (0 = no bullet)
  parse_bullets: true           # Parse "-" lines into bullet levels

  padding:                      # Optional cell padding (EMU)
    top: 45720
    bottom: 45720

  placeholders: true            # Fill missing cells/headers with placeholders
```

---

## Cell Content Formats

Each cell can be:

1. **String** (single paragraph)
2. **List of strings** (multiple paragraphs)
3. **List of paragraph objects** with per-paragraph formatting

Paragraph object fields:

```yaml
- text: "Paragraph text"
  sub: "(optional second line)"  # optional subtitle/unit line in default body color
  lvl: 0            # 0 = no bullet, 1 = bullet, 2 = nested
  size: 14          # font size in pt
  color: accent1    # theme name (tx1, accent1, dk2, ...) or hex (#RRGGBB / RRGGBB)
  font: Arial       # optional font override
  bold: true
  italic: false
  underline: false
```

`sub` behavior:
- rendered on a new line (line break inside the same paragraph block)
- keeps `lvl`/`font`/`size`
- forced non-bold
- uses default body color (`tx1`)

Example with bullets and overrides:

```yaml
cells:
  -
    -
      - { text: "Intro paragraph", size: 14, color: accent1, lvl: 0 }
      - { text: "We will:", lvl: 0 }
      - { text: "First bullet", lvl: 1 }
      - { text: "Nested bullet", lvl: 2 }
```

### Inline Formatting

Within any text string (cells, sidebar, headers), you can use markdown-style inline formatting:

| Syntax | Result |
|--------|--------|
| `**bold**` | **bold** text |
| `*italic*` | *italic* text |
| `[link text](https://example.com)` | clickable hyperlink |

These can be mixed with plain text in the same string:

```yaml
cells:
  - ["Visit **biodiversity.ru** for details"]
  - ["Read the *definitive* article on conservation"]
  - ["[Donate here](https://donate.example.org) to support the project"]
```

Hyperlinks render as blue underlined text and are clickable in PowerPoint presentation mode.

## Icons

Small colored indicators (filled circles) placed in cells instead of text. See `ICONS.md` for the full reference, standard presets, and design rationale.

### Quick syntax

```yaml
# Use a preset
icons:
  preset: severity          # severity | rag | confidence | priority

# Or define custom values
icons:
  size: 0.25               # diameter in inches, default 0.25
  values:
    Critical: "#E5546C"
    High: "#FAA082"
    Medium: "#E8BDAD"
    Low: "accent5"          # scheme colors also supported
  legend: true              # default true; set false to suppress
```

### Cell syntax

```yaml
cells:
  - - "Some text"
    - { icon: "High" }      # renders colored circle, not text
    - { icon: "Critical" }
    - "More text"
```

---

## Validation Rules (v1)

Errors:
- `table.rows` and `table.cols` are required
- `content_area` must include `x`, `y`, `width`, `height` if provided

Warnings:
- Missing `has_col_header` / `has_row_header` (defaults applied)
- Headers missing when enabled (placeholders will be generated)
- `column_widths` list length mismatch (body cols, or total cols when `has_row_header: true`)
- `col_headers` / `row_headers` length mismatch
- Unknown `layout` value (v1 supports: default, content, body, full, 1/1)

---

## Sidebar Content

Split layouts (2/3, 3/4, 1/2) have a secondary content area on the right. Use `sidebar` to fill it with formatted paragraphs:

```yaml
slide_layout: "2/3"

sidebar:
  - { text: "Current understanding", bold: true, color: accent2, size: 18 }
  - "Structure and ownership management is most complex at the time of events."
  - { text: "Questions for discussion", bold: true, color: accent2, size: 18 }
  - "What specific events are particularly difficult to manage?"
```

Each entry uses the same format as cell content: plain strings, paragraph objects with `text`/`bold`/`color`/`size`, or bullet syntax (`- item`). Run `pptx layouts <template>` to see which layouts have a secondary content area.

---

## Row Groups (Superheaders)

Use `row_groups` instead of `rows` + `row_headers` to create category-grouped tables with bold superheader rows spanning the full width. Each group has a `header` label and a list of `rows` beneath it. `header` can be either a string or an object like `{ text: "Group", sub: "(units)" }`.

When using `row_groups`, omit `rows` and `row_headers` — the row count and row-header column are derived from the groups. You still need `cols` (total columns including the superheader column).

**Convenience for chart tables**: if a group has an empty `header`, exactly one row, and that row contains a chart ref (e.g. `wf-1`), the first body cell is automatically promoted to the group header. The promoted header spans the row-header column + first body column, avoiding blank superheader bands for singleton start/total rows.

### Minimal Example

```yaml
title: "Risk Assessment"

table:
  cols: 3
  has_col_header: true
  col_headers: ["Risk", "Mitigation"]

  row_groups:
    - header: "Technical"
      rows:
        - ["API breaks on upgrade", "Pin versions, integration tests"]
        - ["Data loss during migration", "Backup + dry-run procedure"]

    - header: "Commercial"
      rows:
        - ["Customer churn", "Quarterly reviews, SLA guarantees"]
        - ["Pricing pressure", "Value-based pricing, cost transparency"]
```

This produces a table with:
- A column-header row: `[superheader col] | Risk | Mitigation`
- Two superheader rows ("Technical", "Commercial") each spanning the full width
- Two body rows beneath each superheader

### Full Syntax

```yaml
table:
  cols: 3                           # Total columns (including the row-header/superheader column)
  has_col_header: true
  col_headers: ["Col A", "Col B"]   # Body column headers only (excluding superheader column)
  row_header_col_header: "Category" # Optional label for the superheader column in the header row

  # Color overrides for superheader styling
  row_superheader_color: accent2    # Color for superheader text (theme name or hex)
  col_superheader_color: accent2    # Color for column superheader text (if using col_superheaders)

  row_groups:
    - header: "Group A"             # Superheader label (bold, larger font, colored)
      rows:                         # Body rows in this group
        - ["Cell 1", "Cell 2"]      # Each row has (cols - 1) cells (body columns only)
        - ["Cell 3", "Cell 4"]

    - header: "Group B"
      rows:
        - ["Cell 5", "Cell 6"]
```

### Row Groups with Rich Cell Content

Each cell within a row can use the same formats as regular cells: strings, lists of strings, or paragraph objects:

```yaml
row_groups:
  - header: "Deploy"
    rows:
      -
        -
          - { text: "Co-locate in EU data centers", bold: true, lvl: 0 }
        -
          - { text: "Racks in proven sites with network in place", lvl: 1 }
          - { text: "Monthly colocation: ~€1.9k per server", lvl: 2 }
```

### Column Superheaders

Separate from row groups, `col_superheaders` add a header row above the column headers that spans multiple columns:

```yaml
table:
  cols: 4
  has_col_header: true

  col_superheaders:
    - label: ""                  # Empty label above the row-header column
      span: 1
    - label: "Financial Details" # Spans 3 body columns
      span: 3

  col_headers: ["Revenue", "Cost", "Margin"]
  # ...
```

Column superheaders and row groups can be combined in the same table.

---

## Legacy Schema

The old schema (moons, superheaders, bullets, etc.) is preserved at:

```
legacy/INPUT-SCHEMA.md
```
