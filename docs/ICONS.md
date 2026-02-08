# Icons & Indicators

Small visual markers placed in table cells instead of text. Used to convey categorical status at a glance — severity, progress, sentiment, priority.

---

## Principles

1. **One shape per icon** — simple filled oval (`prstGeom prst="ellipse"`), no outlines, no groups. The think-cell MoonLegend5 approach (group of ellipse + arc) is unnecessary when the indicator is always fully filled.
2. **Centered in cell** — positioned at `(cell_x + cell_w/2 - r, cell_y + cell_h/2 - r)` where `r = size/2`.
3. **Fixed size** — `0.25"` (228,600 EMU) diameter by default. Small enough to fit in compact columns, large enough to read.
4. **Colors from the design system** — uses the semantic palette defined in `DESIGN-SYSTEM.md §2.3`. Supports both `srgbClr` hex (`"#E5546C"`) and `schemeClr` names (`"accent5"`).
5. **Legend when icons are present** — auto-rendered in the top-right of the slide, one entry per distinct icon value used, vertically stacked. Each entry is a small circle + label.
6. **Icon columns stay compact** — weight `0.0` in column sizing (no extra space absorption). Min width = icon size + padding.

---

## Standard Icon Sets

### Severity (traffic lights)

The default for risk matrices, due diligence assessments, and any "how bad is this" dimension. Warm tones: worse = redder.

| Value | Color | Hex | Source | Visual |
|-------|-------|-----|--------|--------|
| **Critical** | Red | `#E5546C` | srgbClr | 🔴 |
| **High** | Salmon | `#FAA082` | srgbClr | 🟠 |
| **Medium** | Tan | `#E8BDAD` | srgbClr | 🟡 |
| **Low** | Slate light | `#CBD5E1` | accent5 | ⚪ |

**When to use**: Risk likelihood, risk severity, impact magnitude, vulnerability level. Any scale measuring the size of a *negative* outcome.

**YAML preset name**: `severity`

```yaml
icons:
  preset: severity
```

Equivalent to:
```yaml
icons:
  size: 0.25
  values:
    Critical: "#E5546C"
    High: "#FAA082"
    Medium: "#E8BDAD"
    Low: "accent5"
```

### Status (RAG)

Classic Red/Amber/Green for project tracking and operational health. Includes a neutral grey for items not yet assessed.

| Value | Color | Hex | Source | Visual |
|-------|-------|-----|--------|--------|
| **Red** | Red | `#E5546C` | srgbClr | 🔴 |
| **Amber** | Amber | `#FAA082` | srgbClr | 🟠 |
| **Green** | Green | `#00B050` | srgbClr | 🟢 |
| **Grey** | Slate light | `#CBD5E1` | accent5 | ⚪ |

**When to use**: Project status, workstream health, milestone tracking. Binary-ish assessments of "is this on track?"

**YAML preset name**: `rag`

```yaml
icons:
  preset: rag
```

### Confidence

For expressing conviction or certainty levels. Cool tones: more certain = darker blue.

| Value | Color | Hex | Source | Visual |
|-------|-------|-----|--------|--------|
| **Strong** | Midnight | `#0D193B` | tx1 | 🔵 |
| **Moderate** | Electric blue | `#2251FF` | srgbClr | 🔵 |
| **Weak** | Cyan | `#00A9F4` | srgbClr | 🔵 |
| **None** | Slate light | `#CBD5E1` | accent5 | ⚪ |

**When to use**: Confidence in assumptions, strength of evidence, data quality assessment.

**YAML preset name**: `confidence`

### Priority

For urgency or importance ranking. Uses the same warm palette as severity but with different labels.

| Value | Color | Hex | Source | Visual |
|-------|-------|-----|--------|--------|
| **P1** | Red | `#E5546C` | srgbClr | 🔴 |
| **P2** | Salmon | `#FAA082` | srgbClr | 🟠 |
| **P3** | Tan | `#E8BDAD` | srgbClr | 🟡 |
| **P4** | Slate light | `#CBD5E1` | accent5 | ⚪ |

**When to use**: Task priority, action item urgency, backlog ordering.

**YAML preset name**: `priority`

---

## Custom Icon Sets

Define a fully custom set by providing `values` directly:

```yaml
icons:
  size: 0.20              # optional, default 0.25"
  values:
    Excellent: "#00B050"
    Good: "#92D050"
    Fair: "#FAA082"
    Poor: "#E5546C"
```

Colors can be:
- **Hex RGB**: `"#E5546C"` → produces `<a:srgbClr val="E5546C"/>`
- **Scheme name**: `"accent5"` → produces `<a:schemeClr val="accent5"/>` (adapts if theme changes)

---

## YAML Cell Syntax

Reference icon values by name in cell data:

```yaml
cells:
  - - "MSAs incl. 2028 renewals"
    - { icon: "High" }        # Likelihood column
    - { icon: "Critical" }    # Severity column
    - "Initial perspective bullet text..."
```

An `{ icon: "X" }` cell:
- Produces an oval shape, not a text box
- Is centered in the cell grid position
- Color is looked up from the `icons.values` mapping
- The column gets weight `0.0` (stays compact)

**Empty icon**: Use `{ icon: "" }` or `""` for no indicator (cell left blank).

---

## Legend

When `icons` is defined, a legend is automatically rendered.

### Placement
- **Position**: Top-right of the slide, below the tracker placeholder
- **Default anchor**: `x = slide_width - margin - legend_width`, `y = title_area_bottom` (approximately 0.45" from top)
- **Grows downward**: One row per icon value

### Structure
Each legend entry is:
- A small circle (same `size` as table icons)
- A text label to the right (Arial 10pt, tx1)
- Vertical spacing: `size + 4pt` between entries

### Override
```yaml
icons:
  preset: severity
  legend: false          # suppress legend
  legend_position:       # or override position (EMU)
    x: 10987387
    y: 408583
```

### Ordering
Legend entries appear in the order defined in `values` (top = most severe / worst). This matches the visual convention of "red at top."

---

## Rendering Details

### Shape XML (simplified)

Each icon produces a single `<p:sp>` on the slide:

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="..." name="Icon_R2_C3"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{center_x}" y="{center_y}"/>
      <a:ext cx="228600" cy="228600"/>
    </a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:solidFill>
      <a:srgbClr val="E5546C"/>   <!-- or schemeClr -->
    </a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    <a:p><a:endParaRPr lang="en-US"/></a:p>
  </p:txBody>
</p:sp>
```

### Positioning Algorithm

```
cell = layout.cells[row][col]  →  (x, y, w, h)
icon_size = 228600  # 0.25"

icon_x = x + (w - icon_size) // 2
icon_y = y + (h - icon_size) // 2
```

### Interaction with Sizing

- Icon columns contain no text → `_cell_text_length()` returns 0 → weight `0.0`
- Min column width = `icon_size + 2 * base_pad` (enough for the circle + breathing room)
- Icon cells are skipped in `_body_row_required()` (no text to measure)
- Row height is driven by sibling text columns, not by icon columns

---

## Examples

### Risk table (Tiber slide 4)

```yaml
icons:
  preset: severity

table:
  col_headers: ["Risk from...", "Likelihood", "Severity", "Initial perspective"]
  row_groups:
    - header: "Commercial & delivery"
      rows:
        - - "MSAs incl. 2028 renewals"
          - { icon: "High" }
          - { icon: "Critical" }
          - "All or nothing clause may protect..."
        - - "OLO/POP new bus. and rolling renewals"
          - { icon: "Medium" }
          - { icon: "Low" }
          - "6+6/9+9 contracts..."
```

### Project status tracker

```yaml
icons:
  preset: rag

table:
  col_headers: ["Workstream", "Status", "Owner", "Next milestone"]
  cells:
    - - "Legal DD"
      - { icon: "Green" }
      - "External counsel"
      - "Final report due 15 Feb"
    - - "Financial model"
      - { icon: "Amber" }
      - "Deal team"
      - "Sensitivity cases pending"
    - - "Management meetings"
      - { icon: "Red" }
      - "MD sponsor"
      - "Rescheduling — key person unavailable"
```

---

## Design Rationale

**Why simple ovals, not moons?**
The reference file uses think-cell's MoonLegend5 (a group shape: ellipse background + arc overlay). This supports partial fills (25%, 50%, 75%). But in practice, the indicators are always fully filled — the arc sweep covers 100%. A single filled ellipse is visually identical, produces cleaner XML, and doesn't depend on think-cell. If partial fills are ever needed, the moon approach can be added as a separate icon type.

**Why warm tones for severity?**
Severity measures negative impact magnitude. The warm palette (red → salmon → tan → grey) maps intuitively: redder = worse. This avoids the ambiguity of green in a negative-only scale (green risk? green severity?). Grey/slate for "Low" signals minimal concern without implying positive.

**Why fixed size?**
Variable-size icons would complicate column sizing and create visual noise. A fixed 0.25" circle is:
- Large enough to distinguish colors at presentation scale
- Small enough to fit in a compact column (~0.65" with padding)
- Consistent across all rows, reinforcing the grid structure

**Why auto-legend?**
Icons without a legend are ambiguous. The legend is small (≈0.5" × 0.6"), fits in the top-right without overlapping content, and uses the same visual language as the table. It can be suppressed with `legend: false` for slides where the context is already established (e.g., second slide in a series).
