---
name: clean-slides
description: Inspect, edit, and generate PowerPoint slides. Progressive drill-down inspection, text editing with formatting preservation, table generation from YAML, and PNG rendering for verification.
---

# PowerPoint Skill

Use the `pptx` CLI for all PowerPoint operations.

## When to use
- **Inspect** slide contents, shapes, formatting, charts
- **Edit** text while preserving formatting
- **Generate** table slides from YAML specs
- **Render** slides to PNG for visual verification

## Setup

Before generating slides, ensure the project has a `.clean-slides/` directory. The CLI auto-discovers `template.pptx` and `config.yaml` from there — no flags needed.

```bash
# Quick start — uses the bundled example template
pptx init

# With a custom/corporate template
pptx init -t path/to/corporate-template.pptx
```

This creates `.clean-slides/template.pptx` and `.clean-slides/config.yaml`. All `pptx generate` commands in that directory tree will use them automatically.

If no `.clean-slides/` directory exists, the CLI prints a hint suggesting `pptx init`.

### Overriding auto-discovery

Explicit flags always win:

```bash
# Override template and config for a single run
pptx generate spec.yaml -t other.pptx -c other-config.yaml -o output.pptx
```

See `examples/custom-template/README.md` for a full walkthrough of custom template setup.

## Core workflow: inspect → edit → verify

```bash
# 1. Inspect — progressive drill-down
pptx show deck.pptx                    # slide list
pptx show deck.pptx 3                  # shapes on slide 3
pptx show deck.pptx 3 5                # full detail for shape [5]

# 2. Edit — shape identified by index or name
pptx edit deck.pptx 3 5 "New text" --out edited.pptx

# 3. Verify — render to PNG
pptx render edited.pptx 3 --out images/
```

## Inspect commands

```bash
pptx show <file>                       # slides with layout, shape count, word count
pptx show <file> <slide>               # all shapes sorted top→bottom, left→right
pptx show <file> <slide> <shape>       # full JSON: text runs, formatting, chart data

pptx layouts <file>                    # available slide layouts with placeholders & content areas
pptx theme <file>                      # color scheme (accent1, accent2, etc.)
pptx xml <file> <slide> <shape>        # raw OOXML for debugging
```

**Run `pptx layouts` on your template** before generating — it shows each layout's fillable placeholders (title, subtitle, tracker, logo, etc.) and content areas with dimensions. Layouts with a **secondary** content area (e.g. 2/3, 3/4) have a sidebar for supporting content like key takeaways, charts, or images.

**Shape types in output:** `ph0` (placeholder), `text`, `chart`, `image`, `group`, `line`, `shape`

## Edit commands

```bash
pptx edit <file> <slide> <shape> <text> [--out PATH]
pptx batch <file> <edits.json> [--out PATH]
```

**Text formats — inspect output = edit input:**

```json
{"paragraphs": [
  {"runs": [
    {"text": "Bold ", "bold": true},
    {"text": "normal"}
  ], "level": 0}
]}
```

**Run fields:** `text`, `font`, `size`, `bold`, `italic`, `underline`, `color`, `superscript`, `subscript`

- `"bold": true` → explicitly bold
- `"bold": false` → explicitly not bold (overrides inherited)
- (omit field) → inherit from parent

## Slide management

```bash
pptx add-slide <file> <layout> [--at N] [--out PATH]
pptx delete-slide <file> <slide> --confirm [--out PATH]
pptx delete-shape <file> <slide> <shape> [--out PATH]
pptx insert <deck.pptx> <source.pptx> [--at N] [--slides 1,3-5] [--out PATH]
```

### Replace a slide with a generated table

```bash
# 1. Generate the table slide (title/subtitle are set in the YAML)
pptx generate spec.yaml -t template.pptx -o /tmp/table.pptx

# 2. Insert at position 8, pushing old slide 8 → 9
pptx insert deck.pptx /tmp/table.pptx --at 8 --out deck.pptx

# 3. Delete the old slide (now at position 9)
pptx delete-slide deck.pptx 9 --confirm --out deck.pptx

# 4. Verify
pptx render deck.pptx 8 --out renders/
```

## Render

```bash
pptx render <file> <slides> [--out DIR] [--dpi N]
pptx crop <png> <L> <T> <R> <B> [--out PATH]
```

`<slides>` accepts single (`3`), list (`1,8,16`), or ranges (`1,3-5,8`). Multiple slides share one PDF conversion.

Coordinates for crop are in inches (matches shape positions from `show`).

## Generate tables from YAML

Context: `pptx generate` accepts **multiple YAML files** and produces **one deck** with one slide per YAML, in input order.

```bash
pptx generate spec.yaml -o output.pptx
pptx generate slide1.yaml slide2.yaml slide3.yaml -o output.pptx
pptx generate spec.yaml -t template.pptx -o output.pptx
pptx generate spec.yaml -t template.pptx -c config.yaml -o output.pptx
pptx validate spec.yaml
pptx verify spec.yaml --detail
```

**Top-level YAML keys** (all optional, set before `table:`):

- `title` — **required** — fills the title placeholder
- `subtitle` — fills the subtitle placeholder
- `tracker` — fills the on-page tracker (supports `\n` for line breaks)
- `slide_layout` — template layout name (e.g. `"Default"`, `"2/3"`, `"3/4"`)
- `content_layout` — content area policy: `default` (below header) or `full` (near top)
- `sidebar` — formatted paragraphs for the secondary content area in split layouts (2/3, 3/4, 1/2)

See `INPUT-SCHEMA.md` for the complete schema.

## Inline formatting

Text strings in cells, headers, and sidebar support markdown-style inline formatting:

- `**bold text**` → bold
- `*italic text*` → italic
- `[link text](https://example.com)` → clickable hyperlink (blue, underlined)

## Column widths

By default, columns are sized automatically based on content (wider for cells that need more wrapping room). Override with `column_widths`:

```yaml
table:
  column_widths: equal          # equal body columns — best when columns are items being compared
  column_widths: [1, 2, 1, 1]  # manual proportions
```

## Custom templates

The tool ships with a built-in template config, but supports **any** PowerPoint template. The workflow is: inspect the template, generate a config, review/adjust, then generate slides.

### Step 1: Inspect the template

```bash
pptx layouts template.pptx             # see all layouts, placeholders, content areas
pptx theme template.pptx               # see theme color scheme (dk1, lt1, accent1, etc.)
pptx show template.pptx 1              # inspect shapes on a sample slide
pptx xml template.pptx 1 <shape>       # raw XML for bullet lstStyle details
```

**What to look for:**
- **Layouts**: which layout names exist, which have content areas
- **Placeholders**: what index is title (usually 0), subtitle (usually 1), any tracker/breadcrumb
- **Theme colors**: what dk1, dk2, accent1 map to — these determine text and accent colors
- **Fonts**: what headline vs body fonts the template uses (check actual slide text, not just theme — templates often override theme fonts at the slide level)
- **Bullet hierarchy**: if the template has custom bullet styles, inspect the lstStyle XML on a content placeholder

### Step 2: Generate a starter config

```bash
pptx init-config template.pptx -o template-config.yaml
```

This auto-extracts:
- Theme colors → mapped to semantic names (midnight, electric_blue, etc.)
- Theme fonts → headline/body
- Slide dimensions and estimated layout geometry
- Placeholder indices (title, subtitle, tracker if detected)
- Available layout names (listed as comments)

The output is a **complete, valid config** with sensible defaults for everything — but some values will need adjustment.

### Step 3: Review and adjust the config

Key areas to verify and adjust:

1. **`colors`**: The auto-extracted `electric_blue` and `cyan` may not map correctly to your template's accent colors. Compare with `pptx theme` output and adjust.

2. **`fonts`**: The theme's major/minor fonts may differ from what the template actually uses on slides. Inspect a content slide (`pptx show template.pptx 2 <shape>`) and check the actual font in use.

3. **`font_sizes`**: Defaults are conservative (12pt body). Check your template's actual type scale and adjust title, body, table sizes.

4. **`placeholders`**: Verify `title` (usually 0) and `subtitle` (usually 1) are correct. If your template has a tracker/breadcrumb placeholder, set its index. Remove `tracker` if there isn't one.

5. **`default_colors`**: These control what theme color names are used for body text (`tx1`), headers (`tx1` or `accent1`), dividers, and links. Adjust to match your template's design language.

6. **`bullets`**: The hardest to auto-detect. If your template has custom bullet styles with specific chars (•, –, ›) and indentation, inspect the slide master lstStyle:
   ```bash
   pptx xml template.pptx 1 <content-placeholder>
   # Look for <a:lstStyle> with <a:lvl1pPr>, <a:lvl2pPr>, etc.
   ```
   Update bullet chars, `mar_l_emu`, `indent_emu`, and spacing to match.

7. **`layout`**: The auto-detected margins and Y positions are estimates. If slides look misaligned, measure positions from `pptx show template.pptx <slide>` output and adjust `content_start_y_emu`, `footer_line_y_emu`, etc.

### Step 4: Generate with custom config

```bash
# Generate
pptx generate spec.yaml -t template.pptx -c template-config.yaml -o output.pptx

# Validate / verify also accept --config
pptx validate spec.yaml -c template-config.yaml
pptx verify spec.yaml -c template-config.yaml --detail
```

### Tips for custom templates

- **Start simple**: generate a basic 2×2 table first and render it to check alignment
- **Iterate**: render after each config adjustment to see the effect
- **Content areas from layout**: when using `-t template.pptx`, the generator reads content area dimensions from the actual slide layout placeholders, so layout geometry often works even with approximate config values
- **Bullet levels**: if your template doesn't use custom bullets, the defaults (•, –, ›) work for most templates
- **The config doesn't need to be perfect**: the generator adapts to the template's actual placeholder positions at runtime; the config mainly controls styling (fonts, colors, spacing)

## Backup policy

Before editing, save a backup:
```bash
cp deck.pptx deck.backup.pptx
pptx edit deck.pptx ... --out deck.edited.pptx
```

Always use `--out` to write to a new file unless explicitly asked to overwrite.

## Tips

- **Placeholders** (ph0, ph1, etc.) are structural — edit text only, don't delete or resize
- **Inspect before editing** — use `show` to find the right shape index
- **Render after editing** — visual verification catches formatting issues
- **Copy formatting from inspect** — the JSON output works directly as edit input
