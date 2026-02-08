# Clean Slides

PowerPoint CLI for inspection, editing, generation, and rendering. Features include:

- **YAML → PPTX** table generation with automatic layout solving
- **Custom templates** with configurable colours, fonts, and content areas
- **Inline formatting**: `**bold**`, `*italic*`, `[links](url)` with clickable hyperlinks
- **Icons, bullets, row/column headers, superheaders, sidebars**
- **Render** slides to PNG via PowerPoint or LibreOffice

## Quick Start

Install and generate a slide in two commands:

```bash
pip install -e .
pptx init                              # sets up .clean-slides/ with an example template
```

Then create a YAML spec and generate:

```bash
cat > slide.yaml << 'EOF'
title: Revenue Summary
table:
  rows: 3
  cols: 2
  has_col_header: true
  col_headers: ["Metric", "Value"]
  cells:
    - ["Revenue", "$1.2M"]
    - ["Growth", "15% YoY"]
EOF

pptx generate slide.yaml -o output.pptx
```

The template and config in `.clean-slides/` are picked up automatically — no flags needed.

## Initialising with Your Own Template

To use your corporate or custom template instead of the bundled example:

```bash
# Option A: init with your template (generates a starter config automatically)
pptx init -t my-template.pptx

# Option B: generate config separately, then set up the project
pptx init-config my-template.pptx -o my-config.yaml
# review and adjust my-config.yaml, then:
mkdir .clean-slides
cp my-template.pptx .clean-slides/template.pptx
cp my-config.yaml .clean-slides/config.yaml
```

Once `.clean-slides/` exists, all `pptx generate` commands in that directory (or below) automatically use your template and config. Edit `.clean-slides/config.yaml` to adjust colours, fonts, placeholders, and bullet styles.

See [docs/TEMPLATE-CONFIG.md](docs/TEMPLATE-CONFIG.md) for full config documentation, and [examples/custom-template/](examples/custom-template/) for a worked example.

## Commands

### Inspect

```bash
pptx show <file>                       # slide index with layout, shape count, words
pptx show <file> <slide>               # all shapes sorted by position
pptx show <file> <slide> <shape>       # full detail (text, formatting, chart data)

pptx theme <file>                      # colour scheme
pptx xml <file> <slide> <shape>        # raw XML (escape hatch)
```

### Edit

```bash
pptx edit <file> <slide> <shape> <text> [--out PATH]
pptx batch <file> <edits.json> [--out PATH]
```

**Text formats:**
- Plain: `"Hello world"`
- With formatting: `'{"paragraphs": [{"runs": [{"text": "Bold", "bold": true}]}]}'`

Run fields: `text`, `font`, `size`, `bold`, `italic`, `underline`, `color`, `superscript`, `subscript`

### Slide Management

```bash
pptx add-slide <file> <layout> [--at N] [--out PATH]
pptx delete-slide <file> <slide> --confirm [--out PATH]
pptx delete-shape <file> <slide> <shape> [--out PATH]

# Insert/merge slides from another deck
pptx insert <deck.pptx> <source.pptx> [--at N] [--slides 1,3-5] [--out PATH]
```

### Render

```bash
pptx render <file> <slide> [--out DIR] [--dpi N]
pptx crop <png> <L> <T> <R> <B> [--out PATH]
```

### Setup

```bash
pptx init [-t template.pptx]           # create .clean-slides/ project directory
pptx init-config <template.pptx>       # generate config YAML from a template
```

### Generate (YAML → PPTX)

```bash
pptx generate <yaml...> [-o out.pptx] [-t template.pptx] [-c config.yaml]
pptx validate <yaml...> [-c config.yaml]
pptx verify <yaml...> [--detail] [-c config.yaml]
```

## YAML Input

```yaml
title: Quarterly Update

table:
  rows: 4
  cols: 3
  has_col_header: true
  has_row_header: true
  col_headers: ["Region", "Revenue", "Growth"]
  cells:
    - ["North America", "$4.2M", "12%"]
    - ["Europe", "$2.8M", "8%"]
    - ["Asia-Pacific", "$1.5M", "22%"]
```

Supports `**bold**`, `*italic*`, and `[clickable links](https://example.com)` in any cell.

Full schema: [docs/INPUT-SCHEMA.md](docs/INPUT-SCHEMA.md)

## Development

```bash
pip install -e '.[dev]'
pre-commit install

# Run checks
python -m pytest -q
pyright clean_slides/
```
