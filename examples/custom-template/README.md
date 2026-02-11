# Custom Template Example

This directory contains a complete example of using `pptx generate` with a custom PowerPoint template — from creating the template to producing a finished 11-slide presentation.

## What's here

| File | Description |
|------|-------------|
| `example-template.pptx` | A custom template with 7 layouts (Title - Contrast, Default, 1/2, 2/3, Contrast 1/2, Section, Quote) |
| `example-config.yaml` | Template config: colors, fonts, spacing, placeholder indices |
| `vykhukhol.pptx` | Finished 11-slide presentation about the Russian desman, using all 6 content layouts |

## How it works

`pptx generate` needs two things to produce slides from YAML:

1. **A template** (`.pptx`) — provides slide layouts, master shapes, theme colors, and fonts
2. **A config** (`.yaml`) — tells the generator how to use the template: which colors map to which roles, font sizes, content area geometry, placeholder indices

Without either, the tool falls back to a blank presentation with built-in defaults.

## Quick start

```bash
# Generate a single slide
pptx generate slide.yaml \
  -t examples/custom-template/example-template.pptx \
  -c examples/custom-template/example-config.yaml \
  -o output.pptx

# Render to PNG for verification
pptx render output.pptx 1 --out renders/
```

## Creating your own template

### 1. Start from any PowerPoint file

Any `.pptx` works — a corporate template, a downloaded theme, or a file you design from scratch. The generator uses the **slide layouts** and **theme** from the template; it ignores existing content slides.

### 2. Inspect it

```bash
pptx layouts your-template.pptx    # see layout names, placeholders, content areas
pptx theme your-template.pptx      # see theme color mappings
```

### 3. Generate a starter config

```bash
pptx init-config your-template.pptx -o your-config.yaml
```

This extracts theme colors, fonts, and layout geometry into a config file with sensible defaults. Review and adjust:

- **`colors`** — verify accent colors match the template's design intent
- **`fonts`** — check that headline/body fonts match actual slide text (templates often override theme fonts)
- **`font_sizes`** — adjust title, body, and table sizes to taste
- **`placeholders`** — verify title (usually 0) and subtitle (usually 1) indices
- **`layout`** — tweak `content_start_y_emu` if tables appear too high or low

### 4. Iterate

```bash
# Generate a test slide
pptx generate test.yaml -t your-template.pptx -c your-config.yaml -o test.pptx

# Render and check
pptx render test.pptx 1 --out renders/
```

Start with a simple 3×2 table. Adjust the config, regenerate, re-render. Once it looks right, build out your deck.

## About the example template

The `example-template.pptx` was derived from a corporate template by:

- Stripping 27 layouts down to 7 essential ones
- Replacing proprietary artwork with solid theme-colored rectangles
- Applying a "Warm Editorial" color palette (terracotta, teal, indigo, warm neutrals)
- Setting Rockwell (headings) + Trebuchet MS (body) as theme fonts
- Removing embedded objects, third-party add-in shapes, and client-specific elements

The build script is in `../../sandbox/build-example-template.py` if you want to see the transformation.

## About the example presentation

`vykhukhol.pptx` is an 11-slide presentation about the Russian desman (_Desmana moschata_), a critically endangered semiaquatic mole found only in Russia. It exercises all 6 content layouts:

| Slide | Layout | Content |
|-------|--------|---------|
| 1 | Title - Contrast | Title slide with dark background |
| 2 | Section | "A Mole That Chose Rivers" |
| 3 | Default | Species profile table (row headers, no col header) |
| 4 | 2/3 | Adaptations table + sidebar (snoot, floof, musk) |
| 5 | Section | "Life in the Floodplain" |
| 6 | Default | Ecology & behavior table |
| 7 | Section | "Edge of Extinction" |
| 8 | Default | Population decline timeline |
| 9 | Contrast 1/2 | Threats table + signs of hope (dark sidebar) |
| 10 | 2/3 | Conservation organizations + the people behind them |
| 11 | Quote | Closing quote from researcher Masha Onufrenya |

The build script is in `../../sandbox/vykhukhol-deck.py`.
