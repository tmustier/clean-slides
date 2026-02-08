# Template Config

Template-derived values live in `clean_slides/template-config.yaml` and are
loaded at runtime by `clean_slides.template_config`.

Use this file to align the codebase with a new slide master or template
without touching Python constants.

## Custom Templates

You can use any PPTX template by providing a custom config:

```bash
# Generate a starter config from your template
pptx init-config my-template.pptx -o my-config.yaml

# Use it when generating slides
pptx generate spec.yaml -t my-template.pptx -c my-config.yaml -o output.pptx
```

The `--config` (`-c`) flag is supported by `generate`, `validate`, and `verify`.

## Validation

The loader validates required sections and keys. Missing or malformed
entries raise a `ValueError` with a detailed list of errors.

Optional sections (`placeholders`, `default_colors`) are validated when
present but not required — sensible defaults are used when they're absent.

## Top-level Sections

| Section | Required | Purpose |
|---------|----------|---------|
| `colors` | ✅ | Hex palette for design-system colors |
| `fonts` | ✅ | Headline/body font families |
| `font_sizes` | ✅ | Standard font sizes (pt) |
| `layout` | ✅ | Slide dimensions + key Y positions (EMU) |
| `dividers` | ✅ | Divider line weights (pt) |
| `bullets` | ✅ | Placeholder lstStyle defaults + bullet levels |
| `table_defaults` | ✅ | Table sizing defaults + legend sizing |
| `icons` | ✅ | Default icon size |
| `moon` | ✅ | RAG moon geometry + colors |
| `placeholders` | ❌ | Placeholder indices (title, subtitle, tracker) |
| `default_colors` | ❌ | Default theme color names for body text, headers, dividers, links |
| `text_limits` | ❌ | Soft limits for agent-friendly warnings |

## New Sections (for custom templates)

### `placeholders`

Maps semantic names to PowerPoint placeholder indices. Defaults:

```yaml
placeholders:
  title: 0          # slide title placeholder index
  subtitle: 1       # subtitle placeholder index
  tracker: 17       # on-page tracker/breadcrumb (omit if not present)
```

Use `pptx layouts <template>` to find placeholder indices for your template.

### `default_colors`

Default theme color names used when YAML specs don't override. These are
PowerPoint theme slot names (e.g. `tx1`, `accent1`, `dk2`) or hex colors:

```yaml
default_colors:
  body_text: "tx1"        # body text color
  col_header: "tx1"       # column header text color
  col_superheader: "tx1"  # column superheader text color
  row_header: "accent2"   # row header text color
  row_superheader: "tx1"  # row superheader text color
  divider: "tx1"          # header divider line color
  link: "accent2"         # hyperlink/appendix reference color
```

Use `pptx theme <template>` to see what colors are mapped to each theme slot.

## Editing Guide

1. Run `pptx init-config <template>` to generate a starter config.
2. Update values from the slide master or design system.
3. Run `pptx generate` with `-c` to test, then `pptx render` to verify.

## Notes

- **Units**: EMU values are used for positions and dimensions; font sizes are in points.
- **Bullet levels**: `bullets.levels` mirrors the slide master's `lstStyle`.
- **Moon geometry**: `moon.group` values come from the MoonLegend group on the master.
- **Runtime override**: `set_template_config(path)` replaces the active config at runtime.
