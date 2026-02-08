"""
Clean Slides — PowerPoint CLI for inspection, editing, generation, and rendering.

CLI entry point: `pptx`

Pipeline: YAML → TableSpec → ConstraintSolver → TableLayout → TableRenderer → PPTX

Modules:
    cli             Unified CLI (pptx command)

    # Inspect & edit
    inspect_pptx    Read-only introspection (shapes, text, charts, layouts, themes)
    editor          Safe writing to placeholders and text frames
    screenshot      Render slides to PNG (PowerPoint + LibreOffice)

    # Table generation pipeline
    spec            TableSpec, ContentArea, TableLayout (data model)
    icons           IconSet, presets, cell helpers
    content         Paragraph, normalize_cell (text normalization)
    measure         paragraph_height, cell_content_height (shared measurement)
    sizing          ColumnSizer, RowSizer (width/height algorithms)
    solver          ConstraintSolver (orchestrates sizing + verification)
    renderer        TableRenderer (layout → slide shapes)
    verification    LayoutVerifier, LayoutReport (post-layout checks)
    metadata        fill_slide_metadata (title/subtitle/tracker)

    # Shared
    xml_helpers     OOXML generation + idempotent XML mutation
    constants       Design system values from template config
    template_config Template config loader (template-config.yaml)
    text_metrics    Font measurement utilities
    placeholder     Fill placeholder content
"""

__version__ = "0.3.0"
