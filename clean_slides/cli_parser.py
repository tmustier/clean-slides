"""Argument parser construction for the `pptx` CLI."""

from __future__ import annotations

import argparse
from collections.abc import Callable
from typing import Any

CommandHandler = Callable[[Any], int]


def build_parser(
    *,
    cmd_show: CommandHandler,
    cmd_list: CommandHandler,
    cmd_summary: CommandHandler,
    cmd_slide: CommandHandler,
    cmd_shape: CommandHandler,
    cmd_chart: CommandHandler,
    cmd_layout: CommandHandler,
    cmd_layouts: CommandHandler,
    cmd_theme: CommandHandler,
    cmd_color: CommandHandler,
    cmd_xml: CommandHandler,
    cmd_edit: CommandHandler,
    cmd_batch: CommandHandler,
    cmd_add_slide: CommandHandler,
    cmd_delete_slide: CommandHandler,
    cmd_delete_shape: CommandHandler,
    cmd_insert: CommandHandler,
    cmd_render: CommandHandler,
    cmd_crop: CommandHandler,
    cmd_charts: CommandHandler,
    cmd_generate: CommandHandler,
    cmd_validate: CommandHandler,
    cmd_preview: CommandHandler,
    cmd_verify: CommandHandler,
    cmd_init: CommandHandler,
    cmd_init_config: CommandHandler,
    cmd_screenshot: CommandHandler,
) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx",
        description="PowerPoint CLI — inspect, edit, generate, and render slides.",
    )
    sub = parser.add_subparsers(dest="command", help="Commands")

    # ── Inspect ────────────────────────────────────────────────────────

    p = sub.add_parser("show", help="Inspect file → slides → shapes → detail")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slide", nargs="?", default=None, help="Slide number (1-indexed)")
    p.add_argument("shape", nargs="?", default=None, help="Shape index or name")
    p.set_defaults(func=cmd_show)

    p = sub.add_parser("list", help="List all slides")
    p.add_argument("file", help="PPTX file")
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("summary", help="Text summary of a slide")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slide", help="Slide number (1-indexed)")
    p.set_defaults(func=cmd_summary)

    p = sub.add_parser("slide", help="All shapes on a slide with positions")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slide", help="Slide number (1-indexed)")
    p.set_defaults(func=cmd_slide)

    p = sub.add_parser("shape", help="Full detail for a shape")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slide", help="Slide number (1-indexed)")
    p.add_argument("shape", help="Shape index or name")
    p.set_defaults(func=cmd_shape)

    p = sub.add_parser("chart", help="Chart data and styling")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slide", help="Slide number (1-indexed)")
    p.add_argument("shape", help="Shape index or name")
    p.set_defaults(func=cmd_chart)

    p = sub.add_parser("layout", help="Placeholder map for a slide's layout")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slide", help="Slide number (1-indexed)")
    p.set_defaults(func=cmd_layout)

    p = sub.add_parser("layouts", help="All layouts in presentation")
    p.add_argument("file", help="PPTX file")
    p.set_defaults(func=cmd_layouts)

    p = sub.add_parser("theme", help="Theme color scheme")
    p.add_argument("file", help="PPTX file")
    p.set_defaults(func=cmd_theme)

    p = sub.add_parser("color", help="Reverse lookup hex → theme name")
    p.add_argument("file", help="PPTX file")
    p.add_argument("hex", help="Hex color (e.g. #1B3D6F)")
    p.set_defaults(func=cmd_color)

    p = sub.add_parser("xml", help="Raw XML for a shape")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slide", help="Slide number (1-indexed)")
    p.add_argument("shape", help="Shape index or name")
    p.set_defaults(func=cmd_xml)

    # ── Edit ───────────────────────────────────────────────────────────

    p = sub.add_parser("edit", help="Edit shape text")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slide", help="Slide number (1-indexed)")
    p.add_argument("shape", help="Shape index or name")
    p.add_argument("text", help="New text (plain, JSON runs, or JSON paragraphs)")
    p.add_argument("--out", help="Output path (default: overwrite)")
    p.set_defaults(func=cmd_edit)

    p = sub.add_parser("batch", help="Batch edit from JSON")
    p.add_argument("file", help="PPTX file")
    p.add_argument("edits", help="JSON file with edits (or - for stdin)")
    p.add_argument("--out", help="Output path (default: overwrite)")
    p.set_defaults(func=cmd_batch)

    # ── Slide management ───────────────────────────────────────────────

    p = sub.add_parser("add-slide", help="Add a slide from a layout")
    p.add_argument("file", help="PPTX file")
    p.add_argument("layout", help="Layout index or name")
    p.add_argument("--at", type=int, help="Insert at position (1-indexed)")
    p.add_argument("--out", help="Output path (default: overwrite)")
    p.set_defaults(func=cmd_add_slide)

    p = sub.add_parser("delete-slide", help="Delete a slide")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slide", help="Slide number (1-indexed)")
    p.add_argument("--confirm", action="store_true", help="Required safety flag")
    p.add_argument("--out", help="Output path (default: overwrite)")
    p.set_defaults(func=cmd_delete_slide)

    p = sub.add_parser("delete-shape", help="Delete a shape from a slide")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slide", help="Slide number (1-indexed)")
    p.add_argument("shape", help="Shape index or name")
    p.add_argument("--out", help="Output path (default: overwrite)")
    p.set_defaults(func=cmd_delete_shape)

    p = sub.add_parser("insert", aliases=["merge"], help="Insert slides from another PPTX")
    p.add_argument("file", help="Target PPTX file")
    p.add_argument("source", help="Source PPTX file")
    p.add_argument("--slides", help="Source slide selection (e.g. 1,3-5). Default: all")
    p.add_argument("--at", type=int, help="Insert at position (1-indexed)")
    p.add_argument("--out", help="Output path (default: overwrite)")
    p.set_defaults(func=cmd_insert)

    # ── Render ─────────────────────────────────────────────────────────

    p = sub.add_parser("render", help="Render slides to PNG")
    p.add_argument("file", help="PPTX file")
    p.add_argument("slides", help="Slide numbers: 1, 1,3-5, or 1,8,16 (1-indexed)")
    p.add_argument("--out", help="Output directory")
    p.add_argument("--dpi", type=int, default=150, help="DPI (default: 150)")
    p.add_argument("--engine", choices=["powerpoint", "libreoffice"], help="Rendering engine")
    p.set_defaults(func=cmd_render)

    p = sub.add_parser("crop", help="Crop a region from a PNG")
    p.add_argument("png", help="PNG file")
    p.add_argument("left", type=float, help="Left edge (inches)")
    p.add_argument("top", type=float, help="Top edge (inches)")
    p.add_argument("right", type=float, help="Right edge (inches)")
    p.add_argument("bottom", type=float, help="Bottom edge (inches)")
    p.add_argument("--out", help="Output path")
    p.set_defaults(func=cmd_crop)

    # ── Charts (JSON) ──────────────────────────────────────────────────

    p = sub.add_parser("charts", help="Generate charts from JSON")
    p.add_argument("input", help="Input JSON spec")
    p.add_argument("output", help="Output PPTX")
    p.add_argument("-t", "--template", help="Template PPTX")
    p.add_argument("--layout", help="Layout name when using a template")
    p.add_argument("--expected-template", help="Expected template alias/path")
    p.set_defaults(func=cmd_charts)

    # ── Generate (YAML pipeline) ───────────────────────────────────────

    p = sub.add_parser("generate", aliases=["gen"], help="Generate PPTX from YAML")
    p.add_argument("input", nargs="+", help="Input YAML file(s)")
    p.add_argument("-o", "--output", help="Output PPTX (default: output.pptx)")
    p.add_argument("-t", "--template", help="Template PPTX")
    p.add_argument("-c", "--config", help="Template config YAML (overrides built-in config)")
    p.add_argument("--slide-index", type=int, help="Target existing slide (0-based)")
    p.add_argument("--keep-existing", action="store_true", help="Keep existing shapes")
    p.add_argument("--detail", action="store_true", help="Verbose report")
    p.set_defaults(func=cmd_generate)

    p = sub.add_parser("validate", aliases=["val"], help="Validate YAML schema")
    p.add_argument("input", nargs="+", help="Input YAML file(s)")
    p.add_argument("-c", "--config", help="Template config YAML (overrides built-in config)")
    p.set_defaults(func=cmd_validate)

    p = sub.add_parser("preview", help="Preview table structure from YAML")
    p.add_argument("input", nargs="+", help="Input YAML file(s)")
    p.set_defaults(func=cmd_preview)

    p = sub.add_parser("verify", aliases=["check"], help="Verify layout without generating")
    p.add_argument("input", nargs="+", help="Input YAML file(s)")
    p.add_argument("-c", "--config", help="Template config YAML (overrides built-in config)")
    p.add_argument("--detail", action="store_true", help="Verbose report")
    p.add_argument("--json", help="Write report JSON to file")
    p.set_defaults(func=cmd_verify)

    p = sub.add_parser("init", help="Set up .clean-slides/ project directory")
    p.add_argument("-t", "--template", help="Custom template PPTX (default: bundled example)")
    p.add_argument("-o", "--output", help="Target directory (default: cwd)")
    p.set_defaults(func=cmd_init)

    p = sub.add_parser("init-config", help="Generate starter template-config.yaml from a PPTX")
    p.add_argument("file", help="Template PPTX to introspect")
    p.add_argument("-o", "--output", help="Output YAML path (default: stdout)")
    p.set_defaults(func=cmd_init_config)

    p = sub.add_parser("screenshot", aliases=["ss"], help="Screenshot PPTX via LibreOffice")
    p.add_argument("input", nargs="+", help="Input PPTX file(s)")
    p.add_argument("-o", "--output-dir", help="Output directory")
    p.add_argument("--slide", type=int, default=0, help="Slide index (0-based)")
    p.add_argument("--soffice", help="Path to soffice binary")
    p.set_defaults(func=cmd_screenshot)

    return parser
