"""
Unified CLI for PowerPoint inspection, editing, generation, and rendering.

    pptx <command> [args]

Inspect (progressive drill-down):
    pptx show          <file>                    # slide index with layout + word count
    pptx show          <file> <slide>            # all shapes sorted by position
    pptx show          <file> <slide> <shape>    # full detail (text/formatting/chart data)

Other inspect:
    pptx theme         <file>                    # scheme colors
    pptx xml           <file> <slide> <shape>    # raw XML

Edit:
    pptx edit          <file> <slide> <shape> <text> [--out PATH]
    pptx batch         <file> <edits.json | -> [--out PATH]

Slide management:
    pptx add-slide     <file> <layout> [--at N] [--out PATH]
    pptx delete-slide  <file> <slide> --confirm [--out PATH]
    pptx delete-shape  <file> <slide> <shape> [--out PATH]
    pptx insert        <deck.pptx> <source.pptx> [--at N] [--slides 1,3-5] [--out PATH]

Render:
    pptx render        <file> <slides> [--out DIR] [--dpi N] [--engine E]
    pptx crop          <png> <L> <T> <R> <B> [--out PATH]

Charts (from JSON):
    pptx charts        <spec.json> <output.pptx> [--template PATH] [--layout NAME]

Setup:
    pptx init          [-t template.pptx]        # create .clean-slides/ project dir
    pptx init-config   <template.pptx>           # generate config from a template

Generate (from YAML):
    pptx generate      <yaml...> [-o out.pptx] [-t template.pptx]
    pptx validate      <yaml...>
    pptx verify        <yaml...>

Auto-discovery: when no -t/-c flags are given, looks for .clean-slides/template.pptx
and .clean-slides/config.yaml walking up from the current directory.

Shape identified by index (int) or name (substring match).
Slide numbers are 1-indexed for inspect/edit, 0-indexed for generate --slide-index.
"""

from __future__ import annotations

import argparse
import json
import sys
from typing import Protocol

from .cli_common import find_shape as _find_shape
from .cli_common import get_slide as _get_slide
from .cli_common import is_text_shape as _is_text_shape
from .cli_common import open_presentation as _open
from .cli_common import shape_text as _shape_text
from .cli_deck import cmd_add_slide, cmd_delete_shape, cmd_delete_slide, cmd_insert
from .cli_generate import cmd_generate
from .cli_generate import warn_placeholder_text_limits as _warn_placeholder_text_limits
from .cli_inspect import (
    cmd_chart,
    cmd_color,
    cmd_layout,
    cmd_layouts,
    cmd_list,
    cmd_shape,
    cmd_show,
    cmd_slide,
    cmd_summary,
    cmd_theme,
    cmd_xml,
)
from .cli_render import cmd_charts, cmd_crop, cmd_render, cmd_screenshot
from .cli_spec_commands import cmd_preview, cmd_validate, cmd_verify
from .cli_template import cmd_init, cmd_init_config
from .cli_text import text_preview as _text_preview
from .cli_text import write_to_shape as _write_to_shape

# ============================================================================
# SHARED HELPERS
# ============================================================================


class _FileArgs(Protocol):
    file: str


class _EditArgs(_FileArgs, Protocol):
    slide: str
    shape: str
    text: str
    out: str | None


class _BatchArgs(_FileArgs, Protocol):
    edits: str
    out: str | None




# ============================================================================
# EDIT COMMANDS
# ============================================================================


def cmd_edit(args: _EditArgs) -> int:
    prs = _open(args.file)
    slide = _get_slide(prs, args.slide)
    shape_raw = _find_shape(slide, args.shape)

    if not _is_text_shape(shape_raw):
        print(f"Error: shape '{args.shape}' has no text frame", file=sys.stderr)
        return 1

    shape = shape_raw

    before = _shape_text(shape)
    out_path = args.out or args.file

    print(f"Before: {before}")
    print(f"After:  {_text_preview(args.text)}")

    _write_to_shape(shape, args.text)
    _warn_placeholder_text_limits(slide, shape)

    prs.save(out_path)
    print(f"Saved → {out_path}")
    return 0


def cmd_batch(args: _BatchArgs) -> int:
    prs = _open(args.file)
    out_path = args.out or args.file

    if args.edits == "-":
        edits = json.load(sys.stdin)
    else:
        with open(args.edits) as f:
            edits = json.load(f)

    for idx, edit in enumerate(edits):
        slide_num = edit["slide"]
        shape_id = str(edit["shape"])
        text_arg = edit["text"]

        slide = _get_slide(prs, str(slide_num))
        shape_raw = _find_shape(slide, shape_id)

        if not _is_text_shape(shape_raw):
            print(f"  [{idx+1}] SKIP {shape_id} — no text frame", file=sys.stderr)
            continue

        shape = shape_raw

        before = _shape_text(shape)
        _write_to_shape(shape, text_arg)

        print(f"  [{idx+1}] slide {slide_num} / {shape.name}")
        print(f"      Before: {before}")
        print(f"      After:  {_text_preview(text_arg)}")

    prs.save(out_path)
    print(f"Saved → {out_path}  ({len(edits)} edits)")
    return 0


# ============================================================================
# SLIDE MANAGEMENT COMMANDS
# ============================================================================


# ============================================================================
# RENDER COMMANDS
# ============================================================================


def _build_parser() -> argparse.ArgumentParser:
    from .cli_parser import build_parser

    return build_parser(
        cmd_show=cmd_show,
        cmd_list=cmd_list,
        cmd_summary=cmd_summary,
        cmd_slide=cmd_slide,
        cmd_shape=cmd_shape,
        cmd_chart=cmd_chart,
        cmd_layout=cmd_layout,
        cmd_layouts=cmd_layouts,
        cmd_theme=cmd_theme,
        cmd_color=cmd_color,
        cmd_xml=cmd_xml,
        cmd_edit=cmd_edit,
        cmd_batch=cmd_batch,
        cmd_add_slide=cmd_add_slide,
        cmd_delete_slide=cmd_delete_slide,
        cmd_delete_shape=cmd_delete_shape,
        cmd_insert=cmd_insert,
        cmd_render=cmd_render,
        cmd_crop=cmd_crop,
        cmd_charts=cmd_charts,
        cmd_generate=cmd_generate,
        cmd_validate=cmd_validate,
        cmd_preview=cmd_preview,
        cmd_verify=cmd_verify,
        cmd_init=cmd_init,
        cmd_init_config=cmd_init_config,
        cmd_screenshot=cmd_screenshot,
    )


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
