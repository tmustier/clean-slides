"""Render/chart CLI commands."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Protocol

from .cli_common import open_presentation
from .cli_deck import parse_slide_selection
from .screenshot import crop_region


class FileArgs(Protocol):
    file: str


class RenderArgs(FileArgs, Protocol):
    slides: str
    dpi: int
    out: str | None
    engine: str | None


class CropArgs(Protocol):
    png: str
    left: float
    top: float
    right: float
    bottom: float
    out: str | None


class ChartsArgs(Protocol):
    input: str
    output: str
    template: str | None
    layout: str | None
    expected_template: str | None


def cmd_render(args: RenderArgs) -> int:
    from .screenshot import render_slides

    prs = open_presentation(args.file)
    total = len(prs.slides)
    try:
        slide_nums = parse_slide_selection(args.slides, total)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    paths = render_slides(
        args.file,
        slide_nums,
        dpi=args.dpi,
        output_dir=args.out,
        engine=args.engine,
    )
    for p in paths:
        print(p)
    return 0


def cmd_crop(args: CropArgs) -> int:
    result = crop_region(
        args.png,
        args.left,
        args.top,
        args.right,
        args.bottom,
        output_path=args.out,
    )
    print(result)
    return 0


def cmd_charts(args: ChartsArgs) -> int:
    from .charts import generate_charts_from_json

    input_path = Path(args.input)
    output_path = Path(args.output)
    template_path = Path(args.template) if args.template else None

    try:
        generate_charts_from_json(
            input_path,
            output_path,
            template=template_path,
            layout=args.layout,
            expected_template=args.expected_template,
        )
    except (FileNotFoundError, ImportError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Saved → {output_path}")
    return 0
