"""Edit/batch CLI commands for updating shape text."""

from __future__ import annotations

import json
import sys
from typing import Protocol

from .cli_common import find_shape, get_slide, is_text_shape, open_presentation, shape_text
from .cli_generate import warn_placeholder_text_limits
from .cli_text import text_preview, write_to_shape


class FileArgs(Protocol):
    file: str


class EditArgs(FileArgs, Protocol):
    slide: str
    shape: str
    text: str
    out: str | None


class BatchArgs(FileArgs, Protocol):
    edits: str
    out: str | None


def cmd_edit(args: EditArgs) -> int:
    prs = open_presentation(args.file)
    slide = get_slide(prs, args.slide)
    shape_raw = find_shape(slide, args.shape)

    if not is_text_shape(shape_raw):
        print(f"Error: shape '{args.shape}' has no text frame", file=sys.stderr)
        return 1

    shape = shape_raw

    before = shape_text(shape)
    out_path = args.out or args.file

    print(f"Before: {before}")
    print(f"After:  {text_preview(args.text)}")

    write_to_shape(shape, args.text)
    warn_placeholder_text_limits(slide, shape)

    prs.save(out_path)
    print(f"Saved → {out_path}")
    return 0


def cmd_batch(args: BatchArgs) -> int:
    prs = open_presentation(args.file)
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

        slide = get_slide(prs, str(slide_num))
        shape_raw = find_shape(slide, shape_id)

        if not is_text_shape(shape_raw):
            print(f"  [{idx+1}] SKIP {shape_id} — no text frame", file=sys.stderr)
            continue

        shape = shape_raw

        before = shape_text(shape)
        write_to_shape(shape, text_arg)

        print(f"  [{idx+1}] slide {slide_num} / {shape.name}")
        print(f"      Before: {before}")
        print(f"      After:  {text_preview(text_arg)}")

    prs.save(out_path)
    print(f"Saved → {out_path}  ({len(edits)} edits)")
    return 0
