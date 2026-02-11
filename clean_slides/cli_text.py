"""Text argument parsing and writing helpers for CLI edit/batch commands."""

from __future__ import annotations

import contextlib
import json
from typing import TYPE_CHECKING

from pptx.shapes.autoshape import Shape
from typing_extensions import TypeGuard

if TYPE_CHECKING:
    from .editor import ParagraphSpec

RunOverridesMap = dict[str, object]
TextRun = tuple[str, RunOverridesMap]


def _is_object_list(value: object) -> TypeGuard[list[object]]:
    return isinstance(value, list)


def _is_str_object_dict(value: object) -> TypeGuard[dict[str, object]]:
    return isinstance(value, dict)


def _coerce_overrides(value: object) -> RunOverridesMap:
    if _is_str_object_dict(value):
        return dict(value)
    return {}


def _expand_newlines(result: list[TextRun], text: str, opts: RunOverridesMap) -> None:
    """Split text on newlines, inserting line break markers."""
    parts = text.split("\n")
    for i, part in enumerate(parts):
        if i > 0:
            result.append(("\n", {}))
        if part:
            result.append((part, opts))


def parse_text_arg(text: object) -> list[TextRun]:
    """
    Parse text argument into list of (text, overrides) tuples.

    Formats:
        "Plain text"                                → [("Plain text", {})]
        "Line 1\\nLine 2"                            → with line breaks
        [["Bold ", {"bold": true}], [" normal"]]    → JSON runs
    """

    runs: list[object] | None = None

    if _is_object_list(text):
        runs = text
    elif isinstance(text, str) and text.startswith("["):
        try:
            parsed: object = json.loads(text)
        except (json.JSONDecodeError, ValueError):
            parsed = None

        if _is_object_list(parsed):
            runs = parsed

    if runs is not None:
        result: list[TextRun] = []

        for item in runs:
            if isinstance(item, str):
                t = item
                opts: RunOverridesMap = {}
            elif _is_object_list(item):
                t = str(item[0]) if len(item) > 0 else ""
                opts_raw: object = item[1] if len(item) > 1 else {}
                opts = _coerce_overrides(opts_raw)
            elif _is_str_object_dict(item) and "text" in item:
                t = str(item.get("text", ""))
                opts = {k: v for k, v in item.items() if k != "text"}
            else:
                continue

            _expand_newlines(result, t, opts)

        return result

    text_str = text if isinstance(text, str) else str(text)
    normalized = text_str.replace("\n", "\\n")

    result: list[TextRun] = []
    lines = normalized.split("\\n")
    for i, line in enumerate(lines):
        if i > 0:
            result.append(("\n", {}))
        result.append((line, {}))

    return result


def is_paragraphs_format(text_arg: object) -> bool:
    """Check if text_arg is a multi-paragraph spec."""
    if _is_str_object_dict(text_arg) and "paragraphs" in text_arg:
        return True

    if isinstance(text_arg, str) and text_arg.strip().startswith("{"):
        try:
            parsed: object = json.loads(text_arg)
        except (json.JSONDecodeError, ValueError):
            return False

        return _is_str_object_dict(parsed) and "paragraphs" in parsed

    return False


def parse_paragraphs_arg(text_arg: object) -> list[ParagraphSpec]:
    """Parse a multi-paragraph spec."""

    parsed: object = text_arg
    if isinstance(parsed, str):
        parsed = json.loads(parsed)

    if not _is_str_object_dict(parsed):
        raise ValueError("paragraphs spec must be an object")

    paragraphs_raw = parsed.get("paragraphs")
    if not _is_object_list(paragraphs_raw):
        raise ValueError("paragraphs spec missing 'paragraphs' list")

    paragraphs: list[ParagraphSpec] = []

    for p_obj in paragraphs_raw:
        if not _is_str_object_dict(p_obj):
            continue

        para: ParagraphSpec = {}

        if "runs" in p_obj:
            para["runs"] = p_obj["runs"]

        if "level" in p_obj:
            level_raw = p_obj["level"]
            if isinstance(level_raw, (int, float, str)) and not isinstance(level_raw, bool):
                with contextlib.suppress(ValueError):
                    para["level"] = int(level_raw)

        if "alignment" in p_obj:
            alignment_raw = p_obj["alignment"]
            if isinstance(alignment_raw, str):
                para["alignment"] = alignment_raw

        if "spacing_before" in p_obj:
            sb_raw = p_obj["spacing_before"]
            if isinstance(sb_raw, (int, float)) and not isinstance(sb_raw, bool):
                para["spacing_before"] = float(sb_raw)

        if "spacing_after" in p_obj:
            sa_raw = p_obj["spacing_after"]
            if isinstance(sa_raw, (int, float)) and not isinstance(sa_raw, bool):
                para["spacing_after"] = float(sa_raw)

        if "line_spacing" in p_obj:
            ls_raw = p_obj["line_spacing"]
            if isinstance(ls_raw, (int, float)) and not isinstance(ls_raw, bool):
                para["line_spacing"] = float(ls_raw)

        if "bullet" in p_obj:
            bullet_raw = p_obj["bullet"]
            if isinstance(bullet_raw, (bool, str)):
                para["bullet"] = bullet_raw

        paragraphs.append(para)

    return paragraphs


def write_to_shape(shape: Shape, text_arg: object) -> None:
    """Write content to a shape, dispatching based on format."""
    from .editor import add_line_break, add_run, snapshot_defaults, write_paragraphs

    if is_paragraphs_format(text_arg):
        write_paragraphs(shape, parse_paragraphs_arg(text_arg))
        return

    runs = parse_text_arg(text_arg)
    tf = shape.text_frame
    defaults = snapshot_defaults(tf)
    defaults_map: dict[str, object] = dict(defaults)

    tf.clear()
    p = tf.paragraphs[0]
    for text, overrides in runs:
        if text == "\n":
            add_line_break(p)
            continue

        merged: dict[str, object] = dict(defaults_map)
        merged.update(overrides)
        add_run(p, text, **merged)


def text_preview(text_arg: object) -> str:
    """Compact preview for before/after display."""

    if is_paragraphs_format(text_arg):
        paras = parse_paragraphs_arg(text_arg)
        parts: list[str] = []

        for p in paras:
            lvl = p.get("level", 0)
            prefix = "  " * lvl + ("• " if lvl > 0 else "")
            runs = p.get("runs", "")

            if isinstance(runs, str):
                parts.append(f"{prefix}{runs}")
                continue

            run_texts: list[str] = []
            if _is_object_list(runs):
                for r in runs:
                    if isinstance(r, str):
                        run_texts.append(r)
                    elif _is_str_object_dict(r) and "text" in r:
                        run_texts.append(str(r.get("text", "")))
                    elif _is_object_list(r):
                        run_texts.append(str(r[0]) if len(r) > 0 else "")
                    else:
                        run_texts.append(str(r))
            else:
                run_texts.append(str(runs))

            parts.append(f"{prefix}{''.join(run_texts)}")

        return " ¶ ".join(parts)

    runs = parse_text_arg(text_arg)
    parts: list[str] = []
    for text, opts in runs:
        if text == "\n":
            parts.append("\\n")
        elif opts:
            flags = ",".join(f"{k}={v}" for k, v in opts.items())
            parts.append(f"[{text}|{flags}]")
        else:
            parts.append(text)

    return "".join(parts)
