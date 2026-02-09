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
from collections.abc import Iterable
from copy import deepcopy
from pathlib import Path
from typing import TYPE_CHECKING, Any, Protocol, Union, cast

import yaml
from pptx import Presentation
from pptx.presentation import Presentation as PresentationObj
from pptx.shapes.autoshape import Shape
from pptx.shapes.base import BaseShape
from pptx.slide import Slide, SlideLayout
from pptx.util import Emu
from typing_extensions import TypeGuard

if TYPE_CHECKING:
    from .editor import ParagraphSpec

import contextlib

from .constants import EMU_PER_INCH, Fonts, FontSizes, Layout, TableDefaults
from .metadata import fill_slide_metadata
from .placeholder import fill_placeholders
from .renderer import TableRenderer
from .screenshot import ScreenshotGenerator, crop_region
from .solver import ConstraintSolver, SolveOptions
from .spec import ContentArea, TableSpec
from .template_config import TEMPLATE_CONFIG, set_template_config
from .text_metrics import EMU_PER_PT, TextMetrics

# ============================================================================
# SHARED HELPERS
# ============================================================================


class _ToDict(Protocol):
    def to_dict(self) -> object: ...


def _has_to_dict(obj: object) -> TypeGuard[_ToDict]:
    return hasattr(obj, "to_dict")


def _is_object_list(value: object) -> TypeGuard[list[object]]:
    return isinstance(value, list)


def _cells_can_provide_row_headers(raw_cells: object) -> bool:
    """True when *raw_cells* is a non-empty list of lists each with ≥ 2 elements."""
    if not isinstance(raw_cells, list):
        return False
    rows = cast(list[object], raw_cells)
    if len(rows) == 0:
        return False
    return all(not (not isinstance(row, list) or len(cast(list[object], row)) < 2) for row in rows)


def _is_str_object_dict(value: object) -> TypeGuard[dict[str, object]]:
    return isinstance(value, dict)


def _is_text_shape(shape: BaseShape) -> TypeGuard[Shape]:
    return isinstance(shape, Shape) and shape.has_text_frame


class _FileArgs(Protocol):
    file: str


class _ShowArgs(_FileArgs, Protocol):
    slide: str | None
    shape: str | None


class _SlideArgs(_FileArgs, Protocol):
    slide: str


class _ColorArgs(_FileArgs, Protocol):
    hex: str


class _XmlArgs(_FileArgs, Protocol):
    slide: str
    shape: str


class _EditArgs(_FileArgs, Protocol):
    slide: str
    shape: str
    text: str
    out: str | None


class _BatchArgs(_FileArgs, Protocol):
    edits: str
    out: str | None


class _AddSlideArgs(_FileArgs, Protocol):
    layout: str
    at: int | None
    out: str | None


class _DeleteSlideArgs(_FileArgs, Protocol):
    slide: str
    confirm: bool
    out: str | None


class _DeleteShapeArgs(_FileArgs, Protocol):
    slide: str
    shape: str
    out: str | None


class _InsertArgs(_FileArgs, Protocol):
    source: str
    at: int | None
    slides: str | None
    out: str | None


class _RenderArgs(_FileArgs, Protocol):
    slides: str
    dpi: int
    out: str | None
    engine: str | None


class _CropArgs(Protocol):
    png: str
    left: float
    top: float
    right: float
    bottom: float
    out: str | None


class _ChartsArgs(Protocol):
    input: str
    output: str
    template: str | None
    layout: str | None
    expected_template: str | None
    module_path: str | None


class _InputArgs(Protocol):
    input: list[str]


class _GenerateArgs(_InputArgs, Protocol):
    output: str | None
    template: str | None
    config: str | None
    slide_index: int | None
    keep_existing: bool
    detail: bool


class _VerifyArgs(_InputArgs, Protocol):
    detail: bool
    json: str | None
    config: str | None


class _ScreenshotArgs(_InputArgs, Protocol):
    output_dir: str | None
    slide: int
    soffice: str | None


class _ValidateArgs(_InputArgs, Protocol):
    config: str | None


class _InitArgs(Protocol):
    template: str | None
    output: str | None


class _InitConfigArgs(Protocol):
    file: str
    output: str | None


# ── Project-level auto-discovery ───────────────────────────────────────

PROJECT_DIR_NAME = ".clean-slides"
_CONFIG_NAME = "config.yaml"
_TEMPLATE_NAME = "template.pptx"


def _discover_project_dir() -> Path | None:
    """Walk from CWD to filesystem root looking for a `.clean-slides/` directory."""
    cur = Path.cwd().resolve()
    for parent in [cur, *cur.parents]:
        candidate = parent / PROJECT_DIR_NAME
        if candidate.is_dir():
            return candidate
    return None


def _discover_config() -> Path | None:
    """Return the project config path if auto-discovered."""
    proj = _discover_project_dir()
    if proj is not None:
        cfg = proj / _CONFIG_NAME
        if cfg.is_file():
            return cfg
    return None


def _discover_template() -> Path | None:
    """Return the project template path if auto-discovered."""
    proj = _discover_project_dir()
    if proj is not None:
        tpl = proj / _TEMPLATE_NAME
        if tpl.is_file():
            return tpl
    return None


def _apply_config(config_path: str | None) -> None:
    """Load template config — explicit path, auto-discovered, or built-in defaults."""
    if config_path is not None:
        set_template_config(Path(config_path))
        return
    # Auto-discover from .clean-slides/
    discovered = _discover_config()
    if discovered is not None:
        set_template_config(discovered)


def _open(path: Union[str, Path]) -> PresentationObj:
    """Open a PPTX file."""
    return Presentation(str(path))


def _get_slide(prs: PresentationObj, num: Union[str, int]) -> Slide:
    """Get slide by 1-indexed number."""
    idx = int(num) - 1
    slides: list[Slide] = list(prs.slides)
    if idx < 0 or idx >= len(slides):
        print(f"Error: slide {num} out of range (1-{len(slides)})", file=sys.stderr)
        sys.exit(1)
    return slides[idx]


def _find_shape(slide: Slide, identifier: str) -> BaseShape:
    """Find shape by index (int) or name (substring match)."""
    try:
        idx = int(identifier)
        shapes: list[BaseShape] = list(slide.shapes)
        if 0 <= idx < len(shapes):
            return shapes[idx]
        for s in shapes:
            if s.shape_id == idx:
                return s
    except ValueError:
        pass

    ident_lower = identifier.lower()
    for s in slide.shapes:
        if ident_lower in s.name.lower():
            return s

    print(f"Error: shape '{identifier}' not found", file=sys.stderr)
    sys.exit(1)


def _find_layout(prs: PresentationObj, identifier: str, fallback: bool = False) -> SlideLayout:
    """Find layout by index (int) or name (substring match).

    If fallback=True, returns a best-effort layout instead of exiting.
    """
    try:
        idx = int(identifier)
        layouts: list[SlideLayout] = list(prs.slide_layouts)
        if 0 <= idx < len(layouts):
            return layouts[idx]
    except ValueError:
        pass

    ident_lower = identifier.lower()
    for layout in prs.slide_layouts:
        if ident_lower in layout.name.lower():
            return layout

    if fallback:
        layouts = list(prs.slide_layouts)

        # Prefer a sensible, content-friendly layout.
        # NOTE: templates vary widely; avoid hardcoding indices.
        for key in ("default", "blank"):
            for layout in layouts:
                if layout.name and layout.name.lower() == key:
                    return layout
        for key in ("default", "blank"):
            for layout in layouts:
                if layout.name and key in layout.name.lower():
                    return layout

        return layouts[0]

    print(f"Error: layout '{identifier}' not found", file=sys.stderr)
    sys.exit(1)


def _try_find_layout(prs: PresentationObj, name: str) -> SlideLayout | None:
    """Best-effort lookup by layout name (case-insensitive).

    Returns None if no match.
    """
    name_lower = name.lower()
    layouts: list[SlideLayout] = list(prs.slide_layouts)

    for layout in layouts:
        if layout.name and layout.name.lower() == name_lower:
            return layout

    for layout in layouts:
        if layout.name and name_lower in layout.name.lower():
            return layout

    return None


def _content_area_from_layout(slide_layout: SlideLayout) -> ContentArea | None:
    """Extract the primary content area from a slide layout.

    Finds the best OBJECT / BODY placeholder to use as the table content
    area.  When multiple content placeholders exist (e.g. "Two Content"),
    picks the largest; ties are broken by topmost then leftmost position
    so the primary (upper-left) area wins consistently.

    Returns ``None`` when no suitable placeholder exists.
    """
    from .spec import ContentArea

    # Placeholder type names that represent content areas (OBJECT, BODY,
    # TABLE, CHART, etc.).  We exclude TITLE, SUBTITLE, DATE, FOOTER,
    # SLIDE_NUMBER, HEADER.
    _SKIP_TYPES = {"TITLE", "CENTER_TITLE", "SUBTITLE", "DATE", "FOOTER", "SLIDE_NUMBER", "HEADER"}

    footer_y = int(Layout.FOOTER_LINE_Y)

    # Collect candidate content placeholders.
    candidates: list[tuple[int, int, int, ContentArea]] = []
    for ph in slide_layout.placeholders:
        pf = ph.placeholder_format
        type_name = str(pf.type).split("(")[0].strip() if pf.type is not None else ""
        if type_name in _SKIP_TYPES:
            continue
        ph_area = int(ph.width) * int(ph.height)
        ph_bottom = int(ph.top) + int(ph.height)
        area = ContentArea(
            x=int(ph.left),
            y=int(ph.top),
            width=int(ph.width),
            height=int(min(ph_bottom, footer_y) - int(ph.top)),
        )
        candidates.append((ph_area, int(ph.top), int(ph.left), area))

    if not candidates:
        return None

    # Pick the primary content area: topmost then leftmost among the
    # largest placeholders (within 10% of the max area).
    max_area = max(c[0] for c in candidates)
    threshold = int(max_area * 0.9)
    large = [(top, left, ca) for (a, top, left, ca) in candidates if a >= threshold]
    large.sort()  # topmost, then leftmost
    return large[0][2]


def _sidebar_content_area(slide_layout: SlideLayout) -> ContentArea | None:
    """Extract the secondary (sidebar) content area from a split slide layout.

    Returns the ContentArea for the right-side placeholder in layouts like
    2/3, 3/4, 1/2. Returns None when the layout has no secondary area.
    """
    from .spec import ContentArea

    content_y_threshold = 1600000
    footer_y = int(Layout.FOOTER_LINE_Y)

    # Collect (left, top, width) for content-region placeholders
    candidates: list[tuple[int, int, int]] = []
    for ph in slide_layout.placeholders:  # type: ignore[union-attr]
        top: int = int(ph.top)  # type: ignore[arg-type]
        if top < content_y_threshold:
            continue
        candidates.append((int(ph.left), top, int(ph.width)))  # type: ignore[arg-type]

    if len(candidates) < 2:
        return None

    candidates.sort()
    x, y, w = candidates[1]
    return ContentArea(x=x, y=y, width=w, height=int(footer_y - y))


def _json_out(obj: object) -> None:
    """Serialize to JSON, handling dataclass .to_dict()."""

    payload: object

    if _has_to_dict(obj):
        payload = obj.to_dict()
    elif _is_object_list(obj):
        items: list[object] = []
        for x_obj in obj:
            if _has_to_dict(x_obj):
                items.append(x_obj.to_dict())
            else:
                items.append(x_obj)
        payload = items
    else:
        payload = obj

    print(json.dumps(payload, indent=2, ensure_ascii=False, default=str))


def _shape_text(shape: BaseShape) -> str | None:
    """Current text from a shape."""
    if not _is_text_shape(shape):
        return None
    return shape.text_frame.text.replace("\x0b", "\\n").replace("\n", "\\n")


# ============================================================================
# TEXT PARSING (for edit/batch commands)
# ============================================================================


RunOverridesMap = dict[str, object]
TextRun = tuple[str, RunOverridesMap]


def _coerce_overrides(value: object) -> RunOverridesMap:
    if _is_str_object_dict(value):
        return dict(value)
    return {}


def _parse_text_arg(text: object) -> list[TextRun]:
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


def _expand_newlines(result: list[TextRun], text: str, opts: RunOverridesMap) -> None:
    """Split text on newlines, inserting line break markers."""
    parts = text.split("\n")
    for i, part in enumerate(parts):
        if i > 0:
            result.append(("\n", {}))
        if part:
            result.append((part, opts))


def _is_paragraphs_format(text_arg: object) -> bool:
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


def _parse_paragraphs_arg(text_arg: object) -> list[ParagraphSpec]:
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


def _write_to_shape(shape: Shape, text_arg: object) -> None:
    """Write content to a shape, dispatching based on format."""
    from .editor import add_line_break, add_run, snapshot_defaults, write_paragraphs

    if _is_paragraphs_format(text_arg):
        write_paragraphs(shape, _parse_paragraphs_arg(text_arg))
        return

    runs = _parse_text_arg(text_arg)
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


def _get_text_limit(key: str, default: int) -> int:
    try:
        limits = TEMPLATE_CONFIG.section("text_limits")
    except KeyError:
        return default

    raw = limits.get(key)
    if raw is None:
        return default

    try:
        return int(raw)
    except (TypeError, ValueError):
        return default


_SIDEBAR_MIN_PT = 8  # never shrink below this


def _warn_sidebar_overflow(
    paragraphs: list[Any],
    area: ContentArea,
    metrics: TextMetrics,
) -> None:
    """Warn when sidebar content exceeds available height."""
    total = _sidebar_height(paragraphs, area, metrics)
    if total > area.height:
        overflow_in = (total - area.height) / EMU_PER_INCH
        print(
            f"  WARNING: sidebar content overflows by ~{overflow_in:.1f}in. "
            f"Shorten text, reduce paragraphs, or set sidebar_shrink: true.",
            file=sys.stderr,
        )


def _sidebar_height(
    paragraphs: list[Any],
    area: ContentArea,
    metrics: TextMetrics,
) -> int:
    """Estimate total sidebar height in EMU."""
    PARA_GAP_EMU = int(4 * EMU_PER_PT)
    total = 0
    for para in paragraphs:
        font: str = str(para.font or Fonts.BODY)
        size_pt: int = int(para.size_pt or FontSizes.DEFAULT)
        total += metrics.text_height(str(para.text), area.width, font, size_pt) + PARA_GAP_EMU
    return total


def _shrink_sidebar_to_fit(
    paragraphs: list[Any],
    area: ContentArea,
    metrics: TextMetrics,
) -> None:
    """Proportionally shrink sidebar font sizes until content fits the area.

    Modifies paragraph objects in-place. Warns if content still overflows
    at the minimum font size.
    """
    total = _sidebar_height(paragraphs, area, metrics)
    if total <= area.height:
        return  # fits already

    # Compute scale factor and apply proportionally
    scale = area.height / total
    original_sizes: list[int] = []
    for para in paragraphs:
        orig = int(para.size_pt or FontSizes.DEFAULT)
        original_sizes.append(orig)
        shrunk = max(_SIDEBAR_MIN_PT, int(orig * scale))
        para.size_pt = shrunk

    # Re-check (rounding / min clamp may still overflow)
    total = _sidebar_height(paragraphs, area, metrics)
    if total > area.height:
        overflow_in = (total - area.height) / EMU_PER_INCH
        print(
            f"  WARNING: sidebar content still overflows by ~{overflow_in:.1f}in after "
            f"shrinking fonts (min {_SIDEBAR_MIN_PT}pt). Shorten text or reduce paragraphs.",
            file=sys.stderr,
        )
    else:
        reduced = [
            f"{orig}→{int(para.size_pt or orig)}pt"
            for para, orig in zip(paragraphs, original_sizes)
            if int(para.size_pt or orig) != orig
        ]
        if reduced:
            print(
                f"  sidebar: shrunk fonts to fit ({reduced[0].split('→')[1].rstrip('pt')}pt body)"
            )


def _warn_placeholder_text_limits(slide: Slide, shape: Shape) -> None:
    """Warn (to stderr) when placeholder text likely wraps beyond configured max lines."""
    if not shape.is_placeholder:
        return

    ph_idx = shape.placeholder_format.idx
    if ph_idx not in {0, 1}:
        return

    text = str(shape.text_frame.text or "").strip()
    if not text:
        return

    is_title_slide = "title" in (slide.slide_layout.name or "").lower()

    if ph_idx == 0:
        max_lines = _get_text_limit("title_max_lines", 2)
        font = Fonts.HEADLINE
        size_pt = int(FontSizes.TITLE)
        label = "title"
    else:
        if is_title_slide:
            max_lines = _get_text_limit("title_slide_subtitle_max_lines", 1)
        else:
            max_lines = _get_text_limit("subtitle_max_lines", 1)
        font = Fonts.HEADLINE
        size_pt = int(FontSizes.SUBTITLE)
        label = "subtitle"

    width_emu = int(shape.width)
    metrics = TextMetrics()
    needed = metrics.lines_needed(text, width_emu, font, size_pt)

    if needed > max_lines:
        print(
            f"Warning: {label} text likely wraps to ~{needed} lines (max {max_lines}) in placeholder '{shape.name}'. "
            "Consider shortening.",
            file=sys.stderr,
        )


def _text_preview(text_arg: object) -> str:
    """Compact preview for before/after display."""

    if _is_paragraphs_format(text_arg):
        paras = _parse_paragraphs_arg(text_arg)
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

    runs = _parse_text_arg(text_arg)
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


# ============================================================================
# INSPECT COMMANDS
# ============================================================================


def _classify_shape_type(shape: BaseShape) -> str:
    """Classify shape for display: placeholder, chart, image, table, group, text, connector, decorative."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    if shape.is_placeholder:
        return "placeholder"

    st = shape.shape_type
    if st == MSO_SHAPE_TYPE.CHART:
        return "chart"
    if st == MSO_SHAPE_TYPE.PICTURE:
        return "image"
    if st == MSO_SHAPE_TYPE.TABLE:
        return "table"
    if st == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if st in (MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.FREEFORM):
        return "line"
    if _is_text_shape(shape) and shape.text_frame.text.strip():
        return "text"
    return "shape"


def _shape_word_count(shape: BaseShape) -> int:
    """Count words in a shape's text frame."""
    if not _is_text_shape(shape):
        return 0
    text = shape.text_frame.text.strip()
    return len(text.split()) if text else 0


def _slide_word_count(slide: Slide) -> int:
    """Total words across all text-bearing shapes on a slide."""
    total = 0
    for shape in slide.shapes:
        total += _shape_word_count(shape)
    return total


def cmd_show(args: _ShowArgs) -> int:
    """Progressive drill-down: file → slide list, slide → shapes, shape → detail."""
    from .inspect_pptx import (
        get_slide_comments,
        inspect_chart,
        inspect_shape,
        inspect_slide,
        list_slides,
    )

    prs = _open(args.file)

    # Level 0: slide list
    if args.slide is None:
        total = len(prs.slides)
        print(f"\n  {args.file}  ({total} slide{'s' if total != 1 else ''})\n")
        for entry in list_slides(prs):
            slide = prs.slides[entry["slide"] - 1]
            layout_name = slide.slide_layout.name
            shapes_n = len(slide.shapes)
            words = _slide_word_count(slide)
            title = entry["title"][:60] or "(no title)"
            comments = get_slide_comments(slide)
            cm_str = f"  {len(comments)}cm" if comments else ""
            print(
                f"  {entry['slide']:3d}  {title:62s}  [{layout_name}]  {shapes_n}sh  {words}w{cm_str}"
            )
        print()
        return 0

    # Level 1: shapes on a slide
    slide = _get_slide(prs, args.slide)

    if args.shape is None:
        # Header
        title = "(no title)"
        subtitle = ""

        title_shape = slide.shapes.title
        if title_shape is not None:
            title = title_shape.text.replace("\x0b", " | ").replace("\n", " | ")

        for s in slide.shapes:
            if _is_text_shape(s) and s.is_placeholder and s.placeholder_format.idx == 1:
                subtitle = s.text.replace("\x0b", " | ").replace("\n", " | ")
                break

        layout = slide.slide_layout
        layout_phs: list[str] = []
        for ph in layout.placeholders:
            layout_phs.append(f"ph{ph.placeholder_format.idx}:{ph.name}")

        print(f"\n  Slide {args.slide}: {title}")
        if subtitle:
            print(f"  {subtitle}")
        print(
            f"  Layout: \"{layout.name}\" → {', '.join(layout_phs) if layout_phs else '(no placeholders)'}"
        )
        print()

        # All shapes, sorted by position (inspect_slide already sorts by top, left)
        shapes = inspect_slide(slide)
        for s in shapes:
            real_shape = list(slide.shapes)[s.index]
            cat = _classify_shape_type(real_shape)

            ph_str = f"ph{s.placeholder_idx}" if s.placeholder_idx is not None else ""
            type_str = f"{cat}" if not ph_str else f"{ph_str}"

            text = ""
            if s.text_preview:
                preview = s.text_preview[:60]
                if len(s.text_preview) > 60:
                    preview += "…"
                text = f"  «{preview}»"

            print(
                f"  [{s.index:2d}] {type_str:8s}  {s.left:5.2f},{s.top:5.2f}  {s.width:5.2f}x{s.height:5.2f}  {s.name:30s}{text}"
            )

        # Comments
        comments = get_slide_comments(slide)
        if comments:
            print(f"\n  Comments ({len(comments)}):")
            for cm in comments:
                author = cm.author
                preview = cm.text[:120]
                if len(cm.text) > 120:
                    preview += "…"
                print(f"    [{author}] {preview}")

        print()
        return 0

    # Level 2: shape detail (JSON)
    shape = _find_shape(slide, args.shape)
    if shape.has_chart if hasattr(shape, "has_chart") else False:
        _json_out(inspect_chart(shape))
    else:
        _json_out(inspect_shape(shape))
    return 0


def cmd_list(args: _FileArgs) -> int:
    from .inspect_pptx import list_slides

    prs = _open(args.file)
    for entry in list_slides(prs):
        print(f"  {entry['slide']:3d}  {entry['title']}")
    return 0


def cmd_summary(args: _SlideArgs) -> int:
    from .inspect_pptx import summarize_slide

    prs = _open(args.file)
    slide = _get_slide(prs, args.slide)
    _json_out(summarize_slide(slide))
    return 0


def cmd_slide(args: _SlideArgs) -> int:
    from .inspect_pptx import inspect_slide

    prs = _open(args.file)
    slide = _get_slide(prs, args.slide)
    shapes = inspect_slide(slide)
    for s in shapes:
        ph = f" ph={s.placeholder_idx}" if s.placeholder_idx is not None else ""
        fill = f" fill={s.fill.type}" if s.fill.type != "inherited" else ""
        text = f"  «{s.text_preview}»" if s.text_preview else ""
        print(
            f"  [{s.index:2d}] {s.name:30s}  {s.left:6.2f},{s.top:6.2f}  {s.width:5.2f}x{s.height:5.2f}{ph}{fill}{text}"
        )
    return 0


def cmd_shape(args: _XmlArgs) -> int:
    from .inspect_pptx import inspect_shape

    prs = _open(args.file)
    slide = _get_slide(prs, args.slide)
    shape = _find_shape(slide, args.shape)
    _json_out(inspect_shape(shape))
    return 0


def cmd_chart(args: _XmlArgs) -> int:
    from .inspect_pptx import inspect_chart

    prs = _open(args.file)
    slide = _get_slide(prs, args.slide)
    shape = _find_shape(slide, args.shape)
    _json_out(inspect_chart(shape))
    return 0


def cmd_layout(args: _SlideArgs) -> int:
    from .inspect_pptx import inspect_layout

    prs = _open(args.file)
    slide = _get_slide(prs, args.slide)
    _json_out(inspect_layout(slide.slide_layout))
    return 0


def cmd_layouts(args: _FileArgs) -> int:
    from pptx.enum.shapes import PP_PLACEHOLDER
    from pptx.shapes.placeholder import LayoutPlaceholder

    prs = _open(args.file)

    # Boundary between "structural" placeholders (title, subtitle, tracker) and
    # the slide's main content zone.
    content_y_threshold = int(Layout.CONTENT_START_Y)
    footer_y = int(Layout.FOOTER_LINE_Y)

    for layout in prs.slide_layouts:
        structural: list[str] = []
        # (left_emu, width_in, height_avail_in, name)
        content_phs: list[tuple[int, float, float, str]] = []

        placeholders = cast(Iterable[LayoutPlaceholder], layout.placeholders)
        phs: list[LayoutPlaceholder] = sorted(
            placeholders,
            key=lambda p: (int(p.top), int(p.left)),
        )
        for ph in phs:
            pf = ph.placeholder_format
            w_in = ph.width.inches

            if pf.type == PP_PLACEHOLDER.TITLE:
                structural.append(f"title ({w_in:.1f}in)")
            elif pf.type == PP_PLACEHOLDER.SUBTITLE:
                structural.append(f"subtitle ({w_in:.1f}in)")
            elif pf.type == PP_PLACEHOLDER.PICTURE:
                structural.append(f"image: {ph.name} ({w_in:.1f}×{ph.height.inches:.1f}in)")
            elif int(ph.top) < content_y_threshold:
                # Above content zone — tracker, doc type, etc.
                structural.append(f"{ph.name} ({w_in:.1f}in)")
            else:
                h_avail = Emu(footer_y - int(ph.top)).inches
                content_phs.append((int(ph.left), w_in, h_avail, ph.name))

        # Sort content areas left-to-right, label primary/secondary
        content_phs.sort(key=lambda p: p[0])
        areas: list[str] = []
        for idx, (_x, w, h, _name) in enumerate(content_phs):
            label = "primary" if idx == 0 else "secondary"
            areas.append(f"{label} {w:.1f}×{h:.1f}in")

        print(f"  {layout.name}")
        if structural:
            print(f"    placeholders: {', '.join(structural)}")
        if areas:
            print(f"    content areas: {', '.join(areas)}")
        else:
            print("    content areas: (none)")
    return 0


def cmd_theme(args: _FileArgs) -> int:
    from .inspect_pptx import resolve_theme_colors

    prs = _open(args.file)
    colors = resolve_theme_colors(prs)
    for name, hex_val in sorted(colors.items()):
        if not name.startswith("_"):
            print(f"  {name:12s}  {hex_val}")
    return 0


def cmd_color(args: _ColorArgs) -> int:
    from .inspect_pptx import identify_color

    prs = _open(args.file)
    result = identify_color(prs, args.hex)
    if result:
        print(result)
    else:
        print(f"No theme match for {args.hex}")
    return 0


def cmd_xml(args: _XmlArgs) -> int:
    from .xml_helpers import dump_xml

    prs = _open(args.file)
    slide = _get_slide(prs, args.slide)
    shape = _find_shape(slide, args.shape)
    print(dump_xml(shape._element))
    return 0


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


def _parse_slide_selection(selection: str | None, total: int) -> list[int]:
    """Parse a slide selection string (e.g. "1,3-5") into 1-indexed slide numbers."""
    if not selection:
        return list(range(1, total + 1))

    result: list[int] = []
    seen: set[int] = set()

    for part in selection.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start = int(a)
            end = int(b)
            if start > end:
                start, end = end, start
            nums = range(start, end + 1)
        else:
            nums = [int(part)]

        for n in nums:
            if n < 1 or n > total:
                raise ValueError(f"slide {n} out of range (1-{total})")
            if n not in seen:
                seen.add(n)
                result.append(n)

    return result


def _move_last_slide_to(prs: PresentationObj, at_pos: int) -> None:
    """Move the last slide in *prs* to position *at_pos* (1-indexed)."""
    sldIdLst = prs.slides.element
    sld_ids = sldIdLst.sldId_lst
    if not sld_ids:
        raise ValueError("presentation has no slide ID list")

    last = sld_ids[-1]
    sldIdLst.remove(last)

    insert_idx = at_pos - 1
    if insert_idx < len(sldIdLst.sldId_lst):
        sldIdLst.insert(insert_idx, last)
    else:
        sldIdLst.append(last)


_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_HYPERLINK_RELTYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)


def _collect_relationship_ids(slide: Slide) -> set[str]:
    """Return all r:id values referenced by shapes on this slide."""
    ids: set[str] = set()
    spTree = slide.shapes.element
    for el in spTree.iter():
        for key, val in el.attrib.items():
            if str(key).startswith(f"{{{_NS_R}}}"):
                ids.add(str(val))
    return ids


def _check_and_copy_relationships(src: Slide, dst: Slide) -> None:
    """Copy external hyperlink rels from src to dst, reject embedded content.

    Hyperlinks are external relationships (just a URL string) and can be
    safely re-created on the destination slide. Embedded content (images,
    charts, media) requires OPC-level blob copying which is not yet supported.
    """
    ref_ids = _collect_relationship_ids(src)
    if not ref_ids:
        return

    # Build a rId remapping: src rId → dst rId
    remap: dict[str, str] = {}

    for rid in ref_ids:
        try:
            rel = src.part.rels[rid]
        except KeyError:
            continue

        if rel.is_external and rel.reltype == _HYPERLINK_RELTYPE:
            # Add the hyperlink as a new external relationship on dst
            new_rid = dst.part.rels.get_or_add_ext_rel(_HYPERLINK_RELTYPE, rel.target_ref)
            remap[rid] = new_rid
        else:
            raise ValueError(
                f"source slide contains embedded relationship rId={rid} "
                f"(type={rel.reltype}); pptx insert supports text and hyperlinks only"
            )

    # Remap r:id attributes in the destination shapes
    if remap:
        dst_spTree = dst.shapes.element
        for el in dst_spTree.iter():
            for key in list(el.attrib):
                if str(key).startswith(f"{{{_NS_R}}}"):
                    old_id = str(el.attrib[key])
                    if old_id in remap:
                        el.attrib[key] = remap[old_id]


def _replace_slide_shapes(dst: Slide, src: Slide) -> None:
    """Replace dst slide shapes with a deep-copy of src slide shapes."""
    dst_spTree = dst.shapes.element
    src_spTree = src.shapes.element

    # Keep spTree group headers (nvGrpSpPr + grpSpPr). Remove everything else.
    for child in list(dst_spTree)[2:]:
        dst_spTree.remove(child)

    for child in list(src_spTree)[2:]:
        dst_spTree.append(deepcopy(child))


def cmd_add_slide(args: _AddSlideArgs) -> int:
    prs = _open(args.file)
    layout = _find_layout(prs, args.layout)
    total_before = len(prs.slides)
    out_path = args.out or args.file

    slide = prs.slides.add_slide(layout)

    if args.at is not None:
        at_pos = args.at
        if at_pos < 1 or at_pos > total_before + 1:
            print(f"Error: --at {at_pos} out of range (1-{total_before + 1})", file=sys.stderr)
            return 1

        # Reorder slide by moving its <p:sldId> entry in the presentation XML.
        sldIdLst = prs.slides.element
        sld_ids = sldIdLst.sldId_lst
        if not sld_ids:
            print("Error: presentation has no slide ID list", file=sys.stderr)
            return 1

        new_sldId = sld_ids[-1]
        sldIdLst.remove(new_sldId)

        insert_idx = at_pos - 1
        if insert_idx < len(sldIdLst.sldId_lst):
            sldIdLst.insert(insert_idx, new_sldId)
        else:
            sldIdLst.append(new_sldId)

        final_pos = at_pos
    else:
        final_pos = total_before + 1

    phs: list[str] = [f"{ph.placeholder_format.idx}:{ph.name}" for ph in slide.placeholders]
    print(f'Added slide {final_pos} from layout "{layout.name}"')
    print(f"  Placeholders: {', '.join(phs) if phs else '(none)'}")
    print(f"  Total slides: {total_before + 1}")

    prs.save(out_path)
    print(f"Saved → {out_path}")
    return 0


def cmd_delete_slide(args: _DeleteSlideArgs) -> int:
    if not args.confirm:
        print("Error: delete-slide requires --confirm flag", file=sys.stderr)
        return 1

    prs = _open(args.file)
    total = len(prs.slides)
    out_path = args.out or args.file

    if total <= 1:
        print("Error: cannot delete the only slide in the presentation", file=sys.stderr)
        return 1

    slide = _get_slide(prs, args.slide)

    title = "(no title)"
    title_shape = slide.shapes.title
    if title_shape is not None:
        title = title_shape.text.replace("\x0b", " | ").replace("\n", " | ")[:60]

    print(f'Deleting slide {args.slide}: "{title}" ({len(slide.shapes)} shapes)')

    # Find relationship id (rId) for this slide part
    slide_part = slide.part
    rId: str | None = None
    for rel_key in prs.part.rels:
        rel = prs.part.rels[rel_key]
        if rel.target_part is slide_part:
            rId = rel_key
            break

    if rId is None:
        print("Error: could not find slide relationship", file=sys.stderr)
        return 1

    # Remove the <p:sldId> element referencing this slide
    sldIdLst = prs.slides.element
    for sldId in list(sldIdLst.sldId_lst):
        if sldId.rId == rId:
            sldIdLst.remove(sldId)
            break

    prs.part.rels.pop(rId)

    print(f"  Remaining slides: {total - 1}")
    prs.save(out_path)
    print(f"Saved → {out_path}")
    return 0


def cmd_delete_shape(args: _DeleteShapeArgs) -> int:
    prs = _open(args.file)
    slide = _get_slide(prs, args.slide)
    shape = _find_shape(slide, args.shape)
    out_path = args.out or args.file

    text_preview = ""
    if _is_text_shape(shape):
        text_preview = shape.text_frame.text.replace("\x0b", " ").replace("\n", " ")[:50]

    print(f'Deleting shape: "{shape.name}"')
    if text_preview:
        print(f"  Text: {text_preview}")

    slide.shapes.element.remove(shape.element)

    prs.save(out_path)
    print(f"Saved → {out_path}")
    return 0


def cmd_insert(args: _InsertArgs) -> int:
    """Insert slides from another PPTX into this presentation.

    Intended workflow:
        pptx generate spec.yaml -t template.pptx -o /tmp/table.pptx
        pptx insert  deck.pptx /tmp/table.pptx --at 5

    Notes:
    - Supports text and hyperlink slides. Embedded content (images/charts/media)
      is not yet supported.
    - Matches the source slide's layout name in the destination deck, falling
      back to Default (with a warning) when no match is found.  Native
      PowerPoint would import the source layout/master into the destination
      instead; we don't do that because it requires deep OPC-level copying
      of theme, fonts, and background assets.
    """

    prs = _open(args.file)
    src_prs = _open(args.source)
    out_path = args.out or args.file

    total_before = len(prs.slides)
    src_total = len(src_prs.slides)

    try:
        selected = _parse_slide_selection(args.slides, src_total)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    if args.at is None:
        at_pos = total_before + 1
    else:
        at_pos = args.at
        if at_pos < 1 or at_pos > total_before + 1:
            print(f"Error: --at {at_pos} out of range (1-{total_before + 1})", file=sys.stderr)
            return 1

    insert_pos = at_pos
    for slide_num in selected:
        src_slide = _get_slide(src_prs, slide_num)

        # Match the source slide's layout by name in the destination deck.
        # Fall back to "Default" if no match, then best-effort.
        #
        # NOTE: Native PowerPoint imports the source layout (and its master)
        # into the destination when no matching layout exists, creating
        # numbered duplicates like "1_LayoutName".  We don't do that —
        # importing a full layout/master with its theme, fonts, and
        # backgrounds is a deep OPC-level operation.  Instead we fall back
        # to "Default" and warn.  Shapes are still copied faithfully.
        dst_layout: SlideLayout | None = None
        src_layout_name = getattr(src_slide.slide_layout, "name", "")
        if src_layout_name:
            dst_layout = _try_find_layout(prs, src_layout_name)
        if dst_layout is None and src_layout_name:
            print(
                f"  WARNING: slide {slide_num} layout '{src_layout_name}' not found "
                f"in destination; falling back to Default. "
                f"(Native PowerPoint would import the source layout instead.)"
            )
        if dst_layout is None:
            dst_layout = _try_find_layout(prs, "default")
        if dst_layout is None:
            dst_layout = _find_layout(prs, "default", fallback=True)
            print(
                f"  WARNING: slide {slide_num} — 'Default' layout not found either; "
                f"using '{dst_layout.name}'."
            )

        dst_slide = prs.slides.add_slide(dst_layout)
        _replace_slide_shapes(dst_slide, src_slide)

        # Copy hyperlink relationships from source to destination (remapping rIds).
        # Raises ValueError for embedded content (images/charts) which isn't supported yet.
        try:
            _check_and_copy_relationships(src_slide, dst_slide)
        except ValueError as e:
            print(f"Error: slide {slide_num}: {e}", file=sys.stderr)
            return 1

        # Reorder to requested insertion position.
        _move_last_slide_to(prs, insert_pos)
        insert_pos += 1

    prs.save(out_path)
    print(
        f"Inserted {len(selected)} slide(s) from {args.source} into {args.file} at position {at_pos} → {out_path}"
    )
    return 0


# ============================================================================
# RENDER COMMANDS
# ============================================================================


def cmd_render(args: _RenderArgs) -> int:
    from .screenshot import render_slides

    prs = _open(args.file)
    total = len(prs.slides)
    try:
        slide_nums = _parse_slide_selection(args.slides, total)
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


def cmd_crop(args: _CropArgs) -> int:
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


def cmd_charts(args: _ChartsArgs) -> int:
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
            module_path=args.module_path,
        )
    except (FileNotFoundError, ImportError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Saved → {output_path}")
    return 0


# ============================================================================
# GENERATE COMMANDS (YAML → PPTX table pipeline)
# ============================================================================


YamlDict = dict[str, object]


def load_yaml(path: str) -> YamlDict:
    """Load YAML from disk."""
    with open(path, encoding="utf-8") as handle:
        parsed: object = yaml.safe_load(handle)

    # YAML root must be a mapping for our spec format. If it's not, treat it as
    # empty so downstream validation can surface a friendly error.
    if _is_str_object_dict(parsed):
        return parsed

    return {}


def _to_int(value: object) -> int:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(value)
    return int(str(value))


def _parse_content_area(data: YamlDict, layout_override: str | None = None) -> ContentArea:
    if "content_area" in data:
        area_obj = data.get("content_area")
        area: dict[str, object] = area_obj if _is_str_object_dict(area_obj) else {}

        return ContentArea(
            x=_to_int(area.get("x", 0)),
            y=_to_int(area.get("y", 0)),
            width=_to_int(area.get("width", 0)),
            height=_to_int(area.get("height", 0)),
        )

    # Content-area layout (not slide-master layout)
    layout_obj = data.get("content_layout")
    if layout_obj is None:
        layout_obj = data.get("layout")

    if isinstance(layout_obj, str) and layout_obj:
        layout = layout_obj
    elif layout_override is not None:
        layout = layout_override
    else:
        layout = "default"

    return ContentArea.from_layout(layout)


def _parse_options(table: dict[str, object]) -> SolveOptions:
    fonts_obj = table.get("fonts")
    fonts: dict[str, object] = fonts_obj if _is_str_object_dict(fonts_obj) else {}

    padding_obj = table.get("padding")
    padding: dict[str, object] = padding_obj if _is_str_object_dict(padding_obj) else {}

    # Default to template config font size (16pt) instead of hardcoded 12pt
    body_raw = fonts.get("body", FontSizes.TABLE_BODY)
    header_raw = fonts.get("header", FontSizes.TABLE_HEADER)

    default_pad = int(TableDefaults.CELL_PADDING)

    return SolveOptions(
        body_font_pt=_to_int(body_raw),
        header_font_pt=_to_int(header_raw),
        min_font_pt=_to_int(fonts.get("min", 10)),
        max_font_pt=_to_int(fonts.get("max", 16)),
        pad_top=_to_int(padding.get("top", default_pad)),
        pad_bottom=_to_int(padding.get("bottom", default_pad)),
    )


def parse_spec(
    data: YamlDict,
    layout_override: str | None = None,
) -> tuple[TableSpec, ContentArea, SolveOptions, bool]:
    """Parse YAML dict into structured objects."""

    table_obj = data.get("table")
    table: dict[str, object] = table_obj if _is_str_object_dict(table_obj) else {}

    spec = TableSpec.from_dict(data)
    area = _parse_content_area(data, layout_override=layout_override)
    options = _parse_options(table)

    placeholders_raw = table.get("placeholders", True)
    placeholders = bool(placeholders_raw)

    return spec, area, options, placeholders


def validate_spec(data: YamlDict) -> tuple[list[str], list[str]]:
    """Return (errors, warnings)."""
    errors: list[str] = []
    warnings: list[str] = []

    if not data.get("title"):
        warnings.append("Missing 'title' — slide will have no title text")

    table_obj = data.get("table")
    if not _is_str_object_dict(table_obj) or not table_obj:
        # No table section — metadata-only slide (title, section, quote, etc.).
        # This is valid; generate will create the slide with the correct layout
        # and fill placeholders without rendering a table.
        return errors, warnings

    table = table_obj

    raw_groups_raw: object = table.get("row_groups")
    rows_obj: object = table.get("rows")
    cols_obj: object = table.get("cols")

    if raw_groups_raw:
        if cols_obj is None:
            errors.append("table.cols is required")

        if not _is_object_list(raw_groups_raw) or not raw_groups_raw:
            errors.append("table.row_groups must be a non-empty list")
    else:
        if rows_obj is None:
            errors.append("table.rows is required")
        if cols_obj is None:
            errors.append("table.cols is required")

    # Basic numeric checks (best-effort; actual dimension validation happens in TableSpec.from_dict)
    if isinstance(rows_obj, int) and rows_obj <= 0:
        errors.append("table.rows must be > 0")
    if isinstance(cols_obj, int) and cols_obj <= 0:
        errors.append("table.cols must be > 0")

    if "has_col_header" not in table:
        warnings.append("table.has_col_header not specified (defaulting to true)")
    if not raw_groups_raw and "has_row_header" not in table:
        warnings.append("table.has_row_header not specified (defaulting to false)")

    # Parse normalized spec (body-only dims) so validation stays aligned with renderer/solver.
    try:
        spec = TableSpec.from_dict(data)  # handles row_groups + header-inclusive rows/cols
    except ValueError as e:
        errors.append(str(e))
        return errors, warnings

    if not spec.has_col_header:
        warnings.append("Column headers disabled (table.has_col_header = false)")
    if not spec.has_row_header and not raw_groups_raw:
        warnings.append("Row headers disabled (table.has_row_header = false)")

    body_rows = spec.num_rows
    body_cols = spec.num_cols

    # --- headers ---

    raw_col_headers = table.get("col_headers")
    raw_row_headers = table.get("row_headers")
    has_row_header_col = spec.has_row_header

    if spec.has_col_header and not raw_col_headers:
        warnings.append("Column headers missing; placeholders will be generated")
    # Auto-extract: when has_row_header but no row_headers list, the parser
    # extracts them from the first column of cells.  Only warn if cells are
    # also missing (truly no source for row headers).
    raw_cells: object = table.get("cells")
    if has_row_header_col and not raw_row_headers and not raw_groups_raw:
        cells_can_provide = _cells_can_provide_row_headers(raw_cells)
        if not cells_can_provide:
            warnings.append("Row headers missing; placeholders will be generated")

    # col_headers should match *body* columns.
    # Convenience: if has_row_header and row_header_col_header not provided, allow one extra.
    if _is_object_list(raw_col_headers):
        expected = body_cols
        allow = expected + (
            1 if (has_row_header_col and not table.get("row_header_col_header")) else 0
        )
        if len(raw_col_headers) not in {expected, allow}:
            warnings.append(
                f"col_headers length ({len(raw_col_headers)}) does not match expected body columns ({expected})"
            )

    # row_headers should match *body* rows (except grouped mode where row headers are groups).
    if (
        not spec.is_grouped
        and _is_object_list(raw_row_headers)
        and len(raw_row_headers) != body_rows
    ):
        warnings.append(
            f"row_headers length ({len(raw_row_headers)}) does not match expected body rows ({body_rows})"
        )

    # --- column widths ---

    cw: object = table.get("column_widths")

    if isinstance(cw, str) and cw.lower() not in {"equal"}:
        warnings.append(
            f"column_widths '{cw}' is not recognized (use 'equal' or a list of numbers)"
        )
    elif _is_object_list(cw):
        total_cols = body_cols + (1 if has_row_header_col else 0)
        if has_row_header_col:
            if len(cw) not in {body_cols, total_cols}:
                warnings.append(
                    f"column_widths length ({len(cw)}) does not match expected columns ({body_cols} body, or {total_cols} including row header)"
                )
        else:
            if len(cw) != body_cols:
                warnings.append(
                    f"column_widths length ({len(cw)}) does not match expected body columns ({body_cols})"
                )

    # --- cells grid ---

    cells_obj = table.get("cells")
    # When auto-extracting row headers from cells, the first column is consumed
    # by the parser — account for that in width checks.
    auto_extract_row_hdr = (
        has_row_header_col
        and not raw_row_headers
        and not raw_groups_raw
        and _cells_can_provide_row_headers(raw_cells)
    )
    effective_body_cols = body_cols + (1 if auto_extract_row_hdr else 0)

    if _is_object_list(cells_obj):
        if len(cells_obj) > body_rows:
            warnings.append("cells has more rows than expected body rows")

        for idx, row_obj in enumerate(cells_obj):
            if not _is_object_list(row_obj):
                warnings.append(f"cells row {idx + 1} is not a list")
                continue
            if len(row_obj) > effective_body_cols:
                warnings.append(f"cells row {idx + 1} has more columns than expected body columns")

    # --- layout keys ---

    # Validate content_layout by trying to construct it
    cl_raw = data.get("content_layout") or data.get("layout")
    if cl_raw is not None and isinstance(cl_raw, str) and cl_raw:
        try:
            ContentArea.from_layout(cl_raw)
        except ValueError:
            warnings.append(f"content_layout '{cl_raw}' is not recognized")

    content_area_obj = data.get("content_area")
    if content_area_obj:
        if not _is_str_object_dict(content_area_obj):
            errors.append("content_area must be a mapping when set")
        else:
            for key in ("x", "y", "width", "height"):
                if key not in content_area_obj:
                    errors.append(f"content_area.{key} is required when content_area is set")

    # --- fonts / bullets ---

    fonts_obj = table.get("fonts")
    fonts: dict[str, object] = fonts_obj if _is_str_object_dict(fonts_obj) else {}

    body = _to_int(fonts.get("body", FontSizes.TABLE_BODY))
    header = _to_int(fonts.get("header", body))
    min_font = _to_int(fonts.get("min", 10))
    max_font = _to_int(fonts.get("max", 16))

    if header < body:
        warnings.append("header font is smaller than body font")
    if body < min_font or body > max_font:
        warnings.append("body font outside min/max bounds")

    body_default_lvl = table.get("body_default_lvl", 0)
    if not isinstance(body_default_lvl, int) or body_default_lvl < 0 or body_default_lvl > 8:
        warnings.append("body_default_lvl should be between 0 and 8")

    parse_bullets = table.get("parse_bullets", True)
    if not isinstance(parse_bullets, bool):
        warnings.append("parse_bullets should be true or false")

    # --- icons legend ---

    if spec.icons is not None and spec.icons.show_legend:
        legend_items = list(spec.icons.values.items())
        if len(legend_items) > 5:
            warnings.append(
                "icons.legend has more than 5 items; consider limiting legend to 5 for readability"
            )
        colors = [c for _, c in legend_items]
        if len(set(colors)) != len(colors):
            warnings.append(
                "icons.legend contains duplicate colors; legend should map 1:1 with meaning"
            )

    return errors, warnings


def preview_spec(data: YamlDict) -> str:
    """Generate a text preview of table structure."""

    table_obj = data.get("table")
    table: dict[str, object] = table_obj if _is_str_object_dict(table_obj) else {}

    rows = table.get("rows")
    cols = table.get("cols")
    has_col_header = bool(table.get("has_col_header", True))
    has_row_header = bool(table.get("has_row_header", False))

    lines: list[str] = ["=" * 60, "TABLE PREVIEW", "=" * 60]
    lines.append(f"Rows: {rows} | Cols: {cols}")
    lines.append(f"Column headers: {has_col_header} | Row headers: {has_row_header}")

    col_headers_obj = table.get("col_headers")
    col_headers: list[str] = (
        [str(v) for v in col_headers_obj] if _is_object_list(col_headers_obj) else []
    )
    if has_col_header and col_headers:
        lines.append("")
        lines.append("Column headers:")
        lines.append(" | ".join(col_headers))

    row_headers_obj = table.get("row_headers")
    row_headers: list[str] = (
        [str(v) for v in row_headers_obj] if _is_object_list(row_headers_obj) else []
    )
    if has_row_header and row_headers:
        lines.append("")
        lines.append("Row headers:")
        lines.append(", ".join(row_headers[:5]))
        if len(row_headers) > 5:
            lines.append(f"... and {len(row_headers) - 5} more")

    cells_obj = table.get("cells")
    cells: list[object] = cells_obj if _is_object_list(cells_obj) else []

    if cells:
        lines.append("")
        lines.append("Sample rows:")
        for row_obj in cells[:3]:
            if not _is_object_list(row_obj):
                continue

            preview_row: list[str] = []
            for cell in row_obj[:5]:
                cell_text = str(cell)
                if len(cell_text) > 24:
                    cell_text = cell_text[:21] + "..."
                preview_row.append(cell_text)

            lines.append(" | ".join(preview_row))

        if len(cells) > 3:
            lines.append(f"... and {len(cells) - 3} more rows")

    lines.append("")
    lines.append("=" * 60)
    return "\n".join(lines)


def _expand_inputs(inputs: list[str]) -> list[Path]:
    files: list[Path] = []
    for pattern in inputs:
        path = Path(pattern)
        if path.is_file():
            files.append(path)
        else:
            files.extend(Path(".").glob(pattern))
    return files


def _clear_content_area(slide: Slide, area: ContentArea) -> None:
    for shape in list(slide.shapes):
        if _boxes_intersect(
            (int(shape.left), int(shape.top), int(shape.width), int(shape.height)),
            (area.x, area.y, area.width, area.height),
        ):
            slide.shapes.element.remove(shape.element)


_Box = tuple[int, int, int, int]


def _boxes_intersect(box_a: _Box, box_b: _Box) -> bool:
    ax, ay, aw, ah = box_a
    bx, by, bw, bh = box_b
    return ax < bx + bw and ax + aw > bx and ay < by + bh and ay + ah > by


def _infer_layout_from_slide(slide: Slide) -> str:
    for shape in slide.shapes:
        if shape.top < Layout.CONTENT_START_Y:
            return "content"
    return "default"


def _delete_all_slides(prs: PresentationObj) -> None:
    """Remove all existing slides from a presentation."""
    # Must iterate in reverse to avoid index shifting
    # Access internal _sldIdLst to properly delete slides
    slides = prs.slides
    for i in range(len(slides) - 1, -1, -1):
        rId: str = slides._sldIdLst[i].rId  # type: ignore[union-attr]
        prs.part.drop_rel(rId)  # type: ignore[union-attr]
        del slides._sldIdLst[i]  # type: ignore[union-attr]


def _clear_body_placeholders(slide: Slide) -> None:
    """Remove unfilled body/content placeholders from a slide."""
    # Only keep these specific placeholder types (by idx):
    # 0 = title, 1 = subtitle
    # Tracker placeholder is identified by name, not idx (idx varies by layout)
    KEEP_PLACEHOLDER_INDICES = {0, 1}

    shapes_to_remove: list[BaseShape] = []
    for shape in slide.shapes:
        if not shape.is_placeholder:
            continue
        ph_idx = shape.placeholder_format.idx

        # Always keep title and subtitle
        if ph_idx in KEEP_PLACEHOLDER_INDICES:
            continue

        # Keep tracker placeholder (identified by name pattern)
        name_lower = shape.name.lower()
        if "tracker" in name_lower or "on-page" in name_lower:
            continue

        # Check if placeholder has meaningful content
        if hasattr(shape, "text_frame"):
            text: str = str(shape.text_frame.text).strip()  # type: ignore[union-attr]
            # Keep if it has real content (not empty, not placeholder prompts)
            if text and "click to" not in text.lower() and "master text" not in text.lower():
                continue

        # Remove this placeholder
        shapes_to_remove.append(shape)

    # Remove shapes by deleting their XML elements
    for shape in shapes_to_remove:
        sp = shape.element
        parent = sp.getparent()
        if parent is not None:
            parent.remove(sp)


def _hint_init() -> None:
    """Print a one-time hint about `pptx init` when no project config exists."""
    proj = _discover_project_dir()
    if proj is None:
        print(
            "Hint: run `pptx init` to set up a project template and config "
            f"in {PROJECT_DIR_NAME}/",
            file=sys.stderr,
        )


def cmd_generate(args: _GenerateArgs) -> int:
    """Generate PPTX for text-only tables."""
    _apply_config(args.config)

    input_files = _expand_inputs(args.input)
    if not input_files:
        print("No input files found", file=sys.stderr)
        return 1

    # Resolve template: explicit flag → auto-discovered → blank
    template_path = args.template
    if template_path is None:
        discovered_tpl = _discover_template()
        if discovered_tpl is not None:
            template_path = str(discovered_tpl)

    if template_path:
        prs = Presentation(template_path)
        # Remove template's content slides - keep only layouts/masters
        _delete_all_slides(prs)
    else:
        _hint_init()
        prs = Presentation()
        prs.slide_width = Emu(int(Layout.SLIDE_WIDTH))
        prs.slide_height = Emu(int(Layout.SLIDE_HEIGHT))

    metrics = TextMetrics()
    solver = ConstraintSolver(metrics)

    for path in input_files:
        data = load_yaml(str(path))
        errors, warnings = validate_spec(data)
        if errors:
            print(f"{path}:\n  - " + "\n  - ".join(errors), file=sys.stderr)
            return 1
        if warnings:
            print(f"{path}:\n  - " + "\n  - ".join(warnings))

        has_table = _is_str_object_dict(data.get("table")) and bool(data.get("table"))

        layout_override: str | None = None
        slide: Slide | None = None
        if args.slide_index is not None:
            if args.slide_index < 0 or args.slide_index >= len(prs.slides):
                print("slide_index out of range", file=sys.stderr)
                return 1
            slide = prs.slides[args.slide_index]
            if "content_layout" not in data and "layout" not in data:
                layout_override = _infer_layout_from_slide(slide)

        # ---- Resolve slide layout & create slide ----
        slide_layout_obj: SlideLayout | None = None
        if slide is None:
            slide_layout_name = str(data.get("slide_layout") or "Default")
            slide_layout_obj = _find_layout(prs, slide_layout_name, fallback=True)
            slide = prs.slides.add_slide(slide_layout_obj)

        if has_table:
            spec, area, options, placeholders = parse_spec(data, layout_override=layout_override)
            if placeholders:
                spec = fill_placeholders(spec)

            # Derive content area from the layout's primary content placeholder
            # unless the YAML explicitly overrides via content_area / content_layout.
            if slide_layout_obj is not None and "content_area" not in data:
                layout_area = _content_area_from_layout(slide_layout_obj)
                if layout_area is not None:
                    area = layout_area

            if args.slide_index is not None and not args.keep_existing:
                _clear_content_area(slide, area)

            layout, report = solver.solve(spec, area, options)
            print(f"{path}: {report.to_text(detail=args.detail)}")

            spTree = slide.shapes.element

            shape_id: int = TableDefaults.SHAPE_ID_START

            def next_shape_id() -> int:
                nonlocal shape_id
                shape_id += 1
                return shape_id

            renderer = TableRenderer(spTree, next_shape_id, slide_part=slide.part)
            renderer.render(spec, layout, area)

            # Sidebar: fill secondary content area with formatted paragraphs
            sidebar_raw = data.get("sidebar")
            if sidebar_raw is not None and slide_layout_obj is not None:
                sidebar_area = _sidebar_content_area(slide_layout_obj)
                if sidebar_area is not None:
                    from .content import Paragraph, normalize_cell

                    default_para = Paragraph(text="", lvl=0)
                    sidebar_paras = normalize_cell(sidebar_raw, default_para, parse_bullets=True)
                    if sidebar_paras:
                        if data.get("sidebar_shrink"):
                            _shrink_sidebar_to_fit(sidebar_paras, sidebar_area, metrics)
                        else:
                            _warn_sidebar_overflow(sidebar_paras, sidebar_area, metrics)
                        renderer.render_sidebar(sidebar_paras, sidebar_area)
                else:
                    print(
                        "  WARNING: sidebar content specified but layout has no secondary content area"
                    )
        else:
            print(f"{path}: metadata-only slide (no table)")

        fill_slide_metadata(slide, data)

        # Agent-friendly: warn when title/subtitle likely wrap beyond configured limits.
        for sh in slide.shapes:
            if isinstance(sh, Shape) and sh.has_text_frame:
                _warn_placeholder_text_limits(slide, sh)

        _clear_body_placeholders(slide)

    output_path = args.output or "output.pptx"
    prs.save(output_path)
    print(f"Saved: {output_path}")

    return 0


def _example_template_dir() -> Path:
    """Return path to the bundled example-template directory."""
    return Path(__file__).resolve().parent / "example-template"


def cmd_init(args: _InitArgs) -> int:
    """Initialise a .clean-slides/ project directory.

    Without --template: copies the bundled example template and config.
    With --template <file.pptx>: runs init-config to generate a config,
    then copies the template + generated config into .clean-slides/.
    """
    import shutil

    target = Path(args.output) if args.output else Path.cwd()
    project_dir = target / PROJECT_DIR_NAME

    if project_dir.exists():
        print(f"Already initialised: {project_dir}", file=sys.stderr)
        return 1

    project_dir.mkdir(parents=True)

    if args.template:
        # Copy the user's template
        src_tpl = Path(args.template)
        if not src_tpl.is_file():
            print(f"Template not found: {src_tpl}", file=sys.stderr)
            return 1
        dst_tpl = project_dir / _TEMPLATE_NAME
        shutil.copy2(src_tpl, dst_tpl)

        # Generate config via init-config
        dst_cfg = project_dir / _CONFIG_NAME

        class _FakeArgs:
            file = str(src_tpl)
            output = str(dst_cfg)

        rc = cmd_init_config(_FakeArgs())  # type: ignore[arg-type]
        if rc != 0:
            return rc
    else:
        # Copy bundled example
        example_dir = _example_template_dir()
        src_tpl = example_dir / "example-template.pptx"
        src_cfg = example_dir / "example-config.yaml"
        if not src_tpl.is_file():
            print(f"Bundled example not found: {src_tpl}", file=sys.stderr)
            return 1
        shutil.copy2(src_tpl, project_dir / _TEMPLATE_NAME)
        shutil.copy2(src_cfg, project_dir / _CONFIG_NAME)

    print(f"Initialised {project_dir}/")
    print(f"  {_TEMPLATE_NAME}  — slide template")
    print(f"  {_CONFIG_NAME}    — colours, fonts, layout config")
    print()
    print("Generate slides:  pptx generate spec.yaml -o output.pptx")
    print(f"Edit config:      $EDITOR {project_dir / _CONFIG_NAME}")
    return 0


def cmd_init_config(args: _InitConfigArgs) -> int:
    """Generate a starter template-config.yaml by introspecting a PPTX template.

    Extracts theme colors, fonts, placeholder indices, layout names, and slide
    dimensions.  The output is a complete (but approximate) config that an agent
    or user can review and refine — especially bullet levels and spacing, which
    require manual inspection of the slide master lstStyle.
    """
    from .inspect_pptx import resolve_theme_colors

    prs = _open(args.file)
    theme_colors = resolve_theme_colors(prs)

    # --- Slide dimensions ---
    slide_w = int(prs.slide_width or 12192000)
    slide_h = int(prs.slide_height or 6858000)

    # --- Theme fonts ---
    headline_font = "Calibri"
    body_font = "Calibri"
    try:
        from lxml import etree as _etree

        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        master_part = prs.slide_masters[0].part
        for rel in master_part.rels.values():
            if "theme" in str(rel.reltype):
                theme_el: _etree._Element = rel.target_part.element  # type: ignore[union-attr]
                major_el: _etree._Element | None = theme_el.find(  # type: ignore[assignment]
                    f".//{{{ns}}}majorFont/{{{ns}}}latin"
                )
                minor_el: _etree._Element | None = theme_el.find(  # type: ignore[assignment]
                    f".//{{{ns}}}minorFont/{{{ns}}}latin"
                )
                if major_el is not None:
                    tf = str(major_el.get("typeface") or "")  # type: ignore[arg-type]
                    if tf:
                        headline_font = tf
                if minor_el is not None:
                    tf = str(minor_el.get("typeface") or "")  # type: ignore[arg-type]
                    if tf:
                        body_font = tf
                break
    except Exception:
        pass

    # --- Placeholders from first content layout ---
    placeholder_map: dict[str, int] = {"title": 0, "subtitle": 1}
    layout_names: list[str] = []
    for layout in prs.slide_layouts:
        if layout.name:
            layout_names.append(layout.name)

    # Scan all layouts for placeholder indices
    for layout in prs.slide_layouts:
        for ph in layout.placeholders:
            name_lower = ph.name.lower()
            idx = ph.placeholder_format.idx
            if "tracker" in name_lower or "on-page" in name_lower or "breadcrumb" in name_lower:
                placeholder_map["tracker"] = idx
            elif "source" in name_lower or "footnote" in name_lower:
                placeholder_map.setdefault("source", idx)

    # --- Estimate layout geometry from first content layout ---
    # Find a "Default" or first non-title layout
    content_layout = None
    for layout in prs.slide_layouts:
        if layout.name and "default" in layout.name.lower():
            content_layout = layout
            break
    if content_layout is None and len(prs.slide_layouts) > 1:
        content_layout = prs.slide_layouts[1]

    # Estimate margins and key Y positions from placeholder positions
    left_margin = 554736
    right_margin = 554736
    title_y = 182372
    content_start_y = 1710000
    footer_y = int(slide_h * 0.94)
    header_line_y = 1181907

    if content_layout is not None:
        for ph in content_layout.placeholders:
            idx = ph.placeholder_format.idx
            if idx == 0:  # title
                title_y = int(ph.top)
                left_margin = int(ph.left)
                right_margin = slide_w - int(ph.left) - int(ph.width)
            elif int(ph.top) > slide_h * 0.5:
                # Below midpoint — likely footer area
                footer_y = int(ph.top)
            elif int(ph.top) > content_start_y * 0.8:
                content_start_y = int(ph.top)

        # Header line is typically just below subtitle
        for ph in content_layout.placeholders:
            idx = ph.placeholder_format.idx
            if idx == 1:  # subtitle
                header_line_y = int(ph.top) + int(ph.height) + 10000
                break

    # --- Map theme colors to semantic names ---
    def _theme_hex(name: str, fallback: str) -> str:
        val = theme_colors.get(name, fallback)
        return f"#{val}" if not val.startswith("#") else val

    dk1 = _theme_hex("dk1", "000000")
    lt1 = _theme_hex("lt1", "FFFFFF")
    dk2 = _theme_hex("dk2", "444444")
    lt2 = _theme_hex("lt2", "F0F0F0")
    accent1 = _theme_hex("accent1", "4472C4")
    accent3 = _theme_hex("accent3", "A5A5A5")

    # --- Build config ---
    config_lines = [
        f"# Template config generated from: {Path(args.file).name}",
        "# Generated by: pptx init-config",
        "# Review and adjust values — especially bullets, spacing, and colors.",
        "",
        "colors:",
        f'  midnight: "{dk1}"          # dk1 / primary text',
        f'  light_1: "{lt1}"           # lt1 / primary background',
        f'  light_2: "{lt2}"           # lt2 / secondary background',
        f'  slate: "{dk2}"             # dk2 / secondary text',
        f'  electric_blue: "{accent1}" # accent1 / primary accent',
        f'  cyan: "{accent3}"          # accent3 / secondary accent',
        '  red: "#E5546C"             # RAG: critical (adjust to your palette)',
        '  orange: "#FAA082"          # RAG: high',
        '  beige: "#E8BDAD"           # RAG: neutral',
        '  light_green: "#92D050"     # RAG: medium',
        '  green: "#00B050"           # RAG: low',
        "",
        "fonts:",
        f'  headline: "{headline_font}"',
        f'  body: "{body_font}"',
        "",
        "font_sizes:",
        "  title: 24",
        "  subtitle: 16",
        "  table_header: 14",
        "  table_body: 12",
        "  footnote: 8",
        "  tracker: 8",
        "  default: 12",
        "",
        "# Soft limits used for agent-friendly warnings",
        "text_limits:",
        "  title_max_lines: 2",
        "  subtitle_max_lines: 1",
        "  title_slide_subtitle_max_lines: 1",
        "",
        "layout:",
        f"  slide_width_emu: {slide_w}",
        f"  slide_height_emu: {slide_h}",
        f"  left_margin_emu: {left_margin}",
        f"  right_margin_emu: {right_margin}",
        f"  top_margin_emu: {title_y}",
        f"  tracker_y_emu: {max(title_y - 107000, 75000)}",
        f"  title_y_emu: {title_y}",
        f"  subtitle_y_emu: {title_y + 720000}",
        f"  header_line_y_emu: {header_line_y}",
        f"  content_start_y_emu: {content_start_y}",
        f"  footer_line_y_emu: {footer_y}",
        f"  footer_y_emu: {footer_y + 45000}",
        "",
        "placeholders:",
        f"  title: {placeholder_map.get('title', 0)}",
        f"  subtitle: {placeholder_map.get('subtitle', 1)}",
    ]
    if "tracker" in placeholder_map:
        config_lines.append(f"  tracker: {placeholder_map['tracker']}")
    if "source" in placeholder_map:
        config_lines.append(f"  source: {placeholder_map['source']}")

    config_lines.extend(
        [
            "",
            "default_colors:",
            '  body_text: "tx1"',
            '  col_header: "tx1"',
            '  col_superheader: "tx1"',
            '  row_header: "accent1"     # adjust: which theme slot is your accent color?',
            '  row_superheader: "tx1"',
            '  divider: "tx1"',
            '  link: "accent1"',
            "",
            "dividers:",
            "  header_pt: 1.5",
            "  row_pt: 0.5",
            "  footer_pt: 0.5",
            "",
            "# Bullet hierarchy — extracted from the slide master lstStyle.",
            "# Run: pptx xml <template> 1 <placeholder> to inspect lstStyle levels.",
            "# Adjust bullet chars, margins, and spacing to match your template.",
            "bullets:",
            "  def_tab_sz_emu: 914354",
            "  bu_font:",
            f'    typeface: "{body_font}"',
            '    pitch_family: "34"',
            '    charset: "0"',
            "  def_rpr:",
            "    size_pt: 12",
            "    kern_pt: 12",
            '    scheme: "tx1"',
            '    latin: "+mn-lt"',
            f'    cs: "{body_font}"',
            '    cs_pitch_family: "34"',
            '    cs_charset: "0"',
            "  levels:",
            "    - level: 1",
            '      bullet: ""',
            "      mar_l_emu: 0",
            "      indent_emu: 0",
            "      spc_bef_pt: 5.0",
            "      spc_aft_pt: 0.0",
            "      ln_spc_pct: 114000",
            "    - level: 2",
            '      bullet: "•"',
            "      mar_l_emu: 228600",
            "      indent_emu: -228600",
            "      spc_bef_pt: 2.5",
            "      spc_aft_pt: 0.0",
            "      ln_spc_pct: 114000",
            "    - level: 3",
            '      bullet: "–"',
            "      mar_l_emu: 457200",
            "      indent_emu: -228600",
            "      spc_bef_pt: 1.25",
            "      spc_aft_pt: 0.0",
            "      ln_spc_pct: 114000",
            "    - level: 4",
            '      bullet: "»"',
            "      mar_l_emu: 685800",
            "      indent_emu: -228600",
            "      spc_bef_pt: 1.25",
            "      spc_aft_pt: 0.0",
            "      ln_spc_pct: 114000",
            "    - level: 5",
            '      bullet: "›"',
            "      mar_l_emu: 914400",
            "      indent_emu: -228600",
            "      spc_bef_pt: 1.25",
            "      spc_aft_pt: 0.0",
            "      ln_spc_pct: 114000",
            "    - level: 6",
            '      bullet: "◦"',
            "      mar_l_emu: 1143000",
            "      indent_emu: -228600",
            "      spc_bef_pt: 0.0",
            "      spc_aft_pt: 0.0",
            "      ln_spc_pct: 114000",
            "    - level: 7",
            '      bullet: ""',
            "      mar_l_emu: 914400",
            "      indent_emu: 0",
            "      spc_bef_pt: 0.0",
            "      spc_aft_pt: 0.0",
            "      ln_spc_pct: 100000",
            "    - level: 8",
            '      bullet: "▫"',
            "      mar_l_emu: 1143000",
            "      indent_emu: -228600",
            "      spc_bef_pt: 0.0",
            "      spc_aft_pt: 0.0",
            "      ln_spc_pct: 100000",
            "    - level: 9",
            '      bullet: "▫"',
            "      mar_l_emu: 1143000",
            "      indent_emu: -228600",
            "      spc_bef_pt: 0.0",
            "      spc_aft_pt: 0.0",
            "      ln_spc_pct: 100000",
            "",
            "table_defaults:",
            "  min_row_height_in: 0.4",
            "  cell_padding_emu: 45720",
            "  header_row_height_in: 0.5",
            "  width_step_emu: 45720",
            "  default_width_in: 1.5",
            "  row_header_width_in: 1.0",
            "  moon_width_in: 0.7",
            "  bullets_width_in: 4.0",
            "  superheader_pt_boost: 2",
            "  line_spacing: 1.14",
            "  max_header_lines: 4",
            "  row_header_target_lines: 2",
            "  shape_id_start: 1000",
            "  legend:",
            "    label_pt: 8",
            "    y_offset_in: 0.30",
            "    row_height_ratio: 1.6",
            "    gap_ratio: 0.4",
            "    char_width_ratio: 0.55",
            "",
            "icons:",
            "  default_size_emu: 228600",
            "",
            "moon:",
            "  size_emu: 228600",
            "  arc_adjustments:",
            '    "100": null',
            '    "75": ["16200000", "0"]',
            '    "50": ["16200000", "10800000"]',
            '    "25": ["16200000", "5400000"]',
            '    "0": ["16200000", "16200000"]',
            "  fills:",
            "    critical: 100",
            "    full: 100",
            "    high: 75",
            '    "75": 75',
            "    medium: 50",
            '    "50": 50',
            "    low: 25",
            '    "25": 25',
            "    none: 0",
            "    empty: 0",
            "    na: 0",
            "    disabled: 0",
            "  colors:",
            '    critical: "#E5546C"',
            '    high: "#FAA082"',
            '    medium: "#E8BDAD"',
            '    low: "#CBD5E1"',
            "    none: null",
            '    na: "#CBD5E1"',
            '    disabled: "#CBD5E1"',
            "  group:",
            "    child_offset_x: 762000",
            "    child_offset_y: 1270000",
            "    child_size: 254000",
            "    line_width_emu: 9525",
            '    bg_fill: "#F8FAFC"',
            '    outline_scheme: "tx1"',
            "",
            "# Available layouts in this template:",
        ]
    )
    for name in layout_names:
        config_lines.append(f"#   - {name}")

    output = "\n".join(config_lines) + "\n"

    if args.output:
        Path(args.output).write_text(output)
        print(f"Saved → {args.output}")
        print("Review the config, especially:")
        print("  - colors: map your template's theme colors to semantic names")
        print("  - bullets: inspect slide master lstStyle for accurate margins/chars")
        print("  - placeholders: verify indices match your template")
        print("  - font_sizes: adjust to match your template's type scale")
    else:
        print(output)

    return 0


def cmd_validate(args: _ValidateArgs) -> int:
    """Validate schema for YAML files."""
    _apply_config(args.config)

    input_files = _expand_inputs(args.input)
    if not input_files:
        print("No input files found", file=sys.stderr)
        return 1

    all_valid = True
    for path in input_files:
        data = load_yaml(str(path))
        errors, warnings = validate_spec(data)
        if errors:
            all_valid = False
            print(f"{path}:\n  - " + "\n  - ".join(errors))
        elif warnings:
            print(f"{path}:\n  - " + "\n  - ".join(warnings))
        else:
            print(f"{path}: OK")
    return 0 if all_valid else 1


def cmd_preview(args: _InputArgs) -> int:
    """Preview table structure without generating PPTX."""
    input_files = _expand_inputs(args.input)
    if not input_files:
        print("No input files found", file=sys.stderr)
        return 1

    for path in input_files:
        data = load_yaml(str(path))
        print(preview_spec(data))
    return 0


def cmd_verify(args: _VerifyArgs) -> int:
    """Run layout solver + report without generating PPTX."""
    _apply_config(args.config)

    input_files = _expand_inputs(args.input)
    if not input_files:
        print("No input files found", file=sys.stderr)
        return 1

    metrics = TextMetrics()
    solver = ConstraintSolver(metrics)

    for path in input_files:
        data = load_yaml(str(path))
        errors, warnings = validate_spec(data)
        if errors:
            print(f"{path}:\n  - " + "\n  - ".join(errors), file=sys.stderr)
            return 1
        if warnings:
            print(f"{path}:\n  - " + "\n  - ".join(warnings))

        spec, area, options, placeholders = parse_spec(data)
        if placeholders:
            spec = fill_placeholders(spec)

        _, report = solver.solve(spec, area, options)
        print(f"{path}: {report.to_text(detail=args.detail)}")

        if args.json:
            output_path = Path(args.json)
            output_path.write_text(json.dumps(report.to_json(), indent=2))

    return 0


def cmd_screenshot(args: _ScreenshotArgs) -> int:
    """Generate PNG screenshots from PPTX files."""
    input_files = _expand_inputs(args.input)
    if not input_files:
        print("No input files found", file=sys.stderr)
        return 1

    output_dir = Path(args.output_dir or "outputs/images")
    generator = ScreenshotGenerator(soffice_path=args.soffice)

    for path in input_files:
        png_path = generator.capture(Path(path), output_dir, slide_index=args.slide)
        print(f"{path}: {png_path}")

    return 0


# ============================================================================
# ARGPARSE SETUP
# ============================================================================


def _build_parser() -> argparse.ArgumentParser:
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
    p.add_argument("--module-path", help="Path to chart generator module")
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


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
