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
from pathlib import Path
from typing import Any, Protocol

from pptx import Presentation
from pptx.presentation import Presentation as PresentationObj
from pptx.shapes.autoshape import Shape
from pptx.shapes.base import BaseShape
from pptx.slide import Slide, SlideLayout
from pptx.util import Emu

from .cli_common import find_layout as _find_layout
from .cli_common import find_shape as _find_shape
from .cli_common import get_slide as _get_slide
from .cli_common import is_str_object_dict as _is_str_object_dict
from .cli_common import is_text_shape as _is_text_shape
from .cli_common import open_presentation as _open
from .cli_common import shape_text as _shape_text
from .cli_deck import cmd_add_slide, cmd_delete_shape, cmd_delete_slide, cmd_insert
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
from .cli_render import cmd_charts, cmd_crop, cmd_render
from .cli_text import text_preview as _text_preview
from .cli_text import write_to_shape as _write_to_shape
from .constants import EMU_PER_INCH, Fonts, FontSizes, Layout, TableDefaults
from .metadata import fill_slide_metadata
from .placeholder import fill_placeholders
from .renderer import TableRenderer
from .screenshot import ScreenshotGenerator
from .solver import ConstraintSolver
from .spec import ContentArea, TableLayout, TableSpec
from .spec_pipeline import YamlDict, load_yaml, parse_spec, preview_spec, validate_spec
from .template_config import TEMPLATE_CONFIG, set_template_config
from .text_metrics import EMU_PER_PT, TextMetrics

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


# ============================================================================
# TEXT PARSING (for edit/batch commands)
# ============================================================================


# Text parsing and writing helpers moved to clean_slides.cli_text

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


# ============================================================================
# GENERATE COMMANDS (YAML → PPTX table pipeline)
# ============================================================================



# YAML parsing/validation helpers now live in clean_slides.spec_pipeline

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


def _resolve_generate_template_path(template_path: str | None) -> str | None:
    if template_path is not None:
        return template_path

    discovered_tpl = _discover_template()
    if discovered_tpl is not None:
        return str(discovered_tpl)

    return None


def _init_generate_presentation(template_path: str | None) -> PresentationObj:
    if template_path:
        prs = Presentation(template_path)
        # Remove template's content slides - keep only layouts/masters
        _delete_all_slides(prs)
        return prs

    _hint_init()
    prs = Presentation()
    prs.slide_width = Emu(int(Layout.SLIDE_WIDTH))
    prs.slide_height = Emu(int(Layout.SLIDE_HEIGHT))
    return prs


def _load_validated_generate_data(path: Path) -> tuple[YamlDict, bool] | None:
    data = load_yaml(str(path))
    errors, warnings = validate_spec(data)
    if errors:
        print(f"{path}:\n  - " + "\n  - ".join(errors), file=sys.stderr)
        return None
    if warnings:
        print(f"{path}:\n  - " + "\n  - ".join(warnings))

    has_table = _is_str_object_dict(data.get("table")) and bool(data.get("table"))
    return data, has_table


def _resolve_generate_slide(
    prs: PresentationObj,
    data: YamlDict,
    slide_index: int | None,
) -> tuple[Slide, SlideLayout | None, str | None] | None:
    layout_override: str | None = None

    if slide_index is not None:
        if slide_index < 0 or slide_index >= len(prs.slides):
            print("slide_index out of range", file=sys.stderr)
            return None

        slide = prs.slides[slide_index]
        if "content_layout" not in data and "layout" not in data:
            layout_override = _infer_layout_from_slide(slide)

        return slide, None, layout_override

    slide_layout_name = str(data.get("slide_layout") or "Default")
    slide_layout_obj = _find_layout(prs, slide_layout_name, fallback=True)
    slide = prs.slides.add_slide(slide_layout_obj)
    return slide, slide_layout_obj, None


def _render_chart_cells_for_spec(
    slide: Slide,
    spec: TableSpec,
    layout: TableLayout,
    area: ContentArea,
) -> None:
    if not spec.chart_defs:
        return

    from .chart_render import render_chart_cells
    from .charts import load_charts_module

    charts_mod = load_charts_module()
    body_pt = layout.body_font_size // 100
    render_chart_cells(
        slide,
        spec,
        layout,
        area,
        charts_mod,
        label_font_size_pt=body_pt,
    )


def _render_sidebar_content(
    data: YamlDict,
    slide_layout_obj: SlideLayout | None,
    renderer: TableRenderer,
    metrics: TextMetrics,
) -> None:
    sidebar_raw = data.get("sidebar")
    if sidebar_raw is None or slide_layout_obj is None:
        return

    sidebar_area = _sidebar_content_area(slide_layout_obj)
    if sidebar_area is None:
        print("  WARNING: sidebar content specified but layout has no secondary content area")
        return

    from .content import Paragraph, normalize_cell

    default_para = Paragraph(text="", lvl=0)
    sidebar_paras = normalize_cell(sidebar_raw, default_para, parse_bullets=True)
    if not sidebar_paras:
        return

    if data.get("sidebar_shrink"):
        _shrink_sidebar_to_fit(sidebar_paras, sidebar_area, metrics)
    else:
        _warn_sidebar_overflow(sidebar_paras, sidebar_area, metrics)

    renderer.render_sidebar(sidebar_paras, sidebar_area)


def _render_table_for_input(
    path: Path,
    slide: Slide,
    slide_layout_obj: SlideLayout | None,
    data: YamlDict,
    layout_override: str | None,
    solver: ConstraintSolver,
    metrics: TextMetrics,
    *,
    detail: bool,
    target_slide_index: int | None,
    keep_existing: bool,
) -> None:
    spec, area, options, placeholders = parse_spec(data, layout_override=layout_override)
    if placeholders:
        spec = fill_placeholders(spec)

    # Derive content area from the layout's primary content placeholder
    # unless the YAML explicitly overrides via content_area / content_layout.
    if slide_layout_obj is not None and "content_area" not in data:
        layout_area = _content_area_from_layout(slide_layout_obj)
        if layout_area is not None:
            area = layout_area

    if target_slide_index is not None and not keep_existing:
        _clear_content_area(slide, area)

    layout, report = solver.solve(spec, area, options)
    print(f"{path}: {report.to_text(detail=detail)}")

    sp_tree = slide.shapes.element

    shape_id: int = TableDefaults.SHAPE_ID_START

    def next_shape_id() -> int:
        nonlocal shape_id
        shape_id += 1
        return shape_id

    renderer = TableRenderer(sp_tree, next_shape_id, slide_part=slide.part)
    renderer.render(spec, layout, area)

    _render_chart_cells_for_spec(slide, spec, layout, area)
    _render_sidebar_content(data, slide_layout_obj, renderer, metrics)


def _finalize_generated_slide(slide: Slide, data: YamlDict) -> None:
    fill_slide_metadata(slide, data)

    # Agent-friendly: warn when title/subtitle likely wrap beyond configured limits.
    for shape in slide.shapes:
        if isinstance(shape, Shape) and shape.has_text_frame:
            _warn_placeholder_text_limits(slide, shape)

    _clear_body_placeholders(slide)


def _process_generate_input(
    path: Path,
    prs: PresentationObj,
    args: _GenerateArgs,
    solver: ConstraintSolver,
    metrics: TextMetrics,
) -> bool:
    loaded = _load_validated_generate_data(path)
    if loaded is None:
        return False

    data, has_table = loaded

    resolved = _resolve_generate_slide(prs, data, args.slide_index)
    if resolved is None:
        return False

    slide, slide_layout_obj, layout_override = resolved

    if has_table:
        _render_table_for_input(
            path,
            slide,
            slide_layout_obj,
            data,
            layout_override,
            solver,
            metrics,
            detail=args.detail,
            target_slide_index=args.slide_index,
            keep_existing=args.keep_existing,
        )
    else:
        print(f"{path}: metadata-only slide (no table)")

    _finalize_generated_slide(slide, data)
    return True


def cmd_generate(args: _GenerateArgs) -> int:
    """Generate PPTX for text-only tables."""
    _apply_config(args.config)

    input_files = _expand_inputs(args.input)
    if not input_files:
        print("No input files found", file=sys.stderr)
        return 1

    template_path = _resolve_generate_template_path(args.template)
    prs = _init_generate_presentation(template_path)

    metrics = TextMetrics()
    solver = ConstraintSolver(metrics)

    for path in input_files:
        ok = _process_generate_input(path, prs, args, solver, metrics)
        if not ok:
            return 1

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
    """Generate a starter template-config.yaml by introspecting a PPTX template."""
    from .template_init_config import build_init_config_output

    output = build_init_config_output(args.file)

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
