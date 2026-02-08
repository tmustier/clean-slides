"""
Render slides to PNG and crop regions for inspection.

Rendering engines (in priority order):
  1. PowerPoint via AppleScript (macOS, faithful to on-screen rendering)
  2. LibreOffice (cross-platform fallback, ~95% accurate)

Usage:
    render_slide(pptx_path, slide_num) → PNG path
    crop_region(png_path, left, top, right, bottom) → cropped PNG path

Coordinates for crop_region are in inches (matching inspect_slide output).
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional, Union

from PIL import Image
from PIL.Image import Image as PILImage

# Slide dimensions in inches (widescreen 13.333 x 7.5)
SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5
RENDER_DPI = 150

Pathish = Union[str, Path]


# ── High-level API ─────────────────────────────────────────────────────


def render_slide(
    pptx_path: Pathish,
    slide_num: int = 1,
    dpi: int = RENDER_DPI,
    output_dir: Optional[Pathish] = None,
    engine: Optional[str] = None,
) -> str:
    """Render a single slide to PNG.  Convenience wrapper around render_slides."""
    paths = render_slides(pptx_path, [slide_num], dpi=dpi, output_dir=output_dir, engine=engine)
    return paths[0]


def render_slides(
    pptx_path: Pathish,
    slide_nums: list[int],
    dpi: int = RENDER_DPI,
    output_dir: Optional[Pathish] = None,
    engine: Optional[str] = None,
) -> list[str]:
    """Render multiple slides to PNG (single PDF conversion).

    Args:
        pptx_path: path to .pptx file
        slide_nums: 1-indexed slide numbers
        dpi: render resolution (default 150)
        output_dir: where to save (default: temp dir)
        engine: "powerpoint", "libreoffice", or None (auto-detect)

    Returns:
        List of PNG paths, one per requested slide.
    """

    pptx_path_str = os.path.abspath(str(pptx_path))

    if output_dir is None:
        output_dir_str = tempfile.mkdtemp(prefix="slide_render_")
    else:
        output_dir_str = str(output_dir)

    os.makedirs(output_dir_str, exist_ok=True)

    selected_engine = engine or _detect_engine()

    if selected_engine == "powerpoint":
        return _render_via_powerpoint(pptx_path_str, slide_nums, dpi, output_dir_str)
    return _render_via_libreoffice(pptx_path_str, slide_nums, dpi, output_dir_str)


def crop_region(
    png_path: Pathish,
    left: float,
    top: float,
    right: float,
    bottom: float,
    output_path: Optional[Pathish] = None,
) -> str:
    """Crop a region from a rendered slide PNG.

    Coordinates are in inches (matching inspect_slide positions).

    Args:
        png_path: path to full slide PNG
        left, top, right, bottom: region in inches
        output_path: where to save (default: auto-named in same dir)

    Returns:
        Path to cropped PNG.
    """

    png_path_str = os.fspath(png_path)

    img: PILImage = Image.open(png_path_str)
    px_w, px_h = img.size

    scale_x = px_w / SLIDE_WIDTH
    scale_y = px_h / SLIDE_HEIGHT

    box = (
        int(left * scale_x),
        int(top * scale_y),
        int(right * scale_x),
        int(bottom * scale_y),
    )

    cropped: PILImage = img.crop(box)

    if output_path is None:
        base = os.path.splitext(png_path_str)[0]
        output_path_str = f"{base}_crop_{left:.1f}_{top:.1f}_{right:.1f}_{bottom:.1f}.png"
    else:
        output_path_str = os.fspath(output_path)

    cropped.save(output_path_str, "PNG")
    return output_path_str


# ── Convenience class (used by generate pipeline) ─────────────────────


class ScreenshotGenerator:
    """Generate PNG screenshots from PPTX files."""

    def __init__(self, soffice_path: Optional[str] = None):
        self.soffice_path = soffice_path or "soffice"

    def capture(self, pptx_path: Path, out_dir: Path, slide_index: int = 0) -> Path:
        """Return PNG path for a slide (0-indexed)."""
        out_dir.mkdir(parents=True, exist_ok=True)

        # Use the unified render function (1-indexed)
        png_path = render_slide(
            pptx_path,
            slide_num=slide_index + 1,
            output_dir=out_dir,
        )
        return Path(png_path)


# ── Engine detection ───────────────────────────────────────────────────


def _detect_engine() -> str:
    """Auto-detect best available rendering engine."""
    if sys.platform == "darwin" and _powerpoint_available():
        return "powerpoint"
    return "libreoffice"


def _powerpoint_available() -> bool:
    """Check if PowerPoint is installed on macOS."""
    return os.path.exists("/Applications/Microsoft PowerPoint.app")


# ── PDF conversion helper (shared) ─────────────────────────────────────


def _pdf_to_images(pdf_path: str, dpi: int) -> list[PILImage]:
    """Convert a PDF into a list of PIL Images."""

    try:
        from pdf2image import convert_from_path
    except ImportError as exc:  # pragma: no cover
        raise ImportError("pdf2image required: pip install pdf2image") from exc

    # The `pdf2image` project doesn't ship complete typing information.
    # We provide a local stub (typings/pdf2image) so this call is typed.
    return convert_from_path(pdf_path, dpi=dpi)


# ── PowerPoint via AppleScript (macOS) ─────────────────────────────────


def _render_via_powerpoint(
    pptx_path: str, slide_nums: list[int], dpi: int, output_dir: str
) -> list[str]:
    """Render via PowerPoint's native PDF export + pdf2image."""

    pptx_dir = os.path.dirname(pptx_path)
    pdf_name = Path(pptx_path).stem + "_export.pdf"
    pdf_path = os.path.join(pptx_dir, pdf_name)

    script = f"""
    tell application "Microsoft PowerPoint"
        open POSIX file "{pptx_path}"
        set theDoc to active presentation
        save theDoc in POSIX file "{pdf_path}" as save as PDF
        close theDoc saving no
    end tell
    """

    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
        timeout=60,
    )
    if result.returncode != 0:
        raise RuntimeError(f"PowerPoint export failed: {result.stderr}")

    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PowerPoint did not create PDF at {pdf_path}")

    images = _pdf_to_images(pdf_path, dpi=dpi)
    paths: list[str] = []
    for slide_num in slide_nums:
        if slide_num < 1 or slide_num > len(images):
            raise ValueError(f"Slide {slide_num} out of range (1-{len(images)})")
        img = images[slide_num - 1]
        png_path = os.path.join(output_dir, f"slide_{slide_num}.png")
        img.save(png_path, "PNG")
        paths.append(png_path)

    try:
        os.remove(pdf_path)
    except OSError:
        pass

    return paths


# ── LibreOffice fallback ───────────────────────────────────────────────


def _render_via_libreoffice(
    pptx_path: str, slide_nums: list[int], dpi: int, output_dir: str
) -> list[str]:
    """Render via LibreOffice's PDF export + pdf2image."""

    pdf_path = os.path.join(output_dir, Path(pptx_path).stem + ".pdf")

    cmd: list[str] = [
        "soffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        output_dir,
        pptx_path,
    ]
    subprocess.run(cmd, capture_output=True, check=True)

    images = _pdf_to_images(pdf_path, dpi=dpi)
    paths: list[str] = []
    for slide_num in slide_nums:
        if slide_num < 1 or slide_num > len(images):
            raise ValueError(f"Slide {slide_num} out of range (1-{len(images)})")
        img = images[slide_num - 1]
        png_path = os.path.join(output_dir, f"slide_{slide_num}.png")
        img.save(png_path, "PNG")
        paths.append(png_path)

    return paths
