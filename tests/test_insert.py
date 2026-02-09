from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation

from clean_slides.cli import cmd_generate, cmd_insert


@dataclass
class GenerateArgs:
    input: list[str]
    template: str | None
    slide_index: int | None
    keep_existing: bool
    detail: bool
    output: str | None
    config: str | None = None


@dataclass
class InsertArgs:
    file: str
    source: str
    at: int | None
    slides: str | None
    out: str | None


def _write_yaml(path: Path) -> None:
    path.write_text(
        """
layout: default

table:
  rows: 2
  cols: 2
  has_col_header: true
  col_headers: ["Header A", "Header B"]
  cells:
    - ["Alpha", "Beta"]
    - ["Gamma", "Delta"]
""".strip()
    )


def test_insert_smoke(tmp_path: Path) -> None:
    # 1) Generate a source deck with one slide
    yaml_path = tmp_path / "table.yaml"
    _write_yaml(yaml_path)

    src_path = tmp_path / "src.pptx"
    gen_args = GenerateArgs(
        input=[str(yaml_path)],
        template=None,
        slide_index=None,
        keep_existing=False,
        detail=False,
        output=str(src_path),
    )
    assert cmd_generate(gen_args) == 0

    src_prs = Presentation(str(src_path))
    assert len(src_prs.slides) == 1
    src_shapes = len(src_prs.slides[0].shapes)
    assert src_shapes > 0

    # 2) Create a target deck with two blank slides
    target_path = tmp_path / "target.pptx"
    target_prs = Presentation()
    blank_layout = target_prs.slide_layouts[len(target_prs.slide_layouts) - 1]
    target_prs.slides.add_slide(blank_layout)
    target_prs.slides.add_slide(blank_layout)
    target_prs.save(str(target_path))

    # 3) Insert source slide at position 2
    out_path = tmp_path / "out.pptx"
    ins_args = InsertArgs(
        file=str(target_path),
        source=str(src_path),
        at=2,
        slides="1",
        out=str(out_path),
    )
    assert cmd_insert(ins_args) == 0

    out_prs = Presentation(str(out_path))
    assert len(out_prs.slides) == 3

    # Inserted slide is position 2 (index 1)
    assert len(out_prs.slides[1].shapes) == src_shapes
