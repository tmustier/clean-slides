from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from pptx import Presentation

from clean_slides.cli import cmd_generate, cmd_verify


@dataclass
class GenerateArgs:
    input: List[str]
    template: Optional[str]
    slide_index: Optional[int]
    keep_existing: bool
    detail: bool
    output: Optional[str]
    config: Optional[str] = None


@dataclass
class VerifyArgs:
    input: List[str]
    detail: bool
    json: Optional[str]
    config: Optional[str] = None


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


def test_generate_smoke(tmp_path: Path) -> None:
    yaml_path = tmp_path / "table.yaml"
    _write_yaml(yaml_path)

    output_path = tmp_path / "output.pptx"
    args = GenerateArgs(
        input=[str(yaml_path)],
        template=None,
        slide_index=None,
        keep_existing=False,
        detail=False,
        output=str(output_path),
    )

    result = cmd_generate(args)

    assert result == 0
    assert output_path.exists()
    assert output_path.stat().st_size > 0

    prs = Presentation(str(output_path))
    assert len(prs.slides) == 1
    assert len(prs.slides[0].shapes) > 0


def test_verify_smoke(tmp_path: Path) -> None:
    yaml_path = tmp_path / "table.yaml"
    _write_yaml(yaml_path)

    args = VerifyArgs(
        input=[str(yaml_path)],
        detail=False,
        json=None,
    )

    result = cmd_verify(args)

    assert result == 0
