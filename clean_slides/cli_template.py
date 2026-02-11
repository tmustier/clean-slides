"""Template/project initialization CLI commands."""

from __future__ import annotations

import shutil
import sys
from pathlib import Path
from typing import Protocol

from .cli_project import CONFIG_NAME, PROJECT_DIR_NAME, TEMPLATE_NAME
from .template_init_config import build_init_config_output


class InitArgs(Protocol):
    template: str | None
    output: str | None


class InitConfigArgs(Protocol):
    file: str
    output: str | None


def _example_template_dir() -> Path:
    """Return path to the bundled example-template directory."""
    return Path(__file__).resolve().parent / "example-template"


def cmd_init_config(args: InitConfigArgs) -> int:
    """Generate a starter template-config.yaml by introspecting a PPTX template."""
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


def cmd_init(args: InitArgs) -> int:
    """Initialise a .clean-slides/ project directory."""
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
        dst_tpl = project_dir / TEMPLATE_NAME
        shutil.copy2(src_tpl, dst_tpl)

        # Generate config via init-config
        dst_cfg = project_dir / CONFIG_NAME

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
        shutil.copy2(src_tpl, project_dir / TEMPLATE_NAME)
        shutil.copy2(src_cfg, project_dir / CONFIG_NAME)

    print(f"Initialised {project_dir}/")
    print(f"  {TEMPLATE_NAME}  — slide template")
    print(f"  {CONFIG_NAME}    — colours, fonts, layout config")
    print()
    print("Generate slides:  pptx generate spec.yaml -o output.pptx")
    print(f"Edit config:      $EDITOR {project_dir / CONFIG_NAME}")
    return 0
