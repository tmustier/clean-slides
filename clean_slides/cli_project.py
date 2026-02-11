"""Project discovery and input expansion helpers for CLI commands."""

from __future__ import annotations

from pathlib import Path

from .template_config import set_template_config

PROJECT_DIR_NAME = ".clean-slides"
CONFIG_NAME = "config.yaml"
TEMPLATE_NAME = "template.pptx"


def discover_project_dir() -> Path | None:
    """Walk from CWD to filesystem root looking for a `.clean-slides/` directory."""
    cur = Path.cwd().resolve()
    for parent in [cur, *cur.parents]:
        candidate = parent / PROJECT_DIR_NAME
        if candidate.is_dir():
            return candidate
    return None


def discover_config() -> Path | None:
    """Return the project config path if auto-discovered."""
    proj = discover_project_dir()
    if proj is not None:
        cfg = proj / CONFIG_NAME
        if cfg.is_file():
            return cfg
    return None


def discover_template() -> Path | None:
    """Return the project template path if auto-discovered."""
    proj = discover_project_dir()
    if proj is not None:
        tpl = proj / TEMPLATE_NAME
        if tpl.is_file():
            return tpl
    return None


def apply_config(config_path: str | None) -> None:
    """Load template config — explicit path, auto-discovered, or built-in defaults."""
    if config_path is not None:
        set_template_config(Path(config_path))
        return

    discovered = discover_config()
    if discovered is not None:
        set_template_config(discovered)


def expand_inputs(inputs: list[str]) -> list[Path]:
    files: list[Path] = []
    for pattern in inputs:
        path = Path(pattern)
        if path.is_file():
            files.append(path)
        else:
            files.extend(Path(".").glob(pattern))
    return files
