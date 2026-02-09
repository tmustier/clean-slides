"""
Load template-specific configuration values.

The config lives in ``template-config.yaml`` (packaged with the library)
so template-derived values are not hardcoded in Python.

Custom templates
~~~~~~~~~~~~~~~~
Use :func:`set_template_config` to replace the active config at runtime
(before any generation work).  The CLI ``--config`` flag calls this.
"""

from __future__ import annotations

from dataclasses import dataclass
from importlib import resources
from pathlib import Path
from typing import Any

import yaml
from typing_extensions import TypeGuard

CONFIG_FILENAME = "template-config.yaml"

_REQUIRED_SECTIONS = {
    "colors": [
        "midnight",
        "light_1",
        "light_2",
        "slate",
        "electric_blue",
        "cyan",
        "red",
        "orange",
        "beige",
        "light_green",
        "green",
    ],
    "fonts": ["headline", "body"],
    "font_sizes": [
        "title",
        "subtitle",
        "table_header",
        "table_body",
        "footnote",
        "tracker",
        "default",
    ],
    "layout": [
        "slide_width_emu",
        "slide_height_emu",
        "left_margin_emu",
        "right_margin_emu",
        "top_margin_emu",
        "tracker_y_emu",
        "title_y_emu",
        "subtitle_y_emu",
        "header_line_y_emu",
        "content_start_y_emu",
        "footer_line_y_emu",
        "footer_y_emu",
    ],
    "dividers": ["header_pt", "row_pt", "footer_pt"],
    "bullets": ["def_tab_sz_emu", "bu_font", "def_rpr", "levels"],
    "table_defaults": [
        "min_row_height_in",
        "cell_padding_emu",
        "header_row_height_in",
        "width_step_emu",
        "default_width_in",
        "row_header_width_in",
        "moon_width_in",
        "bullets_width_in",
        "superheader_pt_boost",
        "line_spacing",
        "max_header_lines",
        "row_header_target_lines",
        "shape_id_start",
        "legend",
    ],
    "icons": ["default_size_emu"],
    "moon": ["size_emu", "arc_adjustments", "fills", "colors", "group"],
}

# Optional sections — validated when present but not required.
_OPTIONAL_SECTIONS = {
    "placeholders": ["title"],  # subtitle, tracker are optional within placeholders
    "default_colors": ["body_text", "divider", "link"],
}

_BULLET_LEVEL_KEYS = {
    "level",
    "bullet",
    "mar_l_emu",
    "indent_emu",
    "spc_bef_pt",
    "spc_aft_pt",
    "ln_spc_pct",
}

_BULLET_BU_FONT_KEYS = {"typeface", "pitch_family", "charset"}
_BULLET_DEF_RPR_KEYS = {
    "size_pt",
    "kern_pt",
    "scheme",
    "latin",
    "cs",
    "cs_pitch_family",
    "cs_charset",
}
_LEGEND_KEYS = {"label_pt", "y_offset_in", "row_height_ratio", "gap_ratio", "char_width_ratio"}
_MOON_GROUP_KEYS = {
    "child_offset_x",
    "child_offset_y",
    "child_size",
    "line_width_emu",
    "bg_fill",
    "outline_scheme",
}
_MOON_ARC_KEYS = {"0", "25", "50", "75", "100"}


def _is_dict(value: object) -> TypeGuard[dict[Any, Any]]:
    return isinstance(value, dict)


def _is_list(value: object) -> TypeGuard[list[Any]]:
    return isinstance(value, list)


@dataclass(frozen=True)
class TemplateConfig:
    """Thin wrapper around the YAML config dict."""

    raw: dict[str, Any]

    def section(self, key: str) -> dict[str, Any]:
        value = self.raw.get(key)
        if not _is_dict(value):
            raise KeyError(f"Missing or invalid template config section: {key}")
        return _stringify_keys(value)

    def has_section(self, key: str) -> bool:
        """Return ``True`` if *key* exists and is a dict."""
        return _is_dict(self.raw.get(key))

    def section_or_default(self, key: str, default: dict[str, Any]) -> dict[str, Any]:
        """Return section if present, otherwise *default*."""
        value = self.raw.get(key)
        if not _is_dict(value):
            return dict(default)
        return _stringify_keys(value)


def _stringify_keys(value: dict[Any, Any]) -> dict[str, Any]:
    return {str(k): v for k, v in value.items()}


def _require_dict(name: str, value: Any, errors: list[str]) -> dict[str, Any] | None:
    if not _is_dict(value):
        errors.append(f"{name} must be a mapping")
        return None
    return _stringify_keys(value)


def _require_keys(name: str, value: dict[str, Any], keys: set[str], errors: list[str]) -> None:
    missing = [key for key in keys if key not in value]
    for key in sorted(missing):
        errors.append(f"Missing {name}.{key}")


def validate_template_config(data: dict[str, Any]) -> list[str]:
    """Validate template config structure, returning a list of errors."""
    errors: list[str] = []

    for section, keys in _REQUIRED_SECTIONS.items():
        value = _require_dict(section, data.get(section), errors)
        if value is None:
            continue
        _require_keys(section, value, set(keys), errors)

    # Optional sections — only validate keys when the section exists.
    for section, keys in _OPTIONAL_SECTIONS.items():
        raw = data.get(section)
        if raw is None:
            continue
        value = _require_dict(section, raw, errors)
        if value is None:
            continue
        _require_keys(section, value, set(keys), errors)

    raw_bullets = data.get("bullets")
    bullets: dict[str, Any] | None = None
    if _is_dict(raw_bullets):
        bullets = _stringify_keys(raw_bullets)

    if bullets is not None:
        bu_font = _require_dict("bullets.bu_font", bullets.get("bu_font"), errors)
        if bu_font is not None:
            _require_keys("bullets.bu_font", bu_font, _BULLET_BU_FONT_KEYS, errors)

        def_rpr = _require_dict("bullets.def_rpr", bullets.get("def_rpr"), errors)
        if def_rpr is not None:
            _require_keys("bullets.def_rpr", def_rpr, _BULLET_DEF_RPR_KEYS, errors)

        levels = bullets.get("levels")
        if not _is_list(levels) or not levels:
            errors.append("bullets.levels must be a non-empty list")
        else:
            for idx, level in enumerate(levels):
                level_dict = _require_dict(f"bullets.levels[{idx}]", level, errors)
                if level_dict is None:
                    continue
                _require_keys(
                    f"bullets.levels[{idx}]",
                    level_dict,
                    _BULLET_LEVEL_KEYS,
                    errors,
                )

    raw_table_defaults = data.get("table_defaults")
    table_defaults: dict[str, Any] | None = None
    if _is_dict(raw_table_defaults):
        table_defaults = _stringify_keys(raw_table_defaults)

    if table_defaults is not None:
        legend = _require_dict("table_defaults.legend", table_defaults.get("legend"), errors)
        if legend is not None:
            _require_keys("table_defaults.legend", legend, _LEGEND_KEYS, errors)

    raw_moon = data.get("moon")
    moon: dict[str, Any] | None = None
    if _is_dict(raw_moon):
        moon = _stringify_keys(raw_moon)

    if moon is not None:
        group = _require_dict("moon.group", moon.get("group"), errors)
        if group is not None:
            _require_keys("moon.group", group, _MOON_GROUP_KEYS, errors)

        arc_adjustments = moon.get("arc_adjustments")
        if _is_dict(arc_adjustments):
            arc_adjustments_dict = _stringify_keys(arc_adjustments)
            missing = _MOON_ARC_KEYS.difference(set(arc_adjustments_dict.keys()))
            for key in sorted(missing):
                errors.append(f"Missing moon.arc_adjustments.{key}")
        else:
            errors.append("moon.arc_adjustments must be a mapping")

    return errors


def load_template_config(path: Path | None = None) -> TemplateConfig:
    """Load template config from *path* or package resource."""
    if path is not None:
        data = yaml.safe_load(path.read_text())
    else:
        config_path = resources.files("clean_slides").joinpath(CONFIG_FILENAME)
        data = yaml.safe_load(config_path.read_text())

    if not _is_dict(data):
        raise ValueError("Template config must be a mapping")

    normalized_data: dict[str, Any] = _stringify_keys(data)

    errors = validate_template_config(normalized_data)
    if errors:
        message = "Template config validation failed:\n" + "\n".join(f"- {e}" for e in errors)
        raise ValueError(message)

    return TemplateConfig(normalized_data)


TEMPLATE_CONFIG = load_template_config()


def set_template_config(path: Path) -> None:
    """Replace the active :data:`TEMPLATE_CONFIG` at runtime.

    This reloads the config from *path*, re-validates, and patches the
    module-level ``TEMPLATE_CONFIG`` so that all downstream imports
    (constants, renderer, etc.) pick up the new values.

    Must be called **before** any generation/rendering work — typically
    from the CLI when ``--config`` is supplied.
    """
    import clean_slides.constants as _constants_mod
    import clean_slides.metadata as _metadata_mod
    import clean_slides.template_config as _self

    new_config = load_template_config(path)

    # Patch the singleton in both the canonical module and this alias.
    _self.TEMPLATE_CONFIG = new_config
    globals()["TEMPLATE_CONFIG"] = new_config

    # Re-initialize derived constants.
    _constants_mod.reload_constants(new_config)

    # Re-initialize placeholder map.
    _metadata_mod.reload_placeholders(new_config)
