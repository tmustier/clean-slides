"""
Icon indicator definitions and presets.

Provides IconSet (config for traffic-light circles in table cells)
and standard preset palettes (severity, RAG, confidence, priority).
"""

from dataclasses import dataclass
from typing import Any, Dict, Optional

from typing_extensions import TypeGuard

from .constants import EMU_PER_INCH, ICON_DEFAULT_SIZE_EMU


def _is_dict(value: object) -> TypeGuard[Dict[Any, Any]]:
    return isinstance(value, dict)


# ---------------------------------------------------------------------------
# Standard presets — mapping of value name → color (hex or scheme)
# ---------------------------------------------------------------------------

ICON_PRESETS: Dict[str, Dict[str, str]] = {
    "severity": {
        "Critical": "#E5546C",
        "High": "#FAA082",
        "Medium": "#E8BDAD",
        "Low": "accent5",
    },
    "rag": {
        "Red": "#E5546C",
        "Amber": "#FAA082",
        "Green": "#00B050",
        "Grey": "accent5",
    },
    "confidence": {
        "Strong": "tx1",
        "Moderate": "#2251FF",
        "Weak": "#00A9F4",
        "None": "accent5",
    },
    "priority": {
        "P1": "#E5546C",
        "P2": "#FAA082",
        "P3": "#E8BDAD",
        "P4": "accent5",
    },
}


# ---------------------------------------------------------------------------
# IconSet dataclass
# ---------------------------------------------------------------------------


@dataclass
class IconSet:
    """Configuration for icon indicators in table cells.

    ``values`` maps human-readable names (e.g. "Critical") to colors.
    Colors can be hex RGB (``"#E5546C"``) or PowerPoint scheme names
    (``"accent5"``).
    """

    values: Dict[str, str]  # name → color
    size_emu: int = ICON_DEFAULT_SIZE_EMU
    show_legend: bool = True
    legend_x: Optional[int] = None  # explicit EMU; auto if None
    legend_y: Optional[int] = None

    @classmethod
    def from_dict(cls, raw: Dict[str, Any]) -> "IconSet":
        values: Dict[str, str] = {}

        preset = raw.get("preset")
        if preset:
            preset_name = str(preset)
            values = dict(ICON_PRESETS.get(preset_name, {}))
            if not values:
                raise ValueError(f"Unknown icon preset: {preset_name}")

        # Custom values override/extend preset
        custom = raw.get("values")
        if _is_dict(custom):
            values.update({str(k): str(v) for k, v in custom.items()})

        if not values:
            raise ValueError("icons must specify 'preset' or 'values'")

        default_size_in = ICON_DEFAULT_SIZE_EMU / EMU_PER_INCH
        size_inches = float(raw.get("size", default_size_in))
        size_emu = int(size_inches * EMU_PER_INCH)

        show_legend = bool(raw.get("legend", True))

        legend_pos_raw = raw.get("legend_position")
        legend_pos: Dict[str, Any] = {}
        if _is_dict(legend_pos_raw):
            legend_pos = {str(k): v for k, v in legend_pos_raw.items()}

        legend_x = int(legend_pos["x"]) if "x" in legend_pos else None
        legend_y = int(legend_pos["y"]) if "y" in legend_pos else None

        return cls(
            values=values,
            size_emu=size_emu,
            show_legend=show_legend,
            legend_x=legend_x,
            legend_y=legend_y,
        )


# ---------------------------------------------------------------------------
# Cell helpers
# ---------------------------------------------------------------------------


def is_icon_cell(value: Any) -> bool:
    """Return True if *value* is an icon reference (``{icon: "X"}``)."""
    return isinstance(value, dict) and "icon" in value


def icon_cell_value(value: Any) -> Optional[str]:
    """Extract the icon name from an icon cell, or None."""
    if is_icon_cell(value):
        v = value.get("icon")
        return str(v) if v else None
    return None
