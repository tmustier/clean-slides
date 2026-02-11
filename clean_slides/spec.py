"""Table specification and layout dataclasses."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, TypedDict, Union

from typing_extensions import TypeGuard

from .constants import DefaultColors, Layout, TableDefaults
from .icons import (
    IconSet,
)
from .icons import (
    icon_cell_value as _icon_cell_value,
)
from .icons import (
    is_icon_cell as _is_icon_cell,
)

# Re-export icon helpers for convenience (used by other modules).
is_icon_cell = _is_icon_cell
icon_cell_value = _icon_cell_value

Box = tuple[int, int, int, int]  # x, y, w, h in EMU


# ---------------------------------------------------------------------------
# Chart cells — data model
# ---------------------------------------------------------------------------

# Pattern: "chartname-N" where chartname is [a-zA-Z_][a-zA-Z0-9_]* and N is 1+
_CHART_REF_RE = re.compile(r"^([a-zA-Z_][a-zA-Z0-9_]*)-(\d+)$")


@dataclass
class ChartDef:
    """Definition of a chart that can be embedded in table cells."""

    name: str
    type: str  # "bar" | "waterfall"
    dir: str  # "horizontal" | "vertical"
    values: list[float]
    format: str = "{}"
    color: str | None = None  # theme name or hex; None = template default
    colors: list[str | None] | None = None  # per-point colors (optional)
    label_position: str | None = None  # "above", "right", "on", "none"
    scale_max: float | None = None  # override automatic axis maximum
    scale_min: float | None = None  # override automatic axis minimum
    scale_group: str | None = None  # charts in the same group share axis scale

    # Waterfall-specific fields (ignored for type=bar)
    totals: list[int] | None = None  # 1-based indices that are total bars
    decreases: list[int] | None = None  # 1-based indices that are decrease bars
    total_color: str | None = None  # color for total bars
    connector: bool = True  # show connector lines between bars

    @classmethod
    def from_dict(cls, name: str, data: dict[str, Any]) -> ChartDef:
        raw_type = str(data.get("type", "bar"))
        if raw_type not in ("bar", "waterfall"):
            raise ValueError(f"Chart '{name}': type must be 'bar' or 'waterfall', got '{raw_type}'")

        raw_dir = str(data.get("dir", "vertical"))
        if raw_dir not in ("horizontal", "vertical"):
            raise ValueError(
                f"Chart '{name}': dir must be 'horizontal' or 'vertical', got '{raw_dir}'"
            )

        raw_values: object = data.get("values")
        if not _is_list(raw_values) or not raw_values:
            raise ValueError(f"Chart '{name}': values must be a non-empty list of numbers")
        values: list[float] = []
        for i, item in enumerate(raw_values):
            try:
                values.append(float(item))
            except (TypeError, ValueError) as err:
                raise ValueError(f"Chart '{name}': values[{i}] is not a number: {item!r}") from err

        fmt = str(data.get("format", "{}"))
        color_raw = data.get("color")
        color = str(color_raw) if color_raw is not None else None

        colors_raw: object = data.get("colors")
        colors: list[str | None] | None = None
        if colors_raw is not None:
            if not _is_list(colors_raw):
                raise ValueError(f"Chart '{name}': colors must be a list when provided")
            colors = []
            for item in colors_raw:
                if item is None:
                    colors.append(None)
                else:
                    colors.append(str(item))
            if len(colors) != len(values):
                raise ValueError(
                    f"Chart '{name}': colors must have {len(values)} entries to match values"
                )

        label_pos_raw = data.get("label_position")
        label_position = str(label_pos_raw) if label_pos_raw is not None else None
        scale_max_raw = data.get("scale_max")
        scale_max = float(scale_max_raw) if scale_max_raw is not None else None
        scale_min_raw = data.get("scale_min")
        scale_min = float(scale_min_raw) if scale_min_raw is not None else None
        scale_group_raw = data.get("scale_group")
        scale_group = str(scale_group_raw) if scale_group_raw is not None else None

        # Waterfall fields
        totals: list[int] | None = None
        decreases: list[int] | None = None
        total_color: str | None = None
        connector = True
        if raw_type == "waterfall":
            raw_totals: object = data.get("totals")
            if _is_list(raw_totals):
                totals = [int(x) for x in raw_totals]
            raw_decreases: object = data.get("decreases")
            if _is_list(raw_decreases):
                decreases = [int(x) for x in raw_decreases]
            tc_raw = data.get("total_color")
            total_color = str(tc_raw) if tc_raw is not None else None
            connector = bool(data.get("connector", True))

        return cls(
            name=name,
            type=raw_type,
            dir=raw_dir,
            values=values,
            format=fmt,
            color=color,
            colors=colors,
            label_position=label_position,
            scale_max=scale_max,
            scale_min=scale_min,
            scale_group=scale_group,
            totals=totals,
            decreases=decreases,
            total_color=total_color,
            connector=connector,
        )


@dataclass
class ChartRef:
    """A cell reference to a specific bar in a named chart.

    ``name`` identifies the :class:`ChartDef`; ``index`` is 1-based.
    """

    name: str
    index: int  # 1-based

    def __repr__(self) -> str:
        return f"{self.name}-{self.index}"


def parse_chart_ref(value: object) -> ChartRef | None:
    """Try to parse a cell value as a chart reference (``chartname-N``).

    Returns ``None`` if *value* is not a matching string.
    """
    if not isinstance(value, str):
        return None
    m = _CHART_REF_RE.match(value.strip())
    if m is None:
        return None
    return ChartRef(name=m.group(1), index=int(m.group(2)))


def is_chart_ref(value: object) -> bool:
    """Return True if *value* is a chart-ref string like ``bar1-3``."""
    return parse_chart_ref(value) is not None


def _parse_charts(data: dict[str, Any]) -> dict[str, ChartDef]:
    """Parse the top-level ``charts:`` mapping from YAML."""
    raw: object = data.get("charts")
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise ValueError("'charts' must be a mapping of chart-name → chart-definition")
    # raw is dict[Any, Any] from YAML; iterate with explicit typing.
    raw_dict: dict[str, Any] = {str(k): v for k, v in raw.items()}  # type: ignore[union-attr]
    charts: dict[str, ChartDef] = {}
    for name, raw_value in raw_dict.items():
        if not _is_dict(raw_value):
            raise ValueError(f"Chart '{name}': definition must be a mapping")
        charts[name] = ChartDef.from_dict(name, _stringify_keys(raw_value))
    _resolve_scale_groups(charts)
    return charts


def _resolve_scale_groups(charts: dict[str, ChartDef]) -> None:
    """Resolve ``scale_group``: charts in the same group share axis bounds.

    For each group, the effective ``scale_max`` is the maximum value across
    all member charts and the effective ``scale_min`` is the minimum value.
    Explicit per-chart ``scale_max`` / ``scale_min`` override the group value.
    """
    groups: dict[str, list[ChartDef]] = {}
    for chart in charts.values():
        if chart.scale_group is not None:
            groups.setdefault(chart.scale_group, []).append(chart)

    for members in groups.values():
        group_max = max(max(c.values) for c in members)
        group_min = min(min(c.values) for c in members)
        # Only set negative min if there are negative values
        effective_min = group_min if group_min < 0 else None
        for c in members:
            if c.scale_max is None:
                c.scale_max = group_max
            if c.scale_min is None and effective_min is not None:
                c.scale_min = effective_min


# ---------------------------------------------------------------------------
# Row / column overrides
# ---------------------------------------------------------------------------


@dataclass
class CellOverride:
    """Formatting override applied to all body cells in a row or column.

    Fields that are ``None`` are not overridden (cell keeps its default).
    """

    align: str | None = None  # "l", "ctr", "r"
    anchor: str | None = None  # "t", "ctr", "b"
    bold: bool | None = None
    color: str | None = None
    size: int | None = None  # pt
    font: str | None = None

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> CellOverride:
        raw_size = data.get("size")
        return cls(
            align=str(data["align"]) if "align" in data else None,
            anchor=str(data["anchor"]) if "anchor" in data else None,
            bold=bool(data["bold"]) if "bold" in data else None,
            color=str(data["color"]) if "color" in data else None,
            size=int(raw_size) if raw_size is not None else None,
            font=str(data["font"]) if "font" in data else None,
        )


def _parse_overrides(raw: object) -> dict[int, CellOverride]:
    """Parse ``row_overrides`` or ``col_overrides`` from YAML.

    Accepts ``{0: {align: ctr, ...}, 2: {bold: true}}`` mapping body-row
    (or body-column) indices to override dicts.
    """
    if not _is_dict(raw):
        return {}
    result: dict[int, CellOverride] = {}
    for key, value in raw.items():
        try:
            idx = int(key)
        except (TypeError, ValueError):
            continue
        if _is_dict(value):
            result[idx] = CellOverride.from_dict(_stringify_keys(value))
    return result


def _is_dict(value: object) -> TypeGuard[dict[Any, Any]]:
    return isinstance(value, dict)


def _is_list(value: object) -> TypeGuard[list[Any]]:
    return isinstance(value, list)


def _stringify_keys(value: dict[Any, Any]) -> dict[str, Any]:
    return {str(k): v for k, v in value.items()}


def _as_any_list(value: object) -> list[Any] | None:
    """Return *value* as a list preserving dicts/lists for rich content."""
    if not _is_list(value):
        return None
    return list(value)


def _as_float_list(value: object) -> list[float] | None:
    if not _is_list(value):
        return None

    out: list[float] = []
    for item in value:
        try:
            out.append(float(item))
        except (TypeError, ValueError):
            continue
    return out


def _as_cell_grid(value: object) -> list[list[Any]] | None:
    if not _is_list(value):
        return None

    grid: list[list[Any]] = []
    for row in value:
        if _is_list(row):
            grid.append(list(row))
        else:
            # Best-effort normalization. YAML `cells` should be a list-of-lists.
            grid.append([row])

    return grid


def _to_int(value: object) -> int:
    """Convert YAML-ish value to int."""
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        return int(value)
    return int(str(value))


class _HeaderColorOverrides(TypedDict, total=False):
    col_header_color: str
    col_superheader_color: str
    row_header_color: str
    row_superheader_color: str


# ---------------------------------------------------------------------------
# Content area
# ---------------------------------------------------------------------------


@dataclass
class ContentArea:
    """Defines the drawable area for the table in EMU."""

    x: int
    y: int
    width: int
    height: int

    @classmethod
    def from_layout(cls, layout: str = "default") -> ContentArea:
        """Build a content area from a named layout preset.

        For template-specific layouts (e.g. "3/4", "Contrast 1/3"), prefer
        ``_content_area_from_layout()`` in ``cli.py`` which reads the actual
        placeholder positions from the slide layout.
        """
        layout_key = layout.lower().strip()
        full_width = int(Layout.CONTENT_WIDTH)
        left = int(Layout.LEFT_MARGIN)

        # "full" starts at tracker zone — use when you want to fill entire slide
        if layout_key in {"full", "1/1"}:
            return cls(
                x=left,
                y=int(Layout.TRACKER_Y),
                width=full_width,
                height=int(Layout.FOOTER_LINE_Y - Layout.TRACKER_Y),
            )

        # "default" and "content" start below header line — safe with title/subtitle
        content_y = int(Layout.HEADER_LINE_Y + TableDefaults.CELL_PADDING)
        content_h = int(Layout.FOOTER_LINE_Y - content_y)

        if layout_key in {"default", "content", "body"}:
            return cls(x=left, y=content_y, width=full_width, height=content_h)

        raise ValueError(f"Unsupported layout '{layout}'")

    def contains(self, box: Box) -> bool:
        x, y, w, h = box
        return (
            x >= self.x
            and y >= self.y
            and x + w <= self.x + self.width
            and y + h <= self.y + self.height
        )


# ---------------------------------------------------------------------------
# Row groups (superheader support)
# ---------------------------------------------------------------------------


@dataclass
class ColSuperHeader:
    """A column superheader spanning one or more grid columns.

    Example: "" spanning the superheader column, "Details" spanning
    the Strategic-action + Key-details columns.
    """

    label: str
    span: int  # number of grid columns this header spans
    sub: str | None = None  # optional subtitle (non-bold, body text color)


@dataclass
class RowGroup:
    """A group of sub-rows sharing a superheader.

    Example: "Deploy" superheader spanning two sub-rows
    ("Co-locate..." and "Operate & maintain...").
    """

    header: Any
    num_rows: int  # how many sub-rows in this group
    promoted: bool = False  # header auto-promoted from first body cell


# ---------------------------------------------------------------------------
# Table spec
# ---------------------------------------------------------------------------


@dataclass
class TableSpec:
    """Logical table definition.

    Two modes:
      Flat     — cells + row_headers (groups is None)
      Grouped  — cells + groups (row_headers derived from groups)
    """

    num_rows: int  # total body rows (or total sub-rows when grouped)
    num_cols: int
    has_col_header: bool = True
    has_row_header: bool = False

    col_headers: list[Any] | None = None
    col_superheaders: list[ColSuperHeader] | None = None
    row_header_col_header: Any | None = None  # col header for the row-header column
    row_headers: list[Any] | None = None
    cells: list[list[Any]] | None = None

    groups: list[RowGroup] | None = None  # superheader groups

    col_widths: Union[None, str, list[float]] = None  # None (auto) | "equal" | list of floats
    body_default_lvl: int = 0
    parse_bullets: bool = True

    # Chart definitions — keyed by chart name.  Cell grid may contain
    # ``ChartRef`` objects that reference these.
    chart_defs: dict[str, ChartDef] = field(default_factory=lambda: dict[str, ChartDef]())

    # Header colors — defaults come from template-config.yaml default_colors section.
    # Individual specs can override per table.
    col_header_color: str | None = None
    col_superheader_color: str | None = None
    row_header_color: str | None = None
    row_superheader_color: str | None = None

    @property
    def effective_col_header_color(self) -> str:
        return self.col_header_color or DefaultColors.COL_HEADER

    @property
    def effective_col_superheader_color(self) -> str:
        return self.col_superheader_color or DefaultColors.COL_SUPERHEADER

    @property
    def effective_row_header_color(self) -> str:
        return self.row_header_color or DefaultColors.ROW_HEADER

    @property
    def effective_row_superheader_color(self) -> str:
        return self.row_superheader_color or DefaultColors.ROW_SUPERHEADER

    # Row / column overrides — keyed by body-row or body-column index
    row_overrides: dict[int, CellOverride] = field(
        default_factory=lambda: dict[int, CellOverride]()
    )
    col_overrides: dict[int, CellOverride] = field(
        default_factory=lambda: dict[int, CellOverride]()
    )

    # Icon indicators (traffic lights, RAG, etc.)
    icons: IconSet | None = None

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> TableSpec:
        raw_table: object = data.get("table", {})
        table: dict[str, Any] = _stringify_keys(raw_table) if _is_dict(raw_table) else {}

        raw_groups: object = table.get("row_groups")
        if _is_list(raw_groups) and raw_groups:
            return cls._from_groups(table, raw_groups, data)

        num_rows_raw: object = table.get("rows")
        num_cols_raw: object = table.get("cols")
        if num_rows_raw is None or num_cols_raw is None:
            raise ValueError("table.rows and table.cols are required")

        has_col_header = bool(table.get("has_col_header", True))
        has_row_header = bool(table.get("has_row_header", False))

        # YAML table.rows / table.cols are *total* grid dimensions (including
        # header rows/cols).  Internal num_rows / num_cols are body-only.
        total_rows = _to_int(num_rows_raw)
        total_cols = _to_int(num_cols_raw)

        icons = cls._parse_icons(data)

        col_superheaders = cls._parse_col_superheaders(table)
        header_rows = (1 if has_col_header else 0) + (1 if col_superheaders else 0)
        header_cols = 1 if has_row_header else 0

        cells = _as_cell_grid(table.get("cells"))

        # Auto-extract row headers from cells' first column when
        # has_row_header is set but no explicit row_headers list provided.
        explicit_row_headers = _as_any_list(table.get("row_headers"))
        if (
            has_row_header
            and explicit_row_headers is None
            and cells is not None
            and all(len(row) >= 2 for row in cells)
        ):
            explicit_row_headers = [row[0] for row in cells]
            cells = [row[1:] for row in cells]

        body_rows = total_rows - header_rows
        body_cols = total_cols - header_cols

        if body_rows <= 0:
            raise ValueError(
                f"table.rows={total_rows} is too small for configured header rows ({header_rows}); "
                "body rows must be > 0"
            )
        if body_cols <= 0:
            raise ValueError(
                f"table.cols={total_cols} is too small for configured header cols ({header_cols}); "
                "body cols must be > 0"
            )

        col_widths_parsed = cls._parse_column_widths(table)

        row_header_col_header: Any | None = table.get("row_header_col_header")

        col_headers = _as_any_list(table.get("col_headers"))
        # Convenience: allow the row-header column header to be included as the first col_header.
        if (
            has_row_header
            and col_headers
            and row_header_col_header is None
            and len(col_headers) == (body_cols + header_cols)
        ):
            row_header_col_header = col_headers[0]
            col_headers = col_headers[1:]

        body_default_lvl_raw: object = table.get("body_default_lvl", 0)
        body_default_lvl = _to_int(body_default_lvl_raw) if body_default_lvl_raw is not None else 0

        chart_defs = _parse_charts(data)
        if cells is not None and chart_defs:
            cells = _replace_chart_refs(cells)

        return cls(
            num_rows=body_rows,
            num_cols=body_cols,
            has_col_header=has_col_header,
            has_row_header=has_row_header,
            col_headers=col_headers,
            col_superheaders=col_superheaders,
            row_header_col_header=row_header_col_header,
            row_headers=explicit_row_headers,
            cells=cells,
            col_widths=col_widths_parsed,
            body_default_lvl=body_default_lvl,
            parse_bullets=bool(table.get("parse_bullets", True)),
            row_overrides=_parse_overrides(table.get("row_overrides")),
            col_overrides=_parse_overrides(table.get("col_overrides")),
            icons=icons,
            chart_defs=chart_defs,
            **cls._parse_header_colors(table),
        )

    @classmethod
    def _from_groups(
        cls,
        table: dict[str, Any],
        raw_groups: list[Any],
        data: dict[str, Any] | None = None,
    ) -> TableSpec:
        """Parse row_groups into a flat cell grid + group metadata."""
        num_cols_raw: object = table.get("cols")
        if num_cols_raw is None:
            raise ValueError("table.cols is required")

        total_cols = _to_int(num_cols_raw)
        has_col_header = bool(table.get("has_col_header", True))

        col_superheaders = cls._parse_col_superheaders(table)

        groups: list[RowGroup] = []
        all_rows: list[list[Any]] = []

        for group in raw_groups:
            if not _is_dict(group):
                continue
            g = _stringify_keys(group)

            header_raw: object = g.get("header", "")
            rows_raw: object = g.get("rows", [])

            group_rows: list[list[Any]] = []
            if _is_list(rows_raw):
                for row in rows_raw:
                    if _is_list(row):
                        group_rows.append(list(row))
                    else:
                        group_rows.append([row])

            def normalize_group_header(value: object) -> object:
                if _is_dict(value):
                    d = _stringify_keys(value)
                    if "text" in d or "sub" in d:
                        text = str(d.get("text", ""))
                        sub_raw: object = d.get("sub")
                        sub = str(sub_raw) if sub_raw is not None else None
                        return {"text": text, "sub": sub}
                if isinstance(value, str):
                    if "\n" in value:
                        first, rest = value.split("\n", 1)
                        sub = rest.strip()
                        # Heuristic: treat parenthesized continuation as subtitle.
                        if sub.startswith("("):
                            return {"text": first, "sub": sub}
                    return value
                return str(value)

            header_value = normalize_group_header(header_raw)
            header_text = ""
            if isinstance(header_value, str):
                header_text = header_value
            elif _is_dict(header_value):
                header_text = str(header_value.get("text", ""))

            promoted = False
            # Convenience: singleton groups with an empty header and chart refs
            # promote the first body cell into the group header. This avoids
            # blank row-header sections for start/total rows in waterfall tables.
            if not header_text.strip() and len(group_rows) == 1 and group_rows[0]:
                first_cell = group_rows[0][0]
                has_chart_ref = any(parse_chart_ref(cell) is not None for cell in group_rows[0])
                if (
                    has_chart_ref
                    and isinstance(first_cell, str)
                    and first_cell.strip()
                    and parse_chart_ref(first_cell) is None
                ):
                    header_value = normalize_group_header(first_cell)
                    group_rows[0][0] = ""
                    promoted = True

            groups.append(
                RowGroup(header=header_value, num_rows=len(group_rows), promoted=promoted)
            )
            all_rows.extend(group_rows)

        # groups imply a row-header (superheader) column
        header_cols = 1
        body_cols = total_cols - header_cols

        if body_cols <= 0:
            raise ValueError(
                f"table.cols={total_cols} is too small for configured header cols ({header_cols}); "
                "body cols must be > 0"
            )

        data_dict: dict[str, Any] = data if data is not None else {}
        icons = cls._parse_icons(data_dict)

        col_widths_parsed = cls._parse_column_widths(table)

        row_header_col_header: Any | None = table.get("row_header_col_header")

        col_headers = _as_any_list(table.get("col_headers"))
        if (
            col_headers
            and row_header_col_header is None
            and len(col_headers) == (body_cols + header_cols)
        ):
            row_header_col_header = col_headers[0]
            col_headers = col_headers[1:]

        body_default_lvl_raw: object = table.get("body_default_lvl", 0)
        body_default_lvl = _to_int(body_default_lvl_raw) if body_default_lvl_raw is not None else 0

        chart_defs = _parse_charts(data_dict)
        if all_rows and chart_defs:
            all_rows = _replace_chart_refs(all_rows)

        return cls(
            num_rows=len(all_rows),
            num_cols=body_cols,
            has_col_header=has_col_header,
            has_row_header=True,
            col_headers=col_headers,
            col_superheaders=col_superheaders,
            row_header_col_header=row_header_col_header,
            row_headers=[g.header for g in groups],
            cells=all_rows,
            groups=groups,
            col_widths=col_widths_parsed,
            body_default_lvl=body_default_lvl,
            parse_bullets=bool(table.get("parse_bullets", True)),
            row_overrides=_parse_overrides(table.get("row_overrides")),
            col_overrides=_parse_overrides(table.get("col_overrides")),
            icons=icons,
            chart_defs=chart_defs,
            **cls._parse_header_colors(table),
        )

    @property
    def is_grouped(self) -> bool:
        return self.groups is not None and len(self.groups) > 0

    @property
    def has_col_superheader(self) -> bool:
        return self.col_superheaders is not None and len(self.col_superheaders) > 0

    @property
    def row_offset(self) -> int:
        """Number of header grid rows before body rows."""
        return (1 if self.has_col_superheader else 0) + (1 if self.has_col_header else 0)

    @property
    def col_offset(self) -> int:
        """Number of header grid columns before body columns."""
        return 1 if self.has_row_header else 0

    @staticmethod
    def _parse_column_widths(table: dict[str, Any]) -> Union[None, str, list[float]]:
        """Parse ``column_widths`` from YAML.

        Returns:
            None          – auto (content-aware)
            ``"equal"``   – equal body-column widths
            list[float]   – explicit relative proportions
        """
        raw: object = table.get("column_widths")
        if raw is None:
            return None
        if isinstance(raw, str):
            return "equal" if raw.lower() == "equal" else None
        return _as_float_list(raw)

    @staticmethod
    def _parse_icons(data: dict[str, Any]) -> IconSet | None:
        """Parse top-level ``icons`` config."""
        raw: object = data.get("icons")
        if not _is_dict(raw):
            return None
        return IconSet.from_dict(_stringify_keys(raw))

    @staticmethod
    def _parse_header_colors(table: dict[str, Any]) -> _HeaderColorOverrides:
        """Parse per-element header colors from YAML."""
        colors: _HeaderColorOverrides = {}

        val = table.get("col_header_color")
        if val is not None:
            colors["col_header_color"] = str(val)

        val = table.get("col_superheader_color")
        if val is not None:
            colors["col_superheader_color"] = str(val)

        val = table.get("row_header_color")
        if val is not None:
            colors["row_header_color"] = str(val)

        val = table.get("row_superheader_color")
        if val is not None:
            colors["row_superheader_color"] = str(val)

        return colors

    @staticmethod
    def _parse_col_superheaders(table: dict[str, Any]) -> list[ColSuperHeader] | None:
        raw: object = table.get("col_superheaders")
        if not _is_list(raw) or not raw:
            return None

        headers: list[ColSuperHeader] = []
        for item in raw:
            if not _is_dict(item):
                continue
            d = _stringify_keys(item)
            label_raw: object = d.get("label", "")
            span_raw: object = d.get("span", 1)
            try:
                span = _to_int(span_raw) if span_raw is not None else 1
            except (TypeError, ValueError):
                span = 1
            label = str(label_raw)
            sub_raw: object = d.get("sub")
            sub = str(sub_raw) if sub_raw is not None else None
            headers.append(ColSuperHeader(label=label, span=span, sub=sub))

        return headers or None


# ---------------------------------------------------------------------------
# Chart ref replacement — convert matching strings in the cell grid
# ---------------------------------------------------------------------------


def _replace_chart_refs(cells: list[list[Any]]) -> list[list[Any]]:
    """Walk the cell grid and replace ``"chartname-N"`` strings with :class:`ChartRef`."""
    out: list[list[Any]] = []
    for row in cells:
        new_row: list[Any] = []
        for value in row:
            ref = parse_chart_ref(value)
            if ref is not None:
                new_row.append(ref)
            else:
                new_row.append(value)
        out.append(new_row)
    return out


# ---------------------------------------------------------------------------
# Layout result
# ---------------------------------------------------------------------------


def _cell_matrix() -> list[list[Box]]:
    cells: list[list[Box]] = []
    return cells


@dataclass
class TableLayout:
    """Computed layout information for a table."""

    col_widths: list[int]
    row_heights: list[int]  # one per grid row (header + all sub-rows)
    header_font_size: int  # 1/100 pt
    body_font_size: int  # 1/100 pt
    pad_top: int
    pad_bottom: int
    cells: list[list[Box]] = field(default_factory=_cell_matrix)

    def cell_box(self, r: int, c: int) -> Box:
        return self.cells[r][c]
