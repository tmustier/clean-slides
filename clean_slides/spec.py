"""Table specification and layout dataclasses."""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, TypedDict, Union

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

Box = Tuple[int, int, int, int]  # x, y, w, h in EMU


# ---------------------------------------------------------------------------
# Row / column overrides
# ---------------------------------------------------------------------------


@dataclass
class CellOverride:
    """Formatting override applied to all body cells in a row or column.

    Fields that are ``None`` are not overridden (cell keeps its default).
    """

    align: Optional[str] = None  # "l", "ctr", "r"
    anchor: Optional[str] = None  # "t", "ctr", "b"
    bold: Optional[bool] = None
    color: Optional[str] = None
    size: Optional[int] = None  # pt
    font: Optional[str] = None

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "CellOverride":
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


def _as_str_list(value: object) -> Optional[List[str]]:
    if not _is_list(value):
        return None
    return [str(item) for item in value]


def _as_any_list(value: object) -> Optional[List[Any]]:
    """Like ``_as_str_list`` but preserves dicts/lists for rich content."""
    if not _is_list(value):
        return None
    return list(value)


def _as_float_list(value: object) -> Optional[List[float]]:
    if not _is_list(value):
        return None

    out: List[float] = []
    for item in value:
        try:
            out.append(float(item))
        except (TypeError, ValueError):
            continue
    return out


def _as_cell_grid(value: object) -> Optional[List[List[Any]]]:
    if not _is_list(value):
        return None

    grid: List[List[Any]] = []
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
    def from_layout(cls, layout: str = "default") -> "ContentArea":
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


@dataclass
class RowGroup:
    """A group of sub-rows sharing a superheader.

    Example: "Deploy" superheader spanning two sub-rows
    ("Co-locate..." and "Operate & maintain...").
    """

    header: str
    num_rows: int  # how many sub-rows in this group


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

    col_headers: Optional[List[str]] = None
    col_superheaders: Optional[List[ColSuperHeader]] = None
    row_header_col_header: Optional[str] = None  # col header for the row-header column
    row_headers: Optional[List[Any]] = None
    cells: Optional[List[List[Any]]] = None

    groups: Optional[List[RowGroup]] = None  # superheader groups

    col_widths: Union[None, str, List[float]] = None  # None (auto) | "equal" | list of floats
    body_default_lvl: int = 0
    parse_bullets: bool = True

    # Header colors — defaults come from template-config.yaml default_colors section.
    # Individual specs can override per table.
    col_header_color: Optional[str] = None
    col_superheader_color: Optional[str] = None
    row_header_color: Optional[str] = None
    row_superheader_color: Optional[str] = None

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
    row_overrides: Dict[int, CellOverride] = field(
        default_factory=lambda: dict[int, CellOverride]()
    )
    col_overrides: Dict[int, CellOverride] = field(
        default_factory=lambda: dict[int, CellOverride]()
    )

    # Icon indicators (traffic lights, RAG, etc.)
    icons: Optional[IconSet] = None

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TableSpec":
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

        # NOTE: YAML table.rows / table.cols are interpreted as *total* rows/cols,
        # including enabled header rows/cols.
        #
        # Internal TableSpec.num_rows / num_cols remain *body-only* dimensions.
        #
        # Back-compat: if the provided `cells` shape matches the old semantics,
        # we auto-detect and accept it.
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

        row_header_col_header_raw: object = table.get("row_header_col_header")
        row_header_col_header = (
            str(row_header_col_header_raw) if row_header_col_header_raw is not None else None
        )

        col_headers = _as_str_list(table.get("col_headers"))
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
            **cls._parse_header_colors(table),
        )

    @classmethod
    def _from_groups(
        cls,
        table: dict[str, Any],
        raw_groups: list[Any],
        data: Optional[dict[str, Any]] = None,
    ) -> "TableSpec":
        """Parse row_groups into a flat cell grid + group metadata."""
        num_cols_raw: object = table.get("cols")
        if num_cols_raw is None:
            raise ValueError("table.cols is required")

        total_cols = _to_int(num_cols_raw)
        has_col_header = bool(table.get("has_col_header", True))

        col_superheaders = cls._parse_col_superheaders(table)

        groups: List[RowGroup] = []
        all_rows: List[List[Any]] = []

        for group in raw_groups:
            if not _is_dict(group):
                continue
            g = _stringify_keys(group)

            header_raw: object = g.get("header", "")
            rows_raw: object = g.get("rows", [])

            group_rows: List[List[Any]] = []
            if _is_list(rows_raw):
                for row in rows_raw:
                    if _is_list(row):
                        group_rows.append(list(row))
                    else:
                        group_rows.append([row])

            groups.append(RowGroup(header=str(header_raw), num_rows=len(group_rows)))
            all_rows.extend(group_rows)

        # groups imply a row-header (superheader) column
        header_cols = 1
        body_cols = total_cols - header_cols

        # Back-compat: old semantics counted body cols only.
        max_len = max((len(r) for r in all_rows), default=0)
        if max_len == total_cols:
            body_cols = total_cols

        if body_cols <= 0:
            raise ValueError(
                f"table.cols={total_cols} is too small for configured header cols ({header_cols}); "
                "body cols must be > 0"
            )

        data_dict: dict[str, Any] = data if data is not None else {}
        icons = cls._parse_icons(data_dict)

        col_widths_parsed = cls._parse_column_widths(table)

        row_header_col_header_raw: object = table.get("row_header_col_header")
        row_header_col_header = (
            str(row_header_col_header_raw) if row_header_col_header_raw is not None else None
        )

        col_headers = _as_str_list(table.get("col_headers"))
        if (
            col_headers
            and row_header_col_header is None
            and len(col_headers) == (body_cols + header_cols)
        ):
            row_header_col_header = col_headers[0]
            col_headers = col_headers[1:]

        body_default_lvl_raw: object = table.get("body_default_lvl", 0)
        body_default_lvl = _to_int(body_default_lvl_raw) if body_default_lvl_raw is not None else 0

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
    def _parse_column_widths(table: dict[str, Any]) -> Union[None, str, List[float]]:
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
    def _parse_icons(data: dict[str, Any]) -> Optional[IconSet]:
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
    def _parse_col_superheaders(table: dict[str, Any]) -> Optional[List[ColSuperHeader]]:
        raw: object = table.get("col_superheaders")
        if not _is_list(raw) or not raw:
            return None

        headers: List[ColSuperHeader] = []
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
            headers.append(ColSuperHeader(label=str(label_raw), span=span))

        return headers or None


# ---------------------------------------------------------------------------
# Layout result
# ---------------------------------------------------------------------------


def _cell_matrix() -> List[List[Box]]:
    cells: List[List[Box]] = []
    return cells


@dataclass
class TableLayout:
    """Computed layout information for a table."""

    col_widths: List[int]
    row_heights: List[int]  # one per grid row (header + all sub-rows)
    header_font_size: int  # 1/100 pt
    body_font_size: int  # 1/100 pt
    pad_top: int
    pad_bottom: int
    cells: List[List[Box]] = field(default_factory=_cell_matrix)

    def cell_box(self, r: int, c: int) -> Box:
        return self.cells[r][c]
