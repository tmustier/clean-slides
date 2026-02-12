"""Chart overlay and manual label layout helpers."""

from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from typing import Protocol, Union, cast

from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement

from ..pptx_access import chart_xml_space
from .defaults import (
    DEFAULT_WATERFALL_DLABEL_INSIDE_OFFSET_RATIO,
    DEFAULT_WATERFALL_DLABEL_MIN_INSIDE_RATIO,
    DEFAULT_WATERFALL_DLABEL_OUTSIDE_BOTTOM_RATIO,
    DEFAULT_WATERFALL_DLABEL_OUTSIDE_OFFSET_RATIO,
    DEFAULT_WATERFALL_DLABEL_OUTSIDE_SPACING_RATIO,
    DEFAULT_WATERFALL_DLABEL_OUTSIDE_TOP_RATIO,
    DEFAULT_WATERFALL_DLABEL_Y_OFFSET_RATIO,
    DEFAULT_WATERFALL_DLABEL_Y_OFFSET_RATIO_HORIZONTAL,
    DEFAULT_WATERFALL_LABEL_WIDTH_BASE,
    DEFAULT_WATERFALL_LABEL_WIDTH_PER_CHAR,
    DEFAULT_WATERFALL_PLOT_LAYOUT,
    DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT,
)
from .geometry import (
    compute_category_geometry,
    normalize_orientation,
    value_to_x,
    value_to_y,
)
from .spec_utils import format_label, numeric_value, safe_value

FloatOrNone = Union[float, None]
ChartBox = tuple[int, int, int, int]
LabelKey = tuple[int, int]
LabelLayout = dict[str, float]


class _XmlElementLike(Protocol):
    tag: str

    def find(self, path: str) -> object | None: ...

    def findall(self, path: str) -> list[object]: ...

    def append(self, element: object) -> None: ...

    def insert(self, index: int, element: object) -> None: ...

    def index(self, element: object) -> int: ...

    def remove(self, element: object) -> None: ...

    def set(self, key: str, value: str) -> None: ...

    def __iter__(self) -> Iterator[object]: ...


@dataclass
class SegmentLabel:
    series_idx: int
    cat_idx: int
    span: float
    x_center: float
    y_center: float
    width: int
    height: int
    dx: float = 0.0
    dy: float = 0.0


def _mapping(value: object) -> dict[str, object]:
    if not isinstance(value, dict):
        return {}

    typed_mapping = cast(dict[object, object], value)
    mapped: dict[str, object] = {}
    for key, item in typed_mapping.items():
        if isinstance(key, str):
            mapped[key] = item
    return mapped


def _list(value: object) -> list[object]:
    if not isinstance(value, list):
        return []
    return cast(list[object], value)


def _string_list(value: object) -> list[str]:
    result: list[str] = []
    for item in _list(value):
        if isinstance(item, str):
            result.append(item)
    return result


def _segment_values(value: object) -> dict[str, list[object]]:
    if not isinstance(value, dict):
        return {}

    typed_mapping = cast(dict[object, object], value)
    values: dict[str, list[object]] = {}
    for key, item in typed_mapping.items():
        if isinstance(key, str):
            values[key] = _list(item)
    return values


def _float(value: object, default: float = 0.0) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    return default


def _int(value: object, default: int = 0) -> int:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    return default


def _float_or_none(value: object) -> FloatOrNone:
    if isinstance(value, (int, float)):
        return float(value)
    return None


def _float_or_none_list(value: object) -> list[FloatOrNone]:
    values: list[FloatOrNone] = []
    for item in _list(value):
        values.append(_float_or_none(item))
    return values


def _index_set(value: object) -> set[int]:
    indices: set[int] = set()
    if isinstance(value, set):
        for item in cast(set[object], value):
            if isinstance(item, int):
                indices.add(item)
        return indices

    for item in _list(value):
        if isinstance(item, int):
            indices.add(item)
    return indices


def _plot_layout(value: object) -> dict[str, float]:
    if isinstance(value, dict):
        typed_mapping = cast(dict[object, object], value)
        layout = {
            key: float(item)
            for key, item in typed_mapping.items()
            if isinstance(key, str) and isinstance(item, (int, float))
        }
        if layout:
            return layout

    return dict(DEFAULT_WATERFALL_PLOT_LAYOUT)


def _geometry_float(geometry: Mapping[str, object], key: str) -> float:
    return _float(geometry.get(key), 0.0)


def _geometry_series(geometry: Mapping[str, object], key: str) -> list[float]:
    values = geometry.get(key)
    if not isinstance(values, list):
        return []

    typed_values = cast(list[object], values)
    result: list[float] = []
    for item in typed_values:
        if isinstance(item, (int, float)):
            result.append(float(item))
    return result


def _label_at(values: Sequence[FloatOrNone], index: int) -> FloatOrNone:
    if index < 0 or index >= len(values):
        return None
    return values[index]


def get_chart_series(chart_space: object) -> list[object]:
    chart_space_el = cast(_XmlElementLike, chart_space)

    chart_obj = chart_space_el.find(qn("c:chart"))
    if chart_obj is None:
        return []
    chart_el = cast(_XmlElementLike, chart_obj)

    plot_area_obj = chart_el.find(qn("c:plotArea"))
    if plot_area_obj is None:
        return []
    plot_area = cast(_XmlElementLike, plot_area_obj)

    bar_chart_obj = plot_area.find(qn("c:barChart"))
    if bar_chart_obj is None:
        return []
    bar_chart = cast(_XmlElementLike, bar_chart_obj)

    return list(bar_chart.findall(qn("c:ser")))


def measure_label_width(text: str) -> int:
    length = max(1, len(text))
    return int(DEFAULT_WATERFALL_LABEL_WIDTH_BASE + DEFAULT_WATERFALL_LABEL_WIDTH_PER_CHAR * length)


def _ensure_dlbls(series_element: object) -> object:
    series_el = cast(_XmlElementLike, series_element)
    dlbls_obj = series_el.find(qn("c:dLbls"))
    if dlbls_obj is not None:
        return dlbls_obj

    dlbls = OxmlElement("c:dLbls")
    insert_at_obj = series_el.find(qn("c:val"))
    if insert_at_obj is not None:
        series_el.insert(series_el.index(insert_at_obj), dlbls)
    else:
        series_el.append(dlbls)
    return dlbls


def _set_child_val(parent: object, tag: str, value: Union[str, int]) -> object:
    parent_el = cast(_XmlElementLike, parent)
    elem_obj = parent_el.find(qn(tag))
    if elem_obj is None:
        elem_obj = OxmlElement(tag)
        parent_el.append(elem_obj)

    elem = cast(_XmlElementLike, elem_obj)
    elem.set("val", str(value))
    return elem


def _add_dlbl(
    dlbls: object,
    point_idx: int,
    show_val: bool = True,
    manual_x: Union[float, None] = None,
    manual_y: Union[float, None] = None,
) -> object:
    dlbl = OxmlElement("c:dLbl")
    idx_el = OxmlElement("c:idx")
    idx_el.set("val", str(point_idx))
    dlbl.append(idx_el)

    if manual_x is not None or manual_y is not None:
        layout_el = OxmlElement("c:layout")
        manual_el = OxmlElement("c:manualLayout")
        if manual_x is not None:
            x_el = OxmlElement("c:x")
            x_el.set("val", str(manual_x))
            manual_el.append(x_el)
        if manual_y is not None:
            y_el = OxmlElement("c:y")
            y_el.set("val", str(manual_y))
            manual_el.append(y_el)
        layout_el.append(manual_el)
        dlbl.append(layout_el)

    _set_child_val(dlbl, "c:dLblPos", "ctr")
    _set_child_val(dlbl, "c:showLegendKey", 0)
    _set_child_val(dlbl, "c:showVal", 1 if show_val else 0)
    _set_child_val(dlbl, "c:showCatName", 0)
    _set_child_val(dlbl, "c:showSerName", 0)
    _set_child_val(dlbl, "c:showPercent", 0)
    _set_child_val(dlbl, "c:showBubbleSize", 0)

    dlbls_el = cast(_XmlElementLike, dlbls)
    dlbls_el.append(dlbl)
    return dlbl


def apply_waterfall_data_label_layout(
    chart: object,
    chart_box: ChartBox,
    meta: Mapping[str, object],
) -> None:
    overlay = _mapping(meta.get("overlay")) if meta else {}
    if not overlay:
        return

    categories = _list(overlay.get("categories"))
    chart_series = _string_list(overlay.get("chart_series"))
    segment_values = _segment_values(overlay.get("segment_values"))
    label_bottoms = _float_or_none_list(overlay.get("label_bottoms"))

    orientation_raw = meta.get("orientation")
    orientation = normalize_orientation(
        orientation_raw if isinstance(orientation_raw, str) else None
    )
    offset_indices = _index_set(meta.get("offset_label_indices"))

    axis_min = _float(meta.get("axis_min"), 0.0)
    axis_max = _float(meta.get("axis_max"), 0.0)
    gap_width = _int(meta.get("gap_width"), 80)
    plot_layout = _plot_layout(meta.get("plot_layout"))

    geometry = compute_category_geometry(
        chart_box,
        plot_layout,
        categories,
        gap_width,
        orientation,
    )

    plot_left = _geometry_float(geometry, "plot_left")
    plot_top = _geometry_float(geometry, "plot_top")
    plot_width = _geometry_float(geometry, "plot_width")
    plot_height = _geometry_float(geometry, "plot_height")
    bar_centers = _geometry_series(geometry, "bar_centers")

    bar_span = (
        _geometry_float(geometry, "bar_height")
        if orientation == "horizontal"
        else _geometry_float(geometry, "bar_width")
    )

    inside_offset_x = bar_span * DEFAULT_WATERFALL_DLABEL_INSIDE_OFFSET_RATIO
    outside_offset_x = bar_span * DEFAULT_WATERFALL_DLABEL_OUTSIDE_OFFSET_RATIO
    outside_spacing = (
        int(DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT) * DEFAULT_WATERFALL_DLABEL_OUTSIDE_SPACING_RATIO
    )

    if orientation == "horizontal":
        default_dy = plot_height * DEFAULT_WATERFALL_DLABEL_Y_OFFSET_RATIO_HORIZONTAL
    else:
        default_dy = plot_height * DEFAULT_WATERFALL_DLABEL_Y_OFFSET_RATIO

    data_label_decimals = _int(meta.get("data_label_decimals"), 0)

    label_layout: dict[LabelKey, LabelLayout] = {}
    hide_labels: set[LabelKey] = set()

    for cat_idx in range(len(categories)):
        base_val = _label_at(label_bottoms, cat_idx)
        current = 0.0 if base_val is None else base_val
        labels: list[SegmentLabel] = []

        for series_order, series_name in enumerate(chart_series):
            values = segment_values.get(series_name, [])
            value = numeric_value(safe_value(values, cat_idx))
            if value is None:
                continue

            magnitude = abs(value)
            if magnitude == 0:
                hide_labels.add((series_order + 1, cat_idx))
                continue

            seg_bottom = current
            seg_top = current + magnitude
            current = seg_top

            if orientation == "horizontal":
                x_start = value_to_x(seg_bottom, axis_min, axis_max, plot_left, plot_width)
                x_end = value_to_x(seg_top, axis_min, axis_max, plot_left, plot_width)
                span = abs(x_end - x_start)
                x_center = (x_start + x_end) / 2
                y_center = bar_centers[cat_idx] if cat_idx < len(bar_centers) else 0.0
            else:
                y_top = value_to_y(seg_top, axis_min, axis_max, plot_top, plot_height)
                y_bottom = value_to_y(seg_bottom, axis_min, axis_max, plot_top, plot_height)
                span = abs(y_bottom - y_top)
                x_center = bar_centers[cat_idx] if cat_idx < len(bar_centers) else 0.0
                y_center = (y_top + y_bottom) / 2

            text = format_label(magnitude, decimals=data_label_decimals)
            if text is None:
                continue

            labels.append(
                SegmentLabel(
                    series_idx=series_order + 1,
                    cat_idx=cat_idx,
                    span=span,
                    x_center=x_center,
                    y_center=y_center,
                    width=measure_label_width(text),
                    height=int(DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT),
                )
            )

        inside_labels: list[SegmentLabel] = []
        outside_labels: list[SegmentLabel] = []

        for label in labels:
            if orientation == "horizontal":
                threshold = label.width * DEFAULT_WATERFALL_DLABEL_MIN_INSIDE_RATIO
            else:
                threshold = label.height * DEFAULT_WATERFALL_DLABEL_MIN_INSIDE_RATIO

            if label.span < threshold:
                outside_labels.append(label)
            else:
                inside_labels.append(label)

        if orientation == "vertical" and len(inside_labels) > 1:
            sorted_labels = sorted(inside_labels, key=lambda item: item.y_center)
            overlap = any(
                abs(sorted_labels[idx].y_center - sorted_labels[idx - 1].y_center)
                < int(DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT)
                for idx in range(1, len(sorted_labels))
            )
            if overlap:
                sorted_labels[-1].dx = -inside_offset_x

        if outside_labels and orientation == "vertical":
            sorted_labels = sorted(outside_labels, key=lambda item: item.y_center)
            if len(sorted_labels) == 2:
                sorted_labels[0].dx = outside_offset_x
                sorted_labels[0].dy = (
                    int(DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT)
                    * DEFAULT_WATERFALL_DLABEL_OUTSIDE_TOP_RATIO
                )
                sorted_labels[1].dx = outside_offset_x
                sorted_labels[1].dy = (
                    int(DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT)
                    * DEFAULT_WATERFALL_DLABEL_OUTSIDE_BOTTOM_RATIO
                )
            else:
                count = len(sorted_labels)
                for idx, label in enumerate(sorted_labels):
                    offset_index = idx - (count - 1) / 2
                    label.dx = outside_offset_x
                    label.dy = offset_index * outside_spacing

        if outside_labels and orientation == "horizontal":
            sorted_labels = sorted(outside_labels, key=lambda item: item.x_center)
            count = len(sorted_labels)
            for idx, label in enumerate(sorted_labels):
                offset_index = idx - (count - 1) / 2
                label.dy = bar_span * 0.75 + offset_index * outside_spacing

        for label in labels:
            dx = label.dx
            dy = label.dy + default_dy
            key = (label.series_idx, label.cat_idx)
            label_layout[key] = {
                "x": dx / plot_width if plot_width else 0.0,
                "y": dy / plot_height if plot_height else 0.0,
            }

    chart_space = chart_xml_space(chart)
    if chart_space is None:
        return

    series_elements = get_chart_series(chart_space)
    if not series_elements:
        return

    offset_idx = _int(meta.get("offset_series_idx"), 0)

    if 0 <= offset_idx < len(series_elements):
        series_element = series_elements[offset_idx]
        dlbls = _ensure_dlbls(series_element)
        dlbls_el = cast(_XmlElementLike, dlbls)

        for child_obj in list(dlbls_el):
            child = cast(_XmlElementLike, child_obj)
            if child.tag == qn("c:dLbl"):
                dlbls_el.remove(child)

        _set_child_val(dlbls_el, "c:dLblPos", "ctr")
        _set_child_val(dlbls_el, "c:showLegendKey", 0)
        _set_child_val(dlbls_el, "c:showVal", 0)
        _set_child_val(dlbls_el, "c:showCatName", 0)
        _set_child_val(dlbls_el, "c:showSerName", 0)
        _set_child_val(dlbls_el, "c:showPercent", 0)
        _set_child_val(dlbls_el, "c:showBubbleSize", 0)

        default_manual_y = default_dy / plot_height if plot_height else 0.0

        for idx in sorted(offset_indices):
            layout = label_layout.get((offset_idx, idx))
            manual_x = layout.get("x", 0.0) if layout else 0.0
            manual_y = layout.get("y", default_manual_y) if layout else default_manual_y
            _add_dlbl(
                dlbls_el,
                idx,
                show_val=True,
                manual_x=manual_x,
                manual_y=manual_y,
            )

    for series_idx, series_element in enumerate(series_elements):
        if series_idx == offset_idx:
            continue

        dlbls = _ensure_dlbls(series_element)
        _set_child_val(dlbls, "c:dLblPos", "ctr")
        _set_child_val(dlbls, "c:showVal", 1)

        for key, layout in label_layout.items():
            if key[0] != series_idx:
                continue
            _add_dlbl(
                dlbls,
                key[1],
                show_val=True,
                manual_x=layout.get("x"),
                manual_y=layout.get("y"),
            )

        for hidden_key in hide_labels:
            if hidden_key[0] != series_idx:
                continue
            _add_dlbl(
                dlbls,
                hidden_key[1],
                show_val=False,
                manual_x=None,
                manual_y=None,
            )
