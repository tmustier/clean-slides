"""Chart overlay and manual label layout helpers."""

# pyright: reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false
# pyright: reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false
# pyright: reportMissingParameterType=false
# pyright: reportMissingTypeArgument=false
# pyright: reportGeneralTypeIssues=false
# pyright: reportArgumentType=false
# pyright: reportCallIssue=false
# pyright: reportAttributeAccessIssue=false
# pyright: reportIndexIssue=false
# pyright: reportOperatorIssue=false
# pyright: reportUnknownLambdaType=false
# pyright: reportPossiblyUnboundVariable=false

from __future__ import annotations

from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement

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
from .spec_utils import (
    format_label,
    numeric_value,
    safe_value,
)


def get_chart_series(chart_space) -> list[OxmlElement]:
    chart = chart_space.find(qn("c:chart"))
    if chart is None:
        return []
    plot_area = chart.find(qn("c:plotArea"))
    if plot_area is None:
        return []
    bar_chart = plot_area.find(qn("c:barChart"))
    if bar_chart is None:
        return []
    return bar_chart.findall(qn("c:ser"))


def measure_label_width(text: str) -> int:
    length = max(1, len(text))
    return int(DEFAULT_WATERFALL_LABEL_WIDTH_BASE + DEFAULT_WATERFALL_LABEL_WIDTH_PER_CHAR * length)


def apply_waterfall_data_label_layout(chart, chart_box: tuple, meta: dict) -> None:
    overlay = meta.get("overlay") if meta else None
    if not overlay:
        return

    categories = overlay.get("categories", [])
    chart_series = overlay.get("chart_series", [])
    segment_values = overlay.get("segment_values", {})
    label_bottoms = overlay.get("label_bottoms", [])
    orientation = normalize_orientation(meta.get("orientation"))
    offset_indices = set(meta.get("offset_label_indices") or [])

    axis_min = meta.get("axis_min", 0)
    axis_max = meta.get("axis_max", 0)
    gap_width = int(meta.get("gap_width", 80))
    plot_layout = meta.get("plot_layout") or DEFAULT_WATERFALL_PLOT_LAYOUT

    geometry = compute_category_geometry(
        chart_box,
        plot_layout,
        categories,
        gap_width,
        orientation,
    )
    plot_left = geometry["plot_left"]
    plot_top = geometry["plot_top"]
    plot_width = geometry["plot_width"]
    plot_height = geometry["plot_height"]

    if orientation == "horizontal":
        bar_span = geometry["bar_height"]
    else:
        bar_span = geometry["bar_width"]

    inside_offset_x = bar_span * DEFAULT_WATERFALL_DLABEL_INSIDE_OFFSET_RATIO
    outside_offset_x = bar_span * DEFAULT_WATERFALL_DLABEL_OUTSIDE_OFFSET_RATIO
    outside_spacing = (
        DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT * DEFAULT_WATERFALL_DLABEL_OUTSIDE_SPACING_RATIO
    )
    if orientation == "horizontal":
        default_dy = plot_height * DEFAULT_WATERFALL_DLABEL_Y_OFFSET_RATIO_HORIZONTAL
    else:
        default_dy = plot_height * DEFAULT_WATERFALL_DLABEL_Y_OFFSET_RATIO

    data_label_decimals = meta.get("data_label_decimals")
    if data_label_decimals is None:
        data_label_decimals = 0

    # Build label layout per category
    label_layout: dict[tuple[int, int], dict[str, float]] = {}
    hide_labels: set[tuple[int, int]] = set()

    for cat_idx in range(len(categories)):
        base_val = label_bottoms[cat_idx] if cat_idx < len(label_bottoms) else None
        if base_val is None:
            base_val = 0.0
        current = base_val
        labels = []

        for series_order, series_name in enumerate(chart_series):
            values = segment_values.get(series_name, [])
            value = numeric_value(safe_value(values, cat_idx))
            if value is None:
                continue
            magnitude = abs(value)
            if magnitude == 0:
                # suppress zero labels
                series_idx = series_order + 1
                hide_labels.add((series_idx, cat_idx))
                continue

            seg_bottom = current
            seg_top = current + magnitude
            current = seg_top

            if orientation == "horizontal":
                x_start = value_to_x(seg_bottom, axis_min, axis_max, plot_left, plot_width)
                x_end = value_to_x(seg_top, axis_min, axis_max, plot_left, plot_width)
                span = abs(x_end - x_start)
                x_center = (x_start + x_end) / 2
                y_center = geometry["bar_centers"][cat_idx]
            else:
                y_top = value_to_y(seg_top, axis_min, axis_max, plot_top, plot_height)
                y_bottom = value_to_y(seg_bottom, axis_min, axis_max, plot_top, plot_height)
                span = abs(y_bottom - y_top)
                x_center = geometry["bar_centers"][cat_idx]
                y_center = (y_top + y_bottom) / 2

            text = format_label(magnitude, decimals=int(data_label_decimals))
            label_width = measure_label_width(text)

            labels.append(
                {
                    "series_idx": series_order + 1,
                    "cat_idx": cat_idx,
                    "span": span,
                    "x_center": x_center,
                    "y_center": y_center,
                    "width": label_width,
                    "height": DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT,
                }
            )

        inside_labels = []
        outside_labels = []
        for label in labels:
            if orientation == "horizontal":
                threshold = label["width"] * DEFAULT_WATERFALL_DLABEL_MIN_INSIDE_RATIO
            else:
                threshold = label["height"] * DEFAULT_WATERFALL_DLABEL_MIN_INSIDE_RATIO
            if label["span"] < threshold:
                outside_labels.append(label)
            else:
                inside_labels.append(label)

        # Inside labels: shift horizontally if overlapping (prefer shifting bottom label)
        if orientation == "vertical" and len(inside_labels) > 1:
            sorted_labels = sorted(inside_labels, key=lambda item: item["y_center"])
            overlap = any(
                abs(sorted_labels[i]["y_center"] - sorted_labels[i - 1]["y_center"])
                < DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT
                for i in range(1, len(sorted_labels))
            )
            if overlap:
                sorted_labels[-1]["dx"] = -inside_offset_x

        # Outside labels: move to the right and separate vertically
        if outside_labels and orientation == "vertical":
            sorted_labels = sorted(outside_labels, key=lambda item: item["y_center"])
            if len(sorted_labels) == 2:
                sorted_labels[0]["dx"] = outside_offset_x
                sorted_labels[0]["dy"] = (
                    DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT
                    * DEFAULT_WATERFALL_DLABEL_OUTSIDE_TOP_RATIO
                )
                sorted_labels[1]["dx"] = outside_offset_x
                sorted_labels[1]["dy"] = (
                    DEFAULT_WATERFALL_VALUE_LABEL_HEIGHT
                    * DEFAULT_WATERFALL_DLABEL_OUTSIDE_BOTTOM_RATIO
                )
            else:
                count = len(sorted_labels)
                for idx, label in enumerate(sorted_labels):
                    offset_index = idx - (count - 1) / 2
                    label["dx"] = outside_offset_x
                    label["dy"] = offset_index * outside_spacing

        # Horizontal outside labels: move below the bar with spacing
        if outside_labels and orientation == "horizontal":
            sorted_labels = sorted(outside_labels, key=lambda item: item["x_center"])
            count = len(sorted_labels)
            for idx, label in enumerate(sorted_labels):
                offset_index = idx - (count - 1) / 2
                label["dy"] = bar_span * 0.75 + offset_index * outside_spacing

        for label in labels:
            dx = label.get("dx", 0.0)
            dy = label.get("dy", 0.0)
            dy += default_dy
            key = (label["series_idx"], label["cat_idx"])
            label_layout[key] = {
                "x": dx / plot_width if plot_width else 0,
                "y": dy / plot_height if plot_height else 0,
            }

    series_elements = get_chart_series(chart._chartSpace)
    if not series_elements:
        return

    def ensure_dlbls(ser):
        dlbls = ser.find(qn("c:dLbls"))
        if dlbls is None:
            dlbls = OxmlElement("c:dLbls")
            insert_at = ser.find(qn("c:val"))
            if insert_at is not None:
                ser.insert(ser.index(insert_at), dlbls)
            else:
                ser.append(dlbls)
        return dlbls

    def set_child_val(parent, tag, value):
        elem = parent.find(qn(tag))
        if elem is None:
            elem = OxmlElement(tag)
            parent.append(elem)
        elem.set("val", str(value))
        return elem

    def add_dlbl(dlbls, point_idx, show_val=True, manual_x=None, manual_y=None):
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
        set_child_val(dlbl, "c:dLblPos", "ctr")
        set_child_val(dlbl, "c:showLegendKey", 0)
        set_child_val(dlbl, "c:showVal", 1 if show_val else 0)
        set_child_val(dlbl, "c:showCatName", 0)
        set_child_val(dlbl, "c:showSerName", 0)
        set_child_val(dlbl, "c:showPercent", 0)
        set_child_val(dlbl, "c:showBubbleSize", 0)
        dlbls.append(dlbl)
        return dlbl

    # Offset series labels (row0) for reused start values
    offset_idx = meta.get("offset_series_idx", 0)
    if 0 <= offset_idx < len(series_elements):
        ser = series_elements[offset_idx]
        dlbls = ensure_dlbls(ser)
        # clear existing point labels
        for child in list(dlbls):
            if child.tag == qn("c:dLbl"):
                dlbls.remove(child)
        set_child_val(dlbls, "c:dLblPos", "ctr")
        set_child_val(dlbls, "c:showLegendKey", 0)
        set_child_val(dlbls, "c:showVal", 0)
        set_child_val(dlbls, "c:showCatName", 0)
        set_child_val(dlbls, "c:showSerName", 0)
        set_child_val(dlbls, "c:showPercent", 0)
        set_child_val(dlbls, "c:showBubbleSize", 0)

        for idx in sorted(offset_indices):
            key = (offset_idx, idx)
            layout = label_layout.get(key)
            manual_x = layout.get("x") if layout else 0
            manual_y = layout.get("y") if layout else default_dy / plot_height
            add_dlbl(dlbls, idx, show_val=True, manual_x=manual_x, manual_y=manual_y)

    # Non-offset series label overrides
    for series_idx, ser in enumerate(series_elements):
        if series_idx == offset_idx:
            continue
        dlbls = ensure_dlbls(ser)
        set_child_val(dlbls, "c:dLblPos", "ctr")
        set_child_val(dlbls, "c:showVal", 1)
        for key, layout in label_layout.items():
            if key[0] != series_idx:
                continue
            add_dlbl(
                dlbls, key[1], show_val=True, manual_x=layout.get("x"), manual_y=layout.get("y")
            )
        for hide_key in hide_labels:
            if hide_key[0] != series_idx:
                continue
            add_dlbl(dlbls, hide_key[1], show_val=False, manual_x=None, manual_y=None)
