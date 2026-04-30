#!

from typing import Any

from .colors import normalize_color_value
from .types import FormatDict
from .utils import map_style_value


# ================ [ XLSXWRITER SPECIAL FORMATTING SECTION ] =================

def _normalize_color_xlsxwriter(value: Any) -> str | None:
    """
    Accepts a Boolean color value (HEX string, CSS name, etc.)
    and returns the string '#RRGGBB' for xlsxwriter.
    """
    if value is None:
        return None
    # we want '#RRGGBB'
    hex_rgb = normalize_color_value(value, with_alpha=False)  # 'RRGGBB'
    if hex_rgb is None:
        return None
    return "#" + hex_rgb


def _make_xlsxwriter_format(
        workbook,
        fmt: FormatDict,
) -> Any:
    """
    Creates an xlsxwriter Format based on a logical FormatDict.
    Supports:
    - align, valign, text_wrap, indent, shrink_to_fit, text_rotation
    - border, border_*, border_color, border_*_color
    - pattern, fg_color, bg_color
    - font_name, font_size, bold, italic, underline, strike, font_color
    - num_format
    - locked, hidden

    Color fields can be:
    - HEX: 'FFAACC', '#FFAACC', '80FFAACC'
    - CSS: 'steelblue', 'red', ...
    - (If you want, in the future, tuple (r,g,b), (r,g,b,a) will be supported; then we'll expand the validator)
    """
    props: dict[str, Any] = {}

    # ---------- Alignment ----------
    if "align" in fmt:
        props["align"] = map_style_value("align", fmt["align"], "xlsxwriter")

    if "valign" in fmt:
        props["valign"] = map_style_value("valign", fmt["valign"], "xlsxwriter")

    if "text_wrap" in fmt and fmt["text_wrap"] is not None:
        props["text_wrap"] = bool(fmt["text_wrap"])

    if "indent" in fmt and fmt["indent"] is not None:
        props["indent"] = int(fmt["indent"])

    if "shrink_to_fit" in fmt and fmt["shrink_to_fit"] is not None:
        props["shrink"] = bool(fmt["shrink_to_fit"])

    if "text_rotation" in fmt and fmt["text_rotation"] is not None:
        props["rotation"] = int(fmt["text_rotation"])

    # ---------- Borders ----------
    if "border" in fmt and fmt["border"] is not None:
        props["border"] = map_style_value("border", fmt["border"], "xlsxwriter")

    if "border_color" in fmt and fmt["border_color"] is not None:
        norm = _normalize_color_xlsxwriter(fmt["border_color"])
        if norm is not None:
            props["border_color"] = norm

    side_keys = [
        ("border_left", "left"),
        ("border_right", "right"),
        ("border_top", "top"),
        ("border_bottom", "bottom"),
    ]
    for logical_key, xlsx_key in side_keys:
        if logical_key in fmt and fmt[logical_key] is not None:
            props[xlsx_key] = map_style_value("border", fmt[logical_key], "xlsxwriter")

    side_color_keys = [
        ("border_left_color", "left_color"),
        ("border_right_color", "right_color"),
        ("border_top_color", "top_color"),
        ("border_bottom_color", "bottom_color"),
    ]
    for logical_key, xlsx_key in side_color_keys:
        if logical_key in fmt and fmt[logical_key] is not None:
            norm = _normalize_color_xlsxwriter(fmt[logical_key])
            if norm is not None:
                props[xlsx_key] = norm

    # ---------- Fill ----------
    if "pattern" in fmt:
        mapped_pattern = map_style_value("pattern", fmt["pattern"], "xlsxwriter")
        if mapped_pattern is not None:
            props["pattern"] = mapped_pattern

    if "fg_color" in fmt and fmt["fg_color"] is not None:
        norm = _normalize_color_xlsxwriter(fmt["fg_color"])
        if norm is not None:
            props["fg_color"] = norm

    if "bg_color" in fmt and fmt["bg_color"] is not None:
        norm = _normalize_color_xlsxwriter(fmt["bg_color"])
        if norm is not None:
            props["bg_color"] = norm

    # ---------- Font ----------
    if "bold" in fmt and fmt["bold"] is not None:
        props["bold"] = bool(fmt["bold"])

    if "italic" in fmt and fmt["italic"] is not None:
        props["italic"] = bool(fmt["italic"])

    if "underline" in fmt and fmt["underline"] is not None:
        val = fmt["underline"]
        mapped = map_style_value("underline", val, "xlsxwriter")
        props["underline"] = mapped

    if "strike" in fmt and fmt["strike"] is not None:
        props["font_strikeout"] = bool(fmt["strike"])

    if "font_color" in fmt and fmt["font_color"] is not None:
        norm = _normalize_color_xlsxwriter(fmt["font_color"])
        if norm is not None:
            props["font_color"] = norm

    if "font_name" in fmt and fmt["font_name"] is not None:
        props["font_name"] = fmt["font_name"]

    if "font_size" in fmt and fmt["font_size"] is not None:
        props["font_size"] = fmt["font_size"]

    # ---------- Number format ----------
    if "num_format" in fmt and fmt["num_format"] is not None:
        props["num_format"] = fmt["num_format"]

    # ---------- Protection ----------
    if "locked" in fmt and fmt["locked"] is not None:
        props["locked"] = bool(fmt["locked"])

    if "hidden" in fmt and fmt["hidden"] is not None:
        props["hidden"] = bool(fmt["hidden"])

    return workbook.add_format(props)
