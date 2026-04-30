#!

from collections.abc import Mapping, Sequence
from typing import Any

from .helpers import inspect_name
from .types import FormatDict, ResolvedEngineLiteral

"""
STYLE_MAP is a mapping map between "our logical values" and "engine-specific values."
Therefore, this mapping only makes sense where actual mapping is needed, rather than simply passing the value
directly. For many properties (e.g., font_size, font_name), no mapping is needed—xlsxwriter and openpyxl
accept the value directly.

Therefore, we will only set STYLE_MAP mappings for properties that have:
"""
STYLE_MAP: dict[str, dict[ResolvedEngineLiteral, dict[Any, Any]]] = {
    "align": {
        "xlsxwriter": {
            "left": "left",
            "center": "center",
            "right": "right",
            "justify": "justify",
            "fill": "fill",
        },
        "openpyxl": {
            "left": "left",
            "center": "center",
            "right": "right",
            "justify": "justify",
            "fill": "fill",
        },
    },

    "valign": {
        "xlsxwriter": {
            "top": "top",
            "center": "vcenter",
            "bottom": "bottom",
        },
        "openpyxl": {
            "top": "top",
            "center": "center",
            "bottom": "bottom",
            "justify": "justify",
            "distributed": "distributed",
        },
    },

    # border:  1=thin, 2=medium
    "border": {
        "xlsxwriter": {
            1: 1,
            2: 2,
            "thin": 1,
            "medium": 2,
        },
        "openpyxl": {
            1: "thin",
            2: "medium",
            "thin": "thin",
            "medium": "medium",
        },
    },

    "pattern": {
        "xlsxwriter": {"solid": 1, None: None},
        "openpyxl": {"solid": "solid", None: None},
    },

    # text_wrap: for both engines it's just a bool → mapping isn't needed, but we'll leave it
    "text_wrap": {
        "xlsxwriter": {True: True, False: False},
        "openpyxl": {True: True, False: False},
    },

    # underline: different values
    # xlsxwriter: 1 or "single", "double"
    # openpyxl: "single", "double"
    "underline": {
        "xlsxwriter": {
            True: 1,  # xlsxwriter: 1 == underline single
            False: None,
            "single": 1,
            "double": 2,
        },
        "openpyxl": {
            True: "single",
            False: None,
            "single": "single",
            "double": "double",
        },
    },

    # strike: both engines have bool, mapping is not formally required
    "strike": {
        "xlsxwriter": {True: True, False: False},
        "openpyxl": {True: True, False: False},
    },

    # Colors are accepted as strings in both engines
    # (HEX or ARGB), mapping is not required, but we'll leave it as is
    "font_color": {
        "xlsxwriter": {},
        "openpyxl": {},
    },

    "fg_color": {
        "xlsxwriter": {},
        "openpyxl": {},
    },

    "bg_color": {
        "xlsxwriter": {},
        "openpyxl": {},
    },

    # font_name and font_size are passed directly
    "font_name": {
        "xlsxwriter": {},
        "openpyxl": {},
    },

    "font_size": {
        "xlsxwriter": {},
        "openpyxl": {},
    },

    # num_format: the number format is passed directly as a string
    "num_format": {
        "xlsxwriter": {},
        "openpyxl": {},
    },
}


def map_style_value(prop: str, value: Any, engine: ResolvedEngineLiteral) -> Any:
    """
    Translates a logical cell property into a property value for the xlsxwriter or openpyxl library, if their
    value encoding differs and therefore requires adaptation depending on which library
    is used.

    If STYLE_MAP[prop][engine] has nothing for the value, it will return the value directly.
    That is:
    return STYLE_MAP.get(prop, {}).get(engine, {}).get(value, value)

    If there is no mapping, it will return the original.

    This allows:
    for font_size → pass directly
    for num_format → pass directly
    for underline → perform real mapping
    for border → perform real mapping
    for unexplored properties → can be easily expanded in the future

    Parameters
    ----------
    prop - property name (logical format key)
    value - property value (value by logical format key) format)
    engine - the name of the library for which the translation is performed ('openpyxl', 'xlsxwriter')

    Returns
    -------
    the value of the specified property in a format understandable for the specified library
    """

    if value is None:
        return None
    prop_map = STYLE_MAP.get(prop)
    if prop_map is None:
        return value
    engine_map = prop_map.get(engine)
    if engine_map is None:
        return value
    return engine_map.get(value, value)


def _guess_align_from_num_format(num_fmt: str) -> str:
    """
    A very simple helper:
    - text formats (with '@') → 'left'
    - everything else → 'right'
    """
    if not num_fmt:
        return "left"
    if "@" in num_fmt:
        return "left"
    # You can make the heuristic more complex if you wish.
    return "right"


# clear dictionary of None
def _compact_fmt(d: Mapping[str, Any]) -> FormatDict:
    """Убирает ключи со значением None."""
    return {k: v for k, v in d.items() if v is not None}


# Calculating column widths for an Excel sheet
def compute_column_widths(
        data: Sequence[Sequence[Any]],
        *,
        header_rows: int = 0,
        use_header_in_width: bool = True,
        min_width: int = 5,
        max_width: int = 80,
        padding: int = 1,
        reserve_chars: int = 0,
) -> list[int]:
    """
    Calculates optimal column widths.

    Args:
    data: table (list[list[Any]])
    header_rows: number of header rows (0,1,2,...)
    use_header_in_width:
    True — headers are included in the width calculation,
    False — headers are ignored.
    min_width: minimum column width
    max_width: maximum width
    padding: right padding
    reserve_chars: extra safety width added to every column

    Returns:
    List[int] — column widths.
    """
    if not data:
        return []

    import unicodedata
    from datetime import date, datetime

    def _char_display_width(ch: str) -> int:
        """Returns the display width of the character (ASCII=1, CJK/emoji=2)."""
        return 2 if unicodedata.east_asian_width(ch) in ("F", "W") else 1

    def _text_display_width(text: str) -> int:
        """Calculates the displayed width of a string."""
        return sum(_char_display_width(ch) for ch in text)

    num_cols = max(len(row) for row in data)
    col_widths = [min_width] * num_cols

    for r_idx, row in enumerate(data):
        is_header = r_idx < header_rows

        # ignore headers if specified
        if is_header and not use_header_in_width:
            continue

        for c_idx in range(num_cols):
            if c_idx >= len(row):
                continue

            val = row[c_idx]

            # Convert the value to a string
            if val is None:
                s = ""
            elif isinstance(val, (int, float)):
                s = str(val)
            elif isinstance(val, (date, datetime)):
                s = val.isoformat()
            elif isinstance(val, bool):
                s = "TRUE" if val else "FALSE"
            else:
                s = str(val)

            w = _text_display_width(s) + padding
            col_widths[c_idx] = max(col_widths[c_idx], w)

    reserve = max(0, int(reserve_chars))
    col_widths = [max(min_width, min(w + reserve, max_width)) for w in col_widths]

    return col_widths


def clear_range_openpyxl(max_col: int, max_row: int, ws: Any) -> None:
    from openpyxl.worksheet.worksheet import Worksheet

    if not isinstance(ws, Worksheet):
        raise TypeError(f"{inspect_name()}: ws must be a openpyxl Worksheet")

    ws.merged_cells.ranges.clear()
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.style = "Normal"
            cell.number_format = "General"
