#!

from copy import deepcopy
from typing import Any, cast

from .colors import validate_key_color_value
from .fmt_openpyxl import _make_openpyxl_style
from .fmt_xlsxwriter import _make_xlsxwriter_format
from .types import AxisFormatSpec, COL_FORMAT_KEY, FormatDict, FormatPriorityLiteral, \
    LogicalSpecEntry, ROW_FORMAT_KEY, RowFormats, SheetFormatSpec
from .types import ResolvedEngineLiteral

#
################ [LOGICAL FORMATTING SECTION ] ################
#

""" COLOR_KEYS: List of STYLE_MAP keys that control color - used during logical format validation,
    as an indication that the value for these keys should be additionally validated via validate_key_color_value
    """
COLOR_KEYS = {
    "font_color",
    "fg_color",
    "bg_color",
    "border_color",
    "border_left_color",
    "border_right_color",
    "border_top_color",
    "border_bottom_color",
}

"""
We define a logical format—a list of formatting keys that will be used to describe the cell format.
The logical format is described by a dictionary, where the specified keys correspond to values. 
The dictionary specification is described by the LOGICAL_STYLE_SPEC dictionary, which contains 
additional information necessary for validating the logical dictionary.

How the logical dictionary is mapped to engines such as openpyxl or xlsxwriter (conceptually)
----------------------------------------------------------------------------------------
    In xlsxwriter
        align → the align parameter in workbook.add_format()
        valign → valign (with "center" mapping → "vcenter")
        text_wrap → text_wrap=True/False
        indent → indent
        shrink_to_fit → shrink
        text_rotation → rotation
        border / border_* → border, top, bottom, left, right
        border_color / *_color → corresponding *_color fields
        pattern, fg_color, bg_color → pattern, fg_color, bg_color
        font_* → bold, italic, underline, font_color, font_name, font_size, font_strikeout
        num_format → num_format
        locked / hidden → locked, hidden (xlsxwriter supports this via format)

    In openpyxl
        align, valign, text_wrap, indent, shrink_to_fit, text_rotation → Alignment(...) object
        border / border_*, *_color → Border(left=Side(style=..., color=...), ...)
        pattern, fg_color, bg_color → PatternFill(patternType=..., fgColor=..., bgColor=...)
        font_* → Font(name=..., size=..., bold=..., italic=..., underline=..., strike=..., color=...)
        num_format → cell.number_format = fmt["num_format"]
        locked, hidden → Protection(locked=..., hidden=...) и cell.protection = ...

How to use this in code
---------------------------

We'll fix the API: "The logical style will always be specified by these keys - see LOGICAL_STYLE_SPEC."
    In the translator functions, for each key, we either:
        pass it directly (if the format is the same),
        or transform it via STYLE_MAP and custom logic (underline, border, pattern, valign, etc.).

    Everything else—adding new presets, inheritance, combinations—is done at the logical dictionary level, without 
    touching the xlsxwriter/openpyxl code.
"""

LOGICAL_STYLE_SPEC: dict[str, LogicalSpecEntry] = {
    #
    # This dictionary is a representation of the LogicalFormat dictionary for convenient use in the
    # validate_format_dict validator
    #
    # --- Alignment ---
    "align": {
        "types": (str,),
        "choices": {"left", "center", "right", "justify", "fill"},
        "nullable": True,
    },
    "valign": {
        "types": (str,),
        "choices": {"top", "center", "bottom", "justify", "distributed"},
        "nullable": True,
    },
    "text_wrap": {
        "types": (bool,),
        "nullable": True,
    },
    "indent": {
        "types": (int,),
        "nullable": True,
    },
    "shrink_to_fit": {
        "types": (bool,),
        "nullable": True,
    },
    "text_rotation": {
        "types": (int,),
        "nullable": True,
    },

    # --- Borders ---
    "border": {
        "types": (int, str),
        "choices": {0, 1, 2, "thin", "medium"},
        "nullable": True,
    },
    "border_color": {
        "types": (str, tuple),
        "nullable": True,
    },
    "border_left": {
        "types": (int, str),
        "choices": {0, 1, 2, "thin", "medium"},
        "nullable": True,
    },
    "border_right": {
        "types": (int, str),
        "choices": {0, 1, 2, "thin", "medium"},
        "nullable": True,
    },
    "border_top": {
        "types": (int, str),
        "choices": {0, 1, 2, "thin", "medium"},
        "nullable": True,
    },
    "border_bottom": {
        "types": (int, str),
        "choices": {0, 1, 2, "thin", "medium"},
        "nullable": True,
    },
    "border_left_color": {
        "types": (str, tuple),
        "nullable": True,
    },
    "border_right_color": {
        "types": (str, tuple),
        "nullable": True,
    },
    "border_top_color": {
        "types": (str, tuple),
        "nullable": True,
    },
    "border_bottom_color": {
        "types": (str, tuple),
        "nullable": True,
    },

    # --- Fill ---
    "pattern": {
        "types": (str,),
        "choices": {"solid"},
        "nullable": True,
    },
    "fg_color": {
        "types": (str, tuple),
        "nullable": True,
    },
    "bg_color": {
        "types": (str, tuple),
        "nullable": True,
    },

    # --- Font ---
    "font_name": {
        "types": (str,),
        "nullable": True,
    },
    "font_size": {
        "types": (int, float),
        "nullable": True,
    },
    "bold": {
        "types": (bool,),
        "nullable": True,
    },
    "italic": {
        "types": (bool,),
        "nullable": True,
    },
    "underline": {
        "types": (bool, str),
        "choices": {True, False, "single", "double", "singleAccounting", "doubleAccounting"},
        "nullable": True,
    },
    "strike": {
        "types": (bool,),
        "nullable": True,
    },
    "font_color": {
        "types": (str, tuple),
        "nullable": True,
    },

    # --- Number format ---
    "num_format": {
        "types": (str,),
        "nullable": True,
    },

    # --- Protection ---
    "locked": {
        "types": (bool,),
        "nullable": True,
    },
    "hidden": {
        "types": (bool,),
        "nullable": True,
    },
}


def validate_format_dict(fmt: FormatDict, strict: bool = True) -> None:
    """
    Checks whether the logical format specified by the dictionary matches the specification specified in LOGICAL_STYLE_SPEC
    Args:
        fmt: Logical format dictionary
        strict: Sets the response to an unknown format key:
        If True, throws a ValueError if the dictionary does not match the specification.
        If False, simply ignores the key.

    Returns:
        None
    """
    for key, value in fmt.items():
        spec = LOGICAL_STYLE_SPEC.get(key)
        if spec is None:
            if strict:
                raise ValueError(f"Unknown format key: {key!r}")
            else:
                continue

        if value is None:
            if not spec.get("nullable", False):
                raise ValueError(f"Format key {key!r} does not accept None")
            continue

        allowed_types = spec.get("types")
        if allowed_types is not None and not isinstance(value, allowed_types):
            raise ValueError(
                f"Format key {key!r} expects {allowed_types}, got {type(value)}"
            )

        choices = spec.get("choices")
        if choices is not None and value not in choices:
            raise ValueError(
                f"Format key {key!r} has invalid value {value!r}, allowed: {choices}"
            )

        # --- additional check for color keys ---
        if key in COLOR_KEYS:
            validate_key_color_value(value, key)


def build_row_formats_with_columns(
        format_spec: SheetFormatSpec,
        max_cols: int,
        max_rows: int | None = None,
        inherit_defaults: bool = True,
) -> RowFormats:
    """
    Builds a dictionary of efficient logical row formats, taking into account column formats.

    Input:
        format_spec: {
        "col": { -1: {...}, 0: {...}, 1: {...}, ... },
        "row": { -1: {...}, 2: {...}, 5: {...}, ... }
        }

        max_cols: The number of columns for which formats are built (columns 0..max_cols-1).

        inherit_defaults:
        - True (default):
            The row/column inherits format -1 (row[-1], col[-1]).
        - False:
            either its own format, or format -1 if none exists.

    Exit:
        {
        -1: [merged_fmt_(-1,0), merged_fmt_(-1,1), ... ],
        2: [merged_fmt_(2,0), merged_fmt_(2,1), ... ],
        ...
        }

        * row priority: merged_fmt(row_idx, col_idx) = col_fmt(col_idx) | row_fmt(row_idx)
        * col priority: merged_fmt(row_idx, col_idx) = row_fmt(row_idx) | col_fmt(col_idx)
    """
    col_format_spec: AxisFormatSpec = format_spec.get(COL_FORMAT_KEY, {})
    row_format_spec: AxisFormatSpec = format_spec.get(ROW_FORMAT_KEY, {})
    priority = cast(FormatPriorityLiteral, format_spec.get("priority", "row"))
    if priority not in ("row", "col"):
        priority = "row"

    default_col_format: FormatDict = col_format_spec.get(-1, {})
    default_row_format: FormatDict = row_format_spec.get(-1, {})

    def axis_fmt(axis_spec: AxisFormatSpec, axis_idx: int, default_fmt: FormatDict, *, zebra: bool) -> FormatDict:
        """
        Calculates row/column format, including optional zebra template lookup.
        """
        if inherit_defaults:
            if axis_idx == -1:
                return default_fmt
            if axis_idx in axis_spec:
                return default_fmt | axis_spec[axis_idx]

            template_idx = -1
            if zebra and -2 in axis_spec:
                zebra_start = 0
                while zebra_start in axis_spec:
                    zebra_start += 1
                template_idx = -1 if ((axis_idx - zebra_start) % 2 == 0) else -2
            return default_fmt | axis_spec.get(template_idx, {})

        if axis_idx in axis_spec:
            return axis_spec[axis_idx]
        if zebra and -2 in axis_spec:
            zebra_start = 0
            while zebra_start in axis_spec:
                zebra_start += 1
            template_idx = -1 if ((axis_idx - zebra_start) % 2 == 0) else -2
            return axis_spec.get(template_idx, default_fmt)
        return default_fmt

    def col_fmt(col_idx: int) -> FormatDict:
        """
        Calculates the cell format based on the column format specified in the formatter.
        """
        return axis_fmt(
            col_format_spec,
            col_idx,
            default_col_format,
            zebra=priority == "col",
        )

    def row_fmt(row_idx: int) -> FormatDict:
        """
        Calculates the cell format based on the row format specified in the formatter.
        """
        return axis_fmt(
            row_format_spec,
            row_idx,
            default_row_format,
            zebra=priority == "row",
        )

    formatter: RowFormats = {}
    row_indexes = set(row_format_spec)

    # If only column formats are specified, we still need the default row template (-1),
    # otherwise, the col formats won't be included in the final formatter and won't be applied at all.

    if priority == "col" and max_rows is not None:
        row_indexes.update(range(max_rows))
    elif col_format_spec and -1 not in row_indexes:
        row_indexes.add(-1)

    for row_idx in sorted(row_indexes):
        row_fmt_list: list[FormatDict] = []
        base_row_fmt = row_fmt(row_idx)

        for col_idx in range(max_cols):
            base_col_fmt = col_fmt(col_idx)
            merged = base_row_fmt | base_col_fmt if priority == "col" else base_col_fmt | base_row_fmt
            row_fmt_list.append(deepcopy(merged))  # mutation protection
        formatter[row_idx] = row_fmt_list

    return formatter


################ [ PHYSICAL FORMATTING SECTION ] ##############

def translate_formatter(
        formatter: RowFormats,
        engine: ResolvedEngineLiteral,
        workbook: Any = None,
) -> dict[int, list[Any]]:
    """
    Converts the logical row-based formatter (row -> list[FormatDict]) created by build_row_formats_with_columns()
    to an engine-specific row-based formatter (row -> list[StyleObject]).

    engine:
    "xlsxwriter" — returns Format objects;
    "openpyxl" — returns OpenpyxlCellStyle.

    (!) xlsxwriter requires a workbook (an xlsxwriter.Workbook object) because the Format is created
    via workbook.add_formaft().

    IMPORTANT:
    - Before creating a style, each unique logical format is validated using the validate_format_dict(...) function.
    - If the format does not comply with the public logical API, a ValueError with a user-friendly description
      will be raised.
    """
    normalized_engine = cast(ResolvedEngineLiteral, engine.lower())
    if normalized_engine not in ("xlsxwriter", "openpyxl"):
        raise ValueError(f"Unsupported engine: {engine}")
    engine = normalized_engine

    if engine == "xlsxwriter" and workbook is None:
        raise ValueError("For engine='xlsxwriter' you must provide 'workbook'.")

    engine_formatter: dict[int, list[Any]] = {}

    # Cache: one logical format -> one style object
    style_cache: dict[tuple[tuple[str, Any], ...], Any] = {}

    for row_idx, fmt_list in formatter.items():
        row_styles: list[Any] = []

        for fmt in fmt_list:
            # Cache key: sorted (name, value) pairs
            # (assuming the values are hashable: numbers/strings, etc.)
            key = tuple(sorted(fmt.items()))

            if key in style_cache:
                style_obj = style_cache[key]
            else:
                # --- VALIDATION of logical format ---
                try:
                    validate_format_dict(fmt, strict=True)
                except ValueError as e:
                    raise ValueError(f'{row_idx=}; {fmt=}: {e!r}') from e

                # --- Creating a style ---
                if engine == "xlsxwriter":
                    style_obj = _make_xlsxwriter_format(workbook, fmt)
                else:  # openpyxl
                    style_obj = _make_openpyxl_style(fmt)

                style_cache[key] = style_obj

            row_styles.append(style_obj)

        engine_formatter[row_idx] = row_styles

    return engine_formatter
