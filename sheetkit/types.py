from typing import Any, Literal, TypeAlias

#
################# [CUSTOM DATA TYPES SECTION] ###############
#

# FormatDict:
# ----------
# Dictionary of cell formatting properties {format property: property value, ...},
# for example: {"align":"left", "border_right":0}
FormatDict: TypeAlias = dict[str, Any]

# AxisFormatSpec:
# --------------
# Row or column format dictionary. This is a dictionary where the keys are the indices of the rows or columns for
# which the format is specified. Used to assign a given logical format to the row or column with the given number.
# The -1 key is used to specify the format for those rows or columns for which the format is not explicitly specified.
# The -2 key is used to draw a vertical or horizontal zebra stripe. If specified, even rows/columns
# for which the format is not explicitly specified will be formatted using this format (and odd rows/columns,
# respectively, will be formatted using the format with the -1 key)
AxisFormatSpec: TypeAlias = dict[int, FormatDict]  # ключи: -2, -1, 0, 1, 2, ...

# SheetFormatSpec:
# ----------------
# A sheet format dictionary consisting of row and column format dictionaries. This is a dictionary of logical
# format dictionaries
# of rows and columns of an AxisFormatSpec. It has the following format:
# {'col': AxisFormatSpec, 'row':AxisFormatSpec, 'priority': 'row'|'col'}
SheetFormatSpec: TypeAlias = dict[str, Any]

# RowFormats:
# ------------
# A row-based dictionary of effective formats, in which each row index corresponds to a list of formats
# for each of the row's columns. Each column format in this list is a dictionary obtained by merging
# the row format dictionary with the column format dictionary, with the row format dictionary taking precedence:
# CellFormatDict[i, j] = ColumnAxisFormatSpec[j] | RowAxisFormatSpec[i]
# A logical RowFormats is translated into a library RowFormats by translating property names and their values
# into the names and values used in the specified library (openpyxl/xlsxwriter)
# Row indexes of a row format can have the values -2, -1, 0, 1, 2, ..., where -1, -2 have the same meaning as
# in AxisFormatSpec
RowFormats: TypeAlias = dict[int, list[FormatDict]]  # {row_idx: [fmt_dict0, fmt_dict1, ...]}

# EngineLiteral: library (engine) name
EngineLiteral = Literal["auto", "xlsxwriter", "openpyxl"]
ResolvedEngineLiteral = Literal["xlsxwriter", "openpyxl"]
FormatPriorityLiteral = Literal["row", "col"]

ModeLiteral = Literal["auto", "new", "replace", "update", "patch"]
AxisFormatKey = Literal["row", "col"]

# Logical formatting: dictionary {property: value}
# A logical cell format is an engine-independent format. It is then translated into a format specific to
# openpyxl or xlsxwriter. The property names and values of both engines sometimes coincide and match the
# properties and values of the logical format.
LogicalSpecEntry: TypeAlias = dict[str, Any]

ColorInput: TypeAlias = (
        str
        | tuple[int, int, int]
        | tuple[int, int, int, int]
        | tuple[str, str]
)

COL_FORMAT_KEY: AxisFormatKey = "col"  # columnar AxisFormatSpec key in SheetFormatSpec
ROW_FORMAT_KEY: AxisFormatKey = "row"  # lowercase AxisFormatSpec key in SheetFormatSpec
