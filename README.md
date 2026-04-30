# sheetkit

`sheetkit` is a compact Python library for everyday Excel automation:
format tables, write whole sheets, patch only selected cells, and read sheets or ranges
back as plain `list[list[Any]]`.

It is for the practical middle ground between "raw `openpyxl` styling is too verbose"
and "I still need real `.xlsx` files, existing workbook updates, formulas, templates,
themes, and controlled formatting."

Use it when you need to:

- turn a Python table into a styled Excel worksheet quickly;
- read a whole sheet or a rectangular range (`"A1:B2"` or `((0, 0), (1, 1))`);
- update an existing workbook without rebuilding everything;
- patch only specific cells while preserving formulas, styles, validation, layout, and protection;
- reuse Office-like themes and formatter presets across reports.

## Quick Example

```python
from sheetkit import read_sheet, write_sheet

write_sheet(
    data_table=[
        ["Name", "Qty", "Price"],
        ["Apples", 10, 1.25],
        ["Bananas", 7, 0.90],
    ],
    file_name="report.xlsx",
    sheet_name="Sales",
    formatter={"row": {0: {"bold": True}}, "col": {}},
)

values = read_sheet("report.xlsx", sheet_name="Sales", sheet_range="A1:B2")
```

## Custom Formatter In One Dictionary

You can format a worksheet without presets by passing a plain dictionary:

- `row`: styles by zero-based row index.
- `col`: styles by zero-based column index.
- `0`, `1`, ...: explicit row/column styles.
- `-1`: default style.
- `-2`: zebra style.
- `priority`: `"row"` or `"col"` decides which side wins when both define the same cell.

```python
from sheetkit import write_sheet

formatter = {
    "priority": "row",
    "row": {
        0: {
            "bold": True,
            "align": "center",
            "pattern": "solid",
            "fg_color": "4472C4",
            "font_color": "FFFFFF",
        },
        -1: {"border_bottom": 1},
        -2: {"pattern": "solid", "fg_color": "F2F6FC"},
    },
    "col": {
        1: {"num_format": "#,##0.00", "align": "right"},
    },
}

write_sheet(
    data_table=[["Item", "Amount"], ["Apples", 1234.5], ["Bananas", 900]],
    file_name="custom.xlsx",
    formatter=formatter,
    header=1,
)
```

## What it does

- writes `.xlsx` files from `list[list[Any]]`
- reads full worksheets and ranges into `list[list[Any]]`
- can save `.xltx` Excel templates through `openpyxl`
- supports both `xlsxwriter` and `openpyxl`
- uses a logical, engine-independent formatting model
- supports row formats, column formats, zebra rows, headers, number formats, borders, fonts, fills, and alignment
- supports Excel / Office themes
- can import `.thmx` theme files
- can extract formatter JSON from a manually styled Excel worksheet
- can replace, update, or patch sheets in existing workbooks through `openpyxl`

## Terms

This README uses three key terms:
- `theme`: a color/font theme (typically from `.thmx` or `presets/themes` JSON).
- `formatter`: runtime `SheetFormatSpec` used by `write_sheet(...)`.
- `formatter JSON`: serialized formatter file that stores `row`/`col` specs.

## Why use it

Working with Excel formatting directly in `openpyxl` is powerful, but verbose and often awkward. `sheetkit` keeps the useful part of that power while making the common workflow much simpler:

1. describe table formatting with plain Python dictionaries
2. choose an engine
3. save the workbook

## Formatting Layers

`sheetkit` formatting works through three layers. Keeping this model in mind makes the API much easier to navigate.

### Layer 1: Theme Presets

Theme presets define color/font foundations.

Typical sources:
- built-in JSON theme presets in `sheetkit/presets/themes`
- imported `.thmx` themes (`export_excel_theme`, `export_excel_themes`, `import_theme`, `import_themes`)
- runtime theme resolution (`get_theme`)

### Layer 2: Formatter JSON

Typical sources:
- extracted from styled worksheet (`extract_formatter_from_sheet`, `extract_formatter_to_file`)
- saved/loaded as JSON (`save_formatter`, `load_formatter`)

### Layer 3: Runtime Formatter

Runtime formatter (`SheetFormatSpec`) is what actually styles a specific worksheet during save.

Typical builders:
- `build_formatter_from_theme(...)` for direct theme->runtime conversion (`variant=0` base, `variant=1..N` from theme variants)

Then apply it with:
- `write_sheet(...)`

### Recommended path

1. Start with theme (`get_theme`) or a formatter JSON (`load_formatter` / extraction).
2. Build runtime formatter (`build_formatter_from_theme`).
3. Save output (`write_sheet`).

### Goal -> Function(s) Cheat Sheet

| Goal | Use |
| --- | --- |
| Save a table with a ready formatter | `write_sheet` |
| Read a sheet/range as `list[list[Any]]` | `read_sheet` |
| Build runtime formatter from theme data | `build_formatter_from_theme` |
| Load theme by name/path | `get_theme` |
| Import one `.thmx` theme into runtime cache | `import_theme` |
| Import all `.thmx` themes from folder into runtime cache | `import_themes` |
| Load formatter JSON | `load_formatter` |
| Load built-in formatter preset by name | `load_format_preset` |
| Save formatter JSON | `save_formatter` |
| List built-in theme preset names | `list_preset_themes` |
| List built-in formatter preset names | `list_preset_formatters` |
| Extract formatter from styled worksheet (in-memory) | `extract_formatter_from_sheet` |
| Extract formatter from styled worksheet to JSON | `extract_formatter_to_file` |
| Extract formatter from a fully styled worksheet range | `extract_formatter_range_from_sheet` |
| Read `.thmx` into memory | `load_thmx_theme` |
| Convert `.thmx` to JSON preset | `export_excel_theme` |
| Normalize/resolve color values | `color_to_hex`, `normalize_color_value` |

## Quick Start

```python
from pathlib import Path

from sheetkit import build_formatter_from_theme, get_theme, write_sheet

data = [
    ["Name", "Price", "Discount", "Total"],
    ["Apples", 10.5, 0.05, 9.98],
    ["Bananas", 12.0, 0.10, 10.80],
    ["Oranges", 12.1, 0.00, 12.10],
]

formatter = build_formatter_from_theme(
    get_theme("office_theme"),
    types=["@", "#,##0.00", "0.00%", "#,##0.00"],
    header=1,
    zebra=True,
)

path = write_sheet(
    data_table=data,
    file_name=Path("fruits.xlsx"),
    sheet_name="Invoice",
    formatter=formatter,
    header=1,
)

print(path)
```

```python
from sheetkit import read_sheet

all_values = read_sheet("fruits.xlsx")
part_values = read_sheet("fruits.xlsx", sheet_range="A1:B2")
part_values2 = read_sheet("fruits.xlsx", sheet_range=((0, 0), (1, 1)))
```

## Core API

### `write_sheet(...) -> Path`

Main entry point for saving a formatted workbook.
`save_formatted_xlsx(...)` is kept as a backward-compatible alias.

```python
write_sheet(
    data_table: list[list[Any]],
    file_name: str | PathLike,
    sheet_name: str | None = None,
    formatter: SheetFormatSpec | None = None,
    mode: ModeLiteral = "auto",
    engine: EngineLiteral = "auto",
    header: int | None = None,
    backup: bool = True,
    patch_skip_none: bool = True,
    patch_skip_formulas: bool = True,
    patch_skip_locked: bool = True,
    patch_range: str | tuple[tuple[int, int], tuple[int, int]] | None = None,
    patch_strict_range: bool = False,
    as_template: bool = False,
    progressor: PercentProgress | None = None,
) -> Path
```

It returns the saved file path and raises exceptions on invalid input or save errors.

`header` behavior:
- `None` (default): do not touch formatter indexes `0/1`; keep formatter header mapping as-is.
- `0`, `1`, `2`: explicit runtime override for header depth on the priority axis.

## Public API Reference

All symbols exported from [`sheetkit/__init__.py`](https://github.com/Prevalex/sheetkit/blob/main/sheetkit/__init__.py) are documented in:

- [API reference](https://github.com/Prevalex/sheetkit/blob/main/docs/api.md)

This includes:

- every exported function
- parameter-by-parameter descriptions
- exported type aliases and literals

Quick export list:

- `write_sheet`
- `read_sheet`
- `build_formatter_from_theme`
- `resolve_formatter_colors`
- `get_theme`
- `import_theme`, `import_themes`
- `load_formatter`, `save_formatter`
- `load_format_preset`
- `list_preset_themes`, `list_preset_formatters`
- `load_preset_file`, `save_preset_file`
- `extract_formatter_from_sheet`, `extract_formatter_to_file`
- `extract_formatter_range_from_sheet`, `extract_formatter_range_to_file`
- `export_excel_theme`
- `export_excel_themes`
- `load_thmx_theme`
- `color_to_hex`
- `normalize_color_value`
- `FormatDict`, `AxisFormatSpec`, `SheetFormatSpec`, `ColorInput`, `EngineLiteral`, `ModeLiteral`

### `build_formatter_from_theme(..., color_mode="hex" | "ref") -> SheetFormatSpec`

`color_mode="ref"` stores colors as theme references (`Accent1`, `Accent1+40`, `Dark2`, etc.) for readable JSON.
Use `resolve_formatter_colors(...)` to convert such formatter to pure HEX colors when needed for performance.

### Theme import helpers

- `load_thmx_theme(...)`
- `import_theme(...)`
- `import_themes(...)`
- `get_theme(...)`

### Formatter extraction helpers

- `extract_formatter_from_sheet(...)`
- `extract_formatter_to_file(...)`
- `save_formatter(...)`

## Formatting Model

`sheetkit` uses a logical formatting model that is independent of the output engine.

### `FormatDict`

A dictionary describing a single cell style.

Example:

```python
{
    "bold": True,
    "align": "center",
    "pattern": "solid",
    "fg_color": "4472C4",
    "font_color": "FFFFFF",
    "border_bottom": 1,
}
```

### `AxisFormatSpec`

A dictionary keyed by row or column index:

```python
{
    -1: {"border_bottom": 1},
    -2: {"pattern": "solid", "fg_color": "F7F7F7"},
    0: {"bold": True, "align": "center"},
}
```

Special keys:

- `-1`: default format
- `-2`: second zebra format

### `SheetFormatSpec`

A sheet formatter contains both row and column format maps:

```python
{
    "priority": "row",  # or "col"
    "row": {...},
    "col": {...},
}
```

`priority` controls which axis wins when row and column formats define the same property:

- `"row"`: row format overrides column format; `header` means top header rows; `types` applies by columns.
- `"col"`: column format overrides row format; `header` means left header columns; `types` applies by rows.

If `priority` is missing, `sheetkit` treats the formatter as `"row"` for compatibility with existing formatter JSON.

## Supported Formatting Keys

Supported logical style keys include:

- `align`, `valign`, `text_wrap`, `indent`, `shrink_to_fit`, `text_rotation`
- `border`, `border_left`, `border_right`, `border_top`, `border_bottom`
- `border_color`, `border_left_color`, `border_right_color`, `border_top_color`, `border_bottom_color`
- `pattern`, `fg_color`, `bg_color`
- `font_name`, `font_size`, `bold`, `italic`, `underline`, `strike`, `font_color`
- `num_format`
- `locked`, `hidden`

### `FormatDict` Reference

| Property | Type | Allowed values | Nullable |
| --- | --- | --- | --- |
| `align` | `str` | `"fill"`, `"left"`, `"justify"`, `"center"`, `"right"` | yes |
| `valign` | `str` | `"bottom"`, `"justify"`, `"distributed"`, `"center"`, `"top"` | yes |
| `text_wrap` | `bool` | | yes |
| `indent` | `int` | | yes |
| `shrink_to_fit` | `bool` | | yes |
| `text_rotation` | `int` | | yes |
| `border` | `int \| str` | `0`, `1`, `2`, `"thin"`, `"medium"` | yes |
| `border_color` | `str \| tuple` | see color input formats | yes |
| `border_left` | `int \| str` | `0`, `1`, `2`, `"thin"`, `"medium"` | yes |
| `border_right` | `int \| str` | `0`, `1`, `2`, `"thin"`, `"medium"` | yes |
| `border_top` | `int \| str` | `0`, `1`, `2`, `"thin"`, `"medium"` | yes |
| `border_bottom` | `int \| str` | `0`, `1`, `2`, `"thin"`, `"medium"` | yes |
| `border_left_color` | `str \| tuple` | see color input formats | yes |
| `border_right_color` | `str \| tuple` | see color input formats | yes |
| `border_top_color` | `str \| tuple` | see color input formats | yes |
| `border_bottom_color` | `str \| tuple` | see color input formats | yes |
| `pattern` | `str` | `"solid"` | yes |
| `fg_color` | `str \| tuple` | see color input formats | yes |
| `bg_color` | `str \| tuple` | see color input formats | yes |
| `font_name` | `str` | | yes |
| `font_size` | `int \| float` | | yes |
| `bold` | `bool` | | yes |
| `italic` | `bool` | | yes |
| `underline` | `bool \| str` | `False`, `True`, `"single"`, `"double"`, `"singleAccounting"`, `"doubleAccounting"` | yes |
| `strike` | `bool` | | yes |
| `font_color` | `str \| tuple` | see color input formats | yes |
| `num_format` | `str` | Excel number format string | yes |
| `locked` | `bool` | | yes |
| `hidden` | `bool` | | yes |

## Color Input

Colors can be given in several forms.

### Plain colors

- HEX: `"FFCC99"`, `"#FFCC99"`
- RGB: `(255, 204, 153)`
- RGBA: `(255, 204, 153, 255)`
- CSS names: `"steelblue"`, `"tomato"`, `"lightgray"`

### Theme color references

`sheetkit` also supports Excel-style theme color references:

- `("wisp", "accent3")`
- `"wisp:accent3"`
- `"Accent1+40"`
- `"Background2-20"`
- `"Text1"`
- `"Hyperlink"`

If no theme is explicitly specified in the reference string, `office_theme` is used by default.

## Predefined Themes

Built-in preset theme names currently included in the package:

`banded`, `basis`, `berlin`, `circuit`, `damask`, `dividend`, `droplet`, `facet`, `frame`, `gallery`, `integral`, `ion`, `ion_boardroom`, `main_event`, `mesh`, `metropolitan`, `office_theme`, `organic`, `parallax`, `parcel`, `retrospect`, `savon`, `slate`, `slice`, `tete`, `vapor_trail`, `view`, `wisp`, `wood_type`.

Built-in formatter preset names currently included in the package (each has both `_row` and `_col` variants):

`banded`, `basis`, `berlin`, `circuit`, `damask`, `dividend`, `droplet`, `facet`, `frame`, `gallery`, `integral`, `ion`, `ion_boardroom`, `main_event`, `mesh`, `metropolitan`, `office_theme`, `organic`, `parallax`, `parcel`, `retrospect`, `savon`, `slate`, `slice`, `tete`, `vapor_trail`, `view`, `wisp`, `wood_type`.

Quick access helpers:
- `list_preset_themes()`
- `list_preset_formatters()`
- `load_format_preset(name, priority="row" | "col")`

### Theme resolution order (`get_theme`)

`get_theme(theme, auto_import=True)` resolves in this order:

1. explicit file path (`.json` or `.thmx`)
2. built-in preset file in `sheetkit/presets/themes`
3. user Office theme folder (`%APPDATA%\Microsoft\Templates\Document Themes`)
4. Office installation theme folder (`...\Document Themes 16`)
5. runtime cache (`THEMES`)

### Formatter preset loading

`load_format_preset(name, priority)` is intentionally strict:

- it loads only from `sheetkit/presets/formatters`
- it resolves `<name>_row.json` or `<name>_col.json`
- it does not read external/user paths

## Extracting Formatters from Excel

`sheetkit` can also extract formatter from a worksheet that you styled manually in Excel.

This is useful when:

- you want to design the look visually in Excel first
- you want to reuse that look later from Python
- you do not want to write formatter JSON by hand

### Recommended worksheet layout

For row-priority formatting, prepare a small sample block that starts at a known cell, usually `A1`:

1. header row 1
2. header row 2
3. first data row style
4. second zebra row style

- row and header styles are read from the first column of the sample block
- column formats currently import only `num_format`
- `sheet_name` is optional; if omitted, the active worksheet is used
- `columns` is required and defines how many sample columns to scan

For column-priority formatting, use the same idea rotated 90 degrees:

1. header column 1
2. header column 2
3. first data column style
4. second zebra column style

- column and header styles are read from the first row of the sample block
- row formats currently import only `num_format`
- `rows` is required and defines how many sample rows to scan

### Example

```python
from pathlib import Path

from sheetkit import extract_formatter_to_file

extract_formatter_to_file(
    file_name=Path("styled_template.xlsx"),
    json_file=Path("imported_blue.json"),
    columns=6,
    start_cell="A1",
    header=2,
    zebra=True,
    priority="row",
    name="imported_blue",
)
```

The resulting JSON file can then be loaded with `sheetkit.themes.load_formatter(...)`.

### Extracting full row/column ranges

When every row or every column has its own style, use `extract_formatter_range_from_sheet(...)`.
It copies styles from a rectangular range rather than interpreting the range as header/base/zebra samples.

- `priority="row"`: row styles are read from the first column; column `num_format` values are scanned down each column.
- `priority="col"`: column styles are read from the first row; row `num_format` values are scanned across each row.
- the last row/column style is stored as the default style `-1`.

## CLI Tools

Two small command-line wrappers are available in [scripts](https://github.com/Prevalex/sheetkit/tree/main/scripts).

### `extract_formatter`

Wraps `extract_formatter_to_file(...)` and `extract_formatter_range_to_file(...)`.

```bash
python scripts/extract_formatter.py <file_name> <json_file> [-m sample] -p row -cs <columns> [-sn <sheet_name>] [-sc <start_cell>] [-hr <header>] [-z | --no-zebra] [-name <name>]
python scripts/extract_formatter.py <file_name> <json_file> [-m sample] -p col -rs <rows> [-sn <sheet_name>] [-sc <start_cell>] [-hr <header>] [-z | --no-zebra] [-name <name>]
python scripts/extract_formatter.py <file_name> <json_file> -m range -p row -rs <rows> -cs <columns> [-sn <sheet_name>] [-sc <start_cell>] [-name <name>]
python scripts/extract_formatter.py <file_name> <json_file> -m range -p col -rs <rows> -cs <columns> [-sn <sheet_name>] [-sc <start_cell>] [-name <name>]
```

Example:

```bash
python scripts/extract_formatter.py styled_template.xlsx imported_blue.json -cs 6 -sn StyledTable -hr 2 -z -name imported_blue
python scripts/extract_formatter.py scripts/Columns.xlsx scripts/__temp__/columns_range.json -m range -p col -rs 12 -cs 7 -name columns_range
```

### `format_xlsx`

Loads a formatter JSON and applies it via `write_sheet(...)`.

Supported input formats:

- `.csv`
- `.xlsx`
- `.xls` if optional dependency `xlrd` is installed

```bash
python scripts/format_xlsx.py <input_file> <output_file> -ft <formatter_json> [-hr 0|1|2] [--zebra | --no-zebra]
```

Example:

```bash
python scripts/format_xlsx.py data.csv formatted.xlsx -ft my_formatter.json -hr 1 -z
```

`format_xlsx` override behavior:
- if `-hr` is omitted, formatter header indexes are kept as-is;
- if `--zebra/--no-zebra` is omitted, formatter zebra mapping is kept as-is.

### `export_excel_themes`

Exports Office `.thmx` themes to theme JSON. With `-cm`, it also builds self-contained formatter JSON files
from those themes and saves them into `<save-folder>/formatters` as both
`<theme>_row.json` and `<theme>_col.json`.

```bash
python scripts/export_excel_themes.py -of -uf -sf sheetkit/presets/themes
python scripts/export_excel_themes.py -of -sf exported_themes -cm ref
```

## Engines and Modes

### Engines

- `"auto"`: default; uses `xlsxwriter` for new workbooks and `openpyxl` for existing workbooks, `replace`/`update`/`patch` modes, and templates
- `"xlsxwriter"`: best choice for creating new files quickly
- `"openpyxl"`: required for updating existing workbooks and for `.xltx`

### Modes

- `"new"`: create a new workbook file; raises `FileExistsError` if the target file already exists
- `"replace"`: replace a sheet in an existing workbook
- `"update"`: update an existing sheet in an existing workbook
- `"patch"`: patch existing sheet cells only (safe merge mode: skips `None`, formulas, and locked cells on protected sheets by default)
- `"auto"`: choose mode automatically based on file existence and engine

Notes:

- `replace`, `update`, and `patch` are supported only with `openpyxl`
- `as_template=True` is supported only with `openpyxl`

### Patch Mode Behavior

`mode="patch"` is designed for calculation templates where formulas/styles/protection already exist on the sheet.

- The update anchor is `A1`: `data_table[r][c]` maps to Excel cell `(r+1, c+1)`.
- `data_table` size can be smaller or larger than the existing used range.
- Smaller patch: only provided cells are considered; other cells are unchanged.
- Larger patch: new rows/cells can be created below/right of the old range.
- With `patch_skip_none=True` (default), `None` means "skip cell", not "clear cell".
- With `patch_skip_formulas=True` (default), formula cells are not overwritten.
- With `patch_skip_locked=True` (default), locked cells on protected sheets are not overwritten.
- Optional `patch_range="D2:E7"` changes the patch anchor to `D2`.
- You can also pass zero-based tuple bounds:
  `patch_range=((0, 0), (1, 1))` is equivalent to `A1:B2`.
- With `mode="auto"` and an existing workbook, `patch_range` selects patch behavior
  instead of replacing the sheet.
- With `patch_strict_range=False` (default), out-of-range patch values are silently ignored.
- With `patch_strict_range=True`, sheetkit raises an error if data does not fit into `patch_range`.

Example:

```python
write_sheet(
    data_table=[
        [None],      # do not touch A1
        [123],       # write A2
        [456],       # write A3
    ],
    file_name="report.xlsx",
    sheet_name="Calc",
    formatter={"row": {}, "col": {}},
    engine="openpyxl",
    mode="patch",
)
```

```python
write_sheet(
    data_table=[[11], [21], [32]],
    file_name="report.xlsx",
    sheet_name="Calc",
    formatter={"row": {}, "col": {}},
    engine="openpyxl",
    mode="patch",
    patch_range="D2:D7",  # writes D2, D3, D4
)
```

```python
write_sheet(
    data_table=[[11, 12], [21, 22]],
    file_name="report.xlsx",
    sheet_name="Calc",
    formatter={"row": {}, "col": {}},
    engine="openpyxl",
    mode="patch",
    patch_range=((0, 0), (1, 1)),  # writes A1:B2
)
```

## Merge Rules

When both row and column styles define the same property, formatter `priority` decides which axis wins:

- `priority="row"`: row formatting has priority over column formatting.
- `priority="col"`: column formatting has priority over row formatting.

This is important when combining:

- default column formats
- explicit column formats
- default row formats
- explicit row formats

## Examples

- [01_basic_usage.py](https://github.com/Prevalex/sheetkit/blob/main/examples/01_basic_usage.py) — minimal workbook creation
- [02_head_and_num_columns.py](https://github.com/Prevalex/sheetkit/blob/main/examples/02_head_and_num_columns.py) — headers, zebra rows, number formats
- [03_styled_columns.py](https://github.com/Prevalex/sheetkit/blob/main/examples/03_styled_columns.py) — manual formatter and theme colors in columns
- [04_styled_rows.py](https://github.com/Prevalex/sheetkit/blob/main/examples/04_styled_rows.py) — manual formatter and row-specific styling
- [05_add_update_sheet.py](https://github.com/Prevalex/sheetkit/blob/main/examples/05_add_update_sheet.py) — updating an existing sheet with `openpyxl`
- [06_save_as_template.py](https://github.com/Prevalex/sheetkit/blob/main/examples/06_save_as_template.py) — saving `.xltx` templates
- [07_extract_formatter_from_sheet.py](https://github.com/Prevalex/sheetkit/blob/main/examples/07_extract_formatter_from_sheet.py) — extracting formatter from a styled Excel worksheet
- [08_read_sheet.py](https://github.com/Prevalex/sheetkit/blob/main/examples/08_read_sheet.py) — reading full sheet and ranges (`"A1:B2"` / `((0,0),(1,1))`)

## Current Status

The project is already usable and tested for its core workflow:

- type checking with `mypy` has been cleaned up in the codebase
- automated tests cover formatter logic, themes, color references, and workbook save/update flows

At the same time, the library is still evolving, and we keep refining the API surface and documentation.

## Development

Current checks used in this repository:

```bash
python -m mypy sheetkit --config-file mypy.ini
python -m pytest -q
```

Release flow (local + TestPyPI + PyPI) is documented in [docs/release.md](https://github.com/Prevalex/sheetkit/blob/main/docs/release.md).

## License

MIT




