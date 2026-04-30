# sheetkit Public API

`sheetkit` is a small Excel automation API for Python code that needs to:

- write styled tables to `.xlsx`;
- read worksheets or ranges as plain `list[list[Any]]`;
- patch selected cells in existing workbooks without clearing formulas, styles, validation, or layout;
- reuse compact formatter dictionaries, Office-like themes, and formatter presets.

This page documents symbols exported by [`sheetkit/__init__.py`](https://github.com/Prevalex/sheetkit/blob/main/sheetkit/__init__.py).

## Fast Path

```python
from sheetkit import read_sheet, write_sheet

write_sheet(
    data_table=[["Name", "Qty"], ["Apples", 10]],
    file_name="report.xlsx",
    sheet_name="Sales",
    formatter={"row": {0: {"bold": True}}, "col": {}},
)

values = read_sheet("report.xlsx", sheet_name="Sales", sheet_range=((0, 0), (1, 1)))
```

For existing calculation workbooks, use `mode="patch"` or pass `patch_range` to update
only selected cells while preserving the rest of the sheet.

## Custom Formatter Shape

Custom formatting is just a dictionary. The core rules are:

- `row`: zero-based row styles.
- `col`: zero-based column styles.
- `-1`: default row/column style.
- `-2`: zebra row/column style.
- `priority`: `"row"` or `"col"` controls which style wins when both apply.

```python
formatter = {
    "priority": "row",
    "row": {
        0: {"bold": True, "pattern": "solid", "fg_color": "4472C4", "font_color": "FFFFFF"},
        -1: {"border_bottom": 1},
        -2: {"pattern": "solid", "fg_color": "F2F6FC"},
    },
    "col": {
        1: {"num_format": "#,##0.00", "align": "right"},
    },
}
```

## Terms

- `theme`: Office-like color/font theme dictionary (from `.thmx` or `presets/themes/*.json`).
- `formatter`: runtime `SheetFormatSpec` (`{"priority": "row"|"col", "row": AxisFormatSpec, "col": AxisFormatSpec}`).
- `formatter JSON`: persisted formatter file with `kind="formatter"` and `row`/`col` mappings.

## Exported Functions

### `write_sheet(...) -> Path`

Main write function. Applies formatter to data and saves `.xlsx` (or `.xltx` when `as_template=True`).
`save_formatted_xlsx(...)` is available as a backward-compatible alias.

Key options:
- `engine`: `"auto"` by default; chooses `xlsxwriter` for new files and `openpyxl` for existing files, `replace`/`update`/`patch`, and templates
- `mode`: `"auto"` by default; creates a new file when it does not exist and replaces the target sheet when it does
- `mode="new"`: safe-create mode; raises `FileExistsError` if target file already exists
- `mode="patch"`: sparse update mode for existing sheets; updates only provided cells and by default skips `None`, formulas, and locked cells on protected sheets
- `header`: `None` by default; keeps formatter header indexes `0/1` unchanged
- `header=0..2`: explicit runtime header override; means top header rows in row-priority mode and left header columns in col-priority mode
- `backup`: when replacing/updating through `openpyxl`, keep a versioned copy of the old sheet
- `as_template`: save `.xltx`; uses `openpyxl`

Patch-specific options:
- `patch_skip_none=True`: in `mode="patch"`, `None` does not overwrite target cell.
- `patch_skip_formulas=True`: in `mode="patch"`, existing formula cells are not overwritten.
- `patch_skip_locked=True`: in `mode="patch"`, locked cells are not overwritten when sheet protection is enabled.
- `patch_range: str | tuple[tuple[int, int], tuple[int, int]] | None = None`:
  optional patch range as Excel string (`"D2:E7"`) or zero-based tuple bounds
  `((start_row, start_col), (end_row, end_col))`.
  Patch is anchored at range top-left.
- With `mode="auto"` and an existing file, passing `patch_range` selects patch behavior instead of replacing the sheet.
- `patch_strict_range=False`: when `True`, raises if `data_table` exceeds `patch_range` bounds.

Mode semantics summary:
- `new`: creates workbook/sheet from scratch.
- `replace`: replaces target sheet content by creating a new sheet version (optionally keeping a backup sheet).
- `update`: clears target sheet range, then writes provided data + formatter.
- `patch`: does not clear the sheet; updates only addressed cells and preserves existing layout/style/formulas by default.

Patch range behavior:
- Patch address space is always anchored at `A1`: `data_table[r][c] -> cell(r+1, c+1)`.
- If `patch_range` is set, patch is anchored at its top-left cell.
- Patch table may be smaller than existing sheet data: only provided coordinates are considered.
- Patch table may be larger than existing used range: new rows/cells are created as needed.
- With `patch_range` and default `patch_strict_range=False`, values outside the range are ignored.
- There is no explicit cell-address map in patch mode by design.

Patch examples:

```python
from sheetkit import write_sheet

# Update only first column (A), do not touch other columns.
write_sheet(
    data_table=[[None], [100], [200], [300]],
    file_name="book.xlsx",
    sheet_name="Calc",
    formatter={"row": {}, "col": {}},
    engine="openpyxl",
    mode="patch",
)
```

```python
# Patch inside explicit range (anchor D2). This writes D2, D3, D4.
write_sheet(
    data_table=[[11], [21], [32]],
    file_name="book.xlsx",
    sheet_name="Calc",
    formatter={"row": {}, "col": {}},
    engine="openpyxl",
    mode="patch",
    patch_range="D2:D7",
)
```

```python
# Patch using zero-based tuple range: ((0,0),(1,1)) -> A1:B2.
write_sheet(
    data_table=[[11, 12], [21, 22]],
    file_name="book.xlsx",
    sheet_name="Calc",
    formatter={"row": {}, "col": {}},
    engine="openpyxl",
    mode="patch",
    patch_range=((0, 0), (1, 1)),
)
```

```python
# Force-write None (clear cells) and allow overwriting formulas.
write_sheet(
    data_table=[[None, None], [None, 0]],
    file_name="book.xlsx",
    sheet_name="Calc",
    formatter={"row": {}, "col": {}},
    engine="openpyxl",
    mode="patch",
    patch_skip_none=False,
    patch_skip_formulas=False,
)
```

### `read_sheet(...) -> list[list[Any]]`

Reads worksheet values without styles.

Key options:
- `sheet_name`: worksheet name; if omitted, active sheet is used.
- `sheet_range`: optional range as Excel string (`"A1:B2"`) or zero-based tuple
  `((start_row, start_col), (end_row, end_col))`.
- `data_only=True`: forwarded to `openpyxl.load_workbook(data_only=...)`.

Examples:

```python
from sheetkit import read_sheet

all_values = read_sheet("book.xlsx")
part_values = read_sheet("book.xlsx", sheet_range="A1:B2")
part_values2 = read_sheet("book.xlsx", sheet_range=((0, 0), (1, 1)))
```

### `build_formatter_from_theme(...) -> SheetFormatSpec`

Builds formatter from theme colors/fonts.

Key options:
- `header`: `0..2`; means top header rows in row-priority mode and left header columns in col-priority mode
- `priority`: `"row"` or `"col"`; missing priority is treated as `"row"`
- `zebra`: include `row[-2]` stripe in row-priority mode or `col[-2]` stripe in col-priority mode
- `variant`: `0` = base, `1..N` = theme variants when present
- `accent`: preferred accent (`1..6`, fallback to accent1)
- `color_mode`: `"hex"` or `"ref"` (`Accent1`, `Accent1+40`, `Dark2`, ...)
- `types`: column `num_format` list in row-priority mode; row `num_format` list in col-priority mode

### `resolve_formatter_colors(formatter, *, theme=None) -> SheetFormatSpec`

Converts color references in formatter (`Accent1`, `theme:accent3`, CSS names, etc.) into HEX.

### `get_theme(theme, *, auto_import=True) -> dict[str, Any]`

Resolves theme by:

1. explicit file path (`.json` or `.thmx`)
2. built-in `presets/themes/<name>.json`
3. user themes (`%APPDATA%\\Microsoft\\Templates\\Document Themes`)
4. Office themes (`...\\Document Themes 16`)
5. runtime cache (`THEMES`)

Built-in preset theme names bundled with the package:

`banded`, `basis`, `berlin`, `circuit`, `damask`, `dividend`, `droplet`, `facet`, `frame`, `gallery`, `integral`, `ion`, `ion_boardroom`, `main_event`, `mesh`, `metropolitan`, `office_theme`, `organic`, `parallax`, `parcel`, `retrospect`, `savon`, `slate`, `slice`, `tete`, `vapor_trail`, `view`, `wisp`, `wood_type`.

See also: `list_preset_themes()`.

### `import_theme(...)` / `import_themes(...)`

Imports one or many `.thmx` themes and (optionally) registers them in runtime cache.

### `load_thmx_theme(file_name) -> dict[str, Any]`

Parses `.thmx` and returns normalized theme dictionary:

- `kind="themes"`
- `source` (file name)
- optional `application`, `version`
- `scheme_name`, `colors`, `fonts`, optional `variants`, `color_transforms`

### `extract_formatter_from_sheet(...) -> SheetFormatSpec`

Reads a styled worksheet sample block and extracts formatter:

- row-priority: row styles from first sample column, column `num_format` from sample columns; requires `columns`
- col-priority: column styles from first sample row, row `num_format` from sample rows; requires `rows`
- supports `header`, `header_rows` alias, `zebra`, `start_cell`, `sheet_name`

### `extract_formatter_to_file(...) -> Path`

Wrapper over `extract_formatter_from_sheet(...)` + save formatter JSON.

### `extract_formatter_range_from_sheet(...) -> SheetFormatSpec`

Reads a fully styled rectangular range and extracts formatter:

- `priority="row"`: row styles from the first range column, column `num_format` by scanning columns
- `priority="col"`: column styles from the first range row, row `num_format` by scanning rows
- requires explicit `rows` and `columns`
- stores the last row/column style as default `-1`

### `extract_formatter_range_to_file(...) -> Path`

Wrapper over `extract_formatter_range_from_sheet(...)` + save formatter JSON.

### `scripts/extract_formatter.py` CLI

Supports two extraction modes:

- `-m sample` (default): template extraction with `header`/`zebra` logic
- `-m range`: full range extraction; requires both `rows` and `columns`

In `range` mode, header/zebra flags are ignored because the range is copied directly.

### `scripts/format_xlsx.py` CLI

Formatter application supports optional overrides:

- `-hr/--header`: optional header override (`0..2`)
- `--zebra` / `--no-zebra`: optional zebra override

If these flags are omitted, formatter mappings are kept unchanged.

## Cookbook (CLI)

1. Sample extraction (row priority):
```bash
python scripts/extract_formatter.py styled.xlsx styled_row.json -m sample -p row -cs 8 -hr 2 -z
```

2. Sample extraction (col priority):
```bash
python scripts/extract_formatter.py styled.xlsx styled_col.json -m sample -p col -rs 10 -hr 2 -z
```

3. Full range extraction (row/col priorities):
```bash
python scripts/extract_formatter.py scripts/rows.xlsx scripts/__temp__/rows_range.json -m range -p row -rs 7 -cs 7
python scripts/extract_formatter.py scripts/Columns.xlsx scripts/__temp__/columns_range.json -m range -p col -rs 12 -cs 7
```

4. Apply formatter JSON to CSV:
```bash
python scripts/format_xlsx.py scripts/rows.csv scripts/__temp__/rows_result.xlsx -ft scripts/__temp__/rows_range.json
python scripts/format_xlsx.py scripts/Columns.csv scripts/__temp__/columns_result.xlsx -ft scripts/__temp__/columns_range.json
```

5. Export themes and auto-build formatters:
```bash
python scripts/export_excel_themes.py -of -uf -sf scripts/__temp__/themes -cm ref
```

With `-cm`, formatter files are created as `<theme>_row.json` and `<theme>_col.json`.

### `load_formatter(source) -> SheetFormatSpec`

Loads formatter JSON and normalizes row/col keys to `int`.

### `load_format_preset(name, priority="row"|"col") -> SheetFormatSpec`

Loads formatter preset strictly from `sheetkit/presets/formatters`:

- resolves `<name>_row.json` or `<name>_col.json` based on `priority`
- does not read user-provided external paths
- useful for "out-of-the-box" formatter usage

Built-in formatter preset names bundled with the package (each has both row/col variants):

`banded`, `basis`, `berlin`, `circuit`, `damask`, `dividend`, `droplet`, `facet`, `frame`, `gallery`, `integral`, `ion`, `ion_boardroom`, `main_event`, `mesh`, `metropolitan`, `office_theme`, `organic`, `parallax`, `parcel`, `retrospect`, `savon`, `slate`, `slice`, `tete`, `vapor_trail`, `view`, `wisp`, `wood_type`.

See also: `list_preset_formatters()`.

### `list_preset_themes() -> list[str]`

Returns sorted names of built-in preset themes from `sheetkit/presets/themes`.

### `list_preset_formatters() -> list[str]`

Returns sorted names for which both formatter variants exist:

- `<name>_row.json`
- `<name>_col.json`

### `save_formatter(formatter, json_file, *, name=None, theme=None) -> Path`

Saves formatter JSON with:

- `kind: "formatter"`
- `priority`
- `row`, `col`
- optional `name`
- optional embedded `theme`, useful for self-contained formatter JSON with theme color references

### `load_preset_file(source, *, expected_kind=None) -> dict[str, Any]`

Generic JSON loader for preset files.

### `save_preset_file(preset, json_file) -> Path`

Generic JSON writer for preset files.

### `export_excel_theme(...) -> Path`

Exports one `.thmx` theme to JSON.

### `export_excel_themes(...) -> dict[str, Path]`

Exports all `.thmx` themes from selected directories.

### `color_to_hex(...) -> str` / `normalize_color_value(...) -> str | None`

Color normalization utilities (HEX/CSS/RGB/theme references).

## Exported Types

- `FormatDict`
- `AxisFormatSpec`
- `SheetFormatSpec`
- `ColorInput`
- `EngineLiteral`
- `FormatPriorityLiteral`
- `ModeLiteral`

