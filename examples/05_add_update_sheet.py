from pathlib import Path

from sheetkit import build_formatter_from_theme, get_theme, write_sheet

EXAMPLE_DIR = Path(__file__).resolve().parent

path = EXAMPLE_DIR / "05_add_update_sheet.xlsx"

formatter = build_formatter_from_theme(
    get_theme("office_theme"),
    types=["@", "#,##0.00"],
    header=1,
    zebra=True,
)

write_sheet(
    data_table=[
        ["Product", "Price"],
        ["Initial row", 10.0],
        ["Will be removed", 20.0],
    ],
    file_name=path,
    sheet_name="PriceList",
    formatter=formatter,
    engine="openpyxl",
    mode="new",
    header=1,
)

write_sheet(
    data_table=[
        ["Product", "Price"],
        ["Updated row", 99.9],
    ],
    file_name=path,
    sheet_name="PriceList",
    formatter=formatter,
    engine="openpyxl",
    mode="update",
    header=1,
    backup=True,
)

print(path.resolve())

