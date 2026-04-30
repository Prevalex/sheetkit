from pathlib import Path

from sheetkit import build_formatter_from_theme, get_theme, write_sheet

EXAMPLE_DIR = Path(__file__).resolve().parent


data = [
    ["Customer", "Due date", "Amount"],
    ["", "", ""],
]

formatter = build_formatter_from_theme(
    get_theme("office_theme"),
    types=["@", "@", "#,##0.00"],
    header=1,
    zebra=False,
)

path = write_sheet(
    data_table=data,
    file_name=EXAMPLE_DIR / "06_save_as_template",
    sheet_name="InvoiceTemplate",
    formatter=formatter,
    engine="openpyxl",
    mode="new",
    header=1,
    as_template=True,
)

print(path.resolve())

