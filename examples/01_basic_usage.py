from pathlib import Path

from sheetkit import build_formatter_from_theme, get_theme, write_sheet

EXAMPLE_DIR = Path(__file__).resolve().parent


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
    file_name=EXAMPLE_DIR / "01_basic_usage.xlsx",
    sheet_name="Invoice",
    formatter=formatter,
    engine="xlsxwriter",
    mode="new",
    header=1,
)

print(path.resolve())

