from pathlib import Path

from sheetkit import build_formatter_from_theme, get_theme, write_sheet

EXAMPLE_DIR = Path(__file__).resolve().parent


data = [
    ["Category", "Product", "Price", "Margin", "Share"],
    ["Laptops", "Model A", 25999.9, 0.18, 0.1245],
    ["Laptops", "Model B", 31999.0, 0.22, 0.0831],
    ["Phones", "Model C", 15999.0, 0.16, 0.1010],
]

formatter = build_formatter_from_theme(
    get_theme("office_theme"),
    types=["@", "@", "#,##0.00", "0.00%", "0.00%"],
    header=1,
    zebra=True,
)

path = write_sheet(
    data_table=data,
    file_name=EXAMPLE_DIR / "02_head_and_num_columns.xlsx",
    sheet_name="PriceList",
    formatter=formatter,
    engine="xlsxwriter",
    mode="new",
    header=1,
)

print(path.resolve())

