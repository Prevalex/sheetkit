from pathlib import Path

from sheetkit import write_sheet

EXAMPLE_DIR = Path(__file__).resolve().parent


data = [
    ["Type", "Description", "Amount"],
    ["income", "Salary", 4200],
    ["income", "Bonus", 700],
    ["expense", "Rent", -1500],
    ["expense", "Food", -480],
]

formatter = {
    "row": {
        0: {
            "bold": True,
            "align": "center",
            "pattern": "solid",
            "fg_color": "office_theme:accent1",
            "font_color": "Background1",
        },
        -1: {"border_bottom": 1, "border_bottom_color": "Background2-25"},
        1: {"fg_color": "Accent5+70", "pattern": "solid"},
        2: {"fg_color": "Accent5+55", "pattern": "solid"},
        3: {"fg_color": "Accent2+70", "pattern": "solid"},
        4: {"fg_color": "Accent2+55", "pattern": "solid"},
    },
    "col": {
        2: {"num_format": "#,##0;[Red]-#,##0", "align": "right"},
    },
}

path = write_sheet(
    data_table=data,
    file_name=EXAMPLE_DIR / "04_styled_rows.xlsx",
    sheet_name="Cashflow",
    formatter=formatter,
    engine="openpyxl",
    mode="new",
    header=1,
)

print(path.resolve())

