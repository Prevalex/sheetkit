from pathlib import Path

from sheetkit import write_sheet

EXAMPLE_DIR = Path(__file__).resolve().parent


data = [
    ["Region", "Manager", "Revenue", "Plan", "Status"],
    ["North", "Alice", 120000, 0.97, "On track"],
    ["South", "Bob", 98000, 0.84, "At risk"],
    ["West", "Carol", 143500, 1.08, "Ahead"],
]

formatter = {
    "row": {
        0: {
            "bold": True,
            "align": "center",
            "valign": "center",
            "text_wrap": True,
            "pattern": "solid",
            "fg_color": "office_theme:accent1",
            "font_color": "Text1",
            "border_bottom": 1,
            "border_bottom_color": "Accent1-25",
        },
        -1: {"border_bottom": 1, "border_bottom_color": "Background2-20"},
        -2: {
            "pattern": "solid",
            "fg_color": "Background2",
            "border_bottom": 1,
            "border_bottom_color": "Background2-20",
        },
    },
    "col": {
        0: {"fg_color": ("wisp", "accent6"), "pattern": "solid"},
        2: {"num_format": "#,##0", "align": "right"},
        3: {"num_format": "0.00%", "align": "right"},
        4: {"fg_color": "Accent2+55", "pattern": "solid"},
    },
}

path = write_sheet(
    data_table=data,
    file_name=EXAMPLE_DIR / "03_styled_columns.xlsx",
    sheet_name="Dashboard",
    formatter=formatter,
    engine="xlsxwriter",
    mode="new",
    header=1,
)

print(path.resolve())

