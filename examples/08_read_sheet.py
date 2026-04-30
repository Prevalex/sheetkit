from pathlib import Path

from sheetkit import read_sheet, write_sheet

EXAMPLE_DIR = Path(__file__).resolve().parent
FILE_PATH = EXAMPLE_DIR / "08_read_sheet.xlsx"

write_sheet(
    data_table=[
        ["Name", "Qty", "Price"],
        ["Apples", 10, 1.25],
        ["Bananas", 7, 0.9],
    ],
    file_name=FILE_PATH,
    sheet_name="Data",
    formatter={"row": {}, "col": {}},
    engine="openpyxl",
    mode="new",
)

all_values = read_sheet(FILE_PATH, sheet_name="Data")
range_values_excel = read_sheet(FILE_PATH, sheet_name="Data", sheet_range="A1:B2")
range_values_tuple = read_sheet(FILE_PATH, sheet_name="Data", sheet_range=((0, 0), (1, 1)))

print("File:", FILE_PATH.resolve())
print("All:")
print(all_values)
print("A1:B2:")
print(range_values_excel)
print("((0,0),(1,1)):")
print(range_values_tuple)
