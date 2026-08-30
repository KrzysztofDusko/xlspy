"""Update existing XLSX/XLSB templates.

Set the paths below to workbooks which already exist.  The updater preserves
the workbook structure and updates pivot-cache source metadata when present.
"""

from xlspy import XlsbUpdater, XlsxUpdater


ROWS = [
    ["Alice", 42],
    ["Bob", 37],
]
HEADERS = ["Name", "Amount"]


def update_xlsx(input_path: str, output_path: str) -> None:
    updater = XlsxUpdater(input_path)
    updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    updater.save(output_path)


def update_xlsb(input_path: str, output_path: str) -> None:
    updater = XlsbUpdater(input_path)
    updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    updater.save(output_path)


if __name__ == "__main__":
    update_xlsx("template.xlsx", "updated.xlsx")
    update_xlsb("template.xlsb", "updated.xlsb")
