"""Update existing XLSX/XLSM/XLSB templates.

Set the paths below to workbooks which already exist.  The updater preserves
the workbook structure and updates pivot-cache source metadata when present.
"""

from xlspy import XlsbUpdater, XlsmUpdater, XlsxUpdater


ROWS = [
    ["Alice", 42],
    ["Bob", 37],
]
HEADERS = ["Name", "Amount"]


def row_generator():
    """Replace this generator with a database cursor adapter in production."""

    yield from ROWS


def rows_from_cursor(cursor):
    """Yield database rows one at a time without building a result list."""

    while True:
        row = cursor.fetchone()
        if row is None:
            break
        yield list(row)


def update_xlsx(input_path: str, output_path: str) -> None:
    updater = XlsxUpdater(input_path)
    updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    updater.save(output_path)


def update_xlsm(input_path: str, output_path: str) -> None:
    """Update XLSM data while preserving the opaque VBA project part."""

    updater = XlsmUpdater(input_path)
    updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    updater.save(output_path)


def update_xlsb(input_path: str, output_path: str) -> None:
    updater = XlsbUpdater(input_path)
    updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    updater.save(output_path)


def update_xlsx_stream(input_path: str, output_path: str) -> None:
    updater = XlsxUpdater(input_path)
    updater.replace_sheet_data_stream("Data", row_generator(), headers=HEADERS)
    updater.save(output_path)


def update_xlsm_stream(input_path: str, output_path: str) -> None:
    """Stream rows into XLSM without loading or changing its VBA project."""

    updater = XlsmUpdater(input_path)
    updater.replace_sheet_data_stream("Data", row_generator(), headers=HEADERS)
    updater.save(output_path)


def update_xlsb_stream(input_path: str, output_path: str) -> None:
    updater = XlsbUpdater(input_path)
    updater.replace_sheet_data_stream("Data", row_generator(), headers=HEADERS)
    updater.save(output_path)


if __name__ == "__main__":
    update_xlsx("template.xlsx", "updated.xlsx")
    update_xlsb("template.xlsb", "updated.xlsb")
