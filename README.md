# Python XLSB/XLSX/XLSM Reader & Writer

A Python library for reading and writing XLSB and XLSX files efficiently, with
existing-workbook updates for macro-enabled XLSM packages.

## Installation

```bash
pip install xlspy
```

## Usage

### Basic Example

```python
from xlspy import XlsbWriter
import datetime
from decimal import Decimal

data = [
    ["Name", "Age", "City", "info"],
    [-123, 2147483647, 2147483648, 2147483999],
    ["x", "y", "z", datetime.datetime.today()],
    ["Alice", 25, "New York", datetime.date.today()],
    ["Bob", 30, "London", Decimal(3.14)],
    ["Charlie", 35, "Paris", datetime.datetime.now()],
    [True, False, None, datetime.datetime.utcnow()]
]

# Initialize writer with a specific compression level
with XlsbWriter("output.xlsb", compressionLevel=6) as writer:
    # Add a visible sheet
    writer.add_sheet("Visible Sheet")
    writer.write_sheet(data)

    # Add a hidden sheet
    writer.add_sheet("Hidden Sheet", hidden=True)
    writer.write_sheet([["This sheet is hidden."]])
```

### XlsxWriter Example

```python
from xlspy import XlsxWriter
import datetime
from decimal import Decimal

data = [
    ["Name", "Age", "City", "info"],
    [-123, 2147483647, 2147483648, 2147483999],
    ["x", "y", "z", datetime.datetime.today()],
    ["Alice", 25, "New York", datetime.date.today()],
    ["Bob", 30, "London", Decimal(3.14)],
    ["Charlie", 35, "Paris", datetime.datetime.now()],
    [True, False, None, datetime.datetime.utcnow()]
]

# Initialize writer with a specific compression level
with XlsxWriter("output.xlsx", compressionLevel=6) as writer:
    # Add a visible sheet
    writer.add_sheet("Visible Sheet")
    writer.write_sheet(data)

    # Add a hidden sheet
    writer.add_sheet("Hidden Sheet", hidden=True)
    writer.write_sheet([["This sheet is hidden."]])
```

### Cell Formatting (Number, Date, DateTime)

`xlspy` supports custom cell formatting for numbers, dates, and datetimes in both XLSB and XLSX output. Pass a `(value, format_string)` tuple to apply a format to a specific cell.

```python
from xlspy import XlsxWriter, F  # or XlsbWriter
import datetime

with XlsxWriter("formatted.xlsx") as writer:
    writer.add_sheet("Formats")
    writer.write_sheet([
        ["Description", "Value"],
        ["Thousands separator", (100000, F.THOUSANDS_SEP)],
        ["Currency PLN",       (100000, F.CURRENCY_PLN)],
        ["Currency EUR",       (100000, F.CURRENCY_EUR)],
        ["Percentage",         (100000, F.PERCENTAGE)],
        ["Scientific",         (100000, F.SCIENTIFIC)],
        ["Two decimals",       (100000, F.TWO_DECIMALS)],
        ["Text",               (100000, F.TEXT)],
        ["Leading zeros",      (100000, F.LEADING_ZEROS)],
        ["Short date",         (datetime.date(2026,6,1), F.DATE_SHORT)],
        ["Long date",          (datetime.date(2026,6,1), F.DATE_LONG)],
        ["ISO date",           (datetime.date(2026,6,1), F.DATE_ISO)],
        ["Month + year",       (datetime.date(2026,6,1), F.DATE_MONTH_YEAR)],
        ["Weekday + date",     (datetime.date(2026,6,1), F.DATE_WEEKDAY)],
        ["Short datetime",     (datetime.datetime(2026,6,1,14,34), F.DATETIME_SHORT)],
        ["Time only",          (datetime.datetime(2026,6,1,14,34), F.TIME_HH_MM)],
        ["12h time",           (datetime.datetime(2026,6,1,14,34), F.TIME_12H)],
        ["ISO datetime",       (datetime.datetime(2026,6,1,14,34), F.DATETIME_ISO)],
    ])
```

#### Available Format Constants (`xlspy.F`)

| Number | Date | DateTime |
|--------|------|----------|
| `F.THOUSANDS_SEP` — `#,##0` | `F.DATE_SHORT` — `dd.mm.yyyy` | `F.DATETIME_SHORT` — `dd.mm.yyyy hh:mm` |
| `F.CURRENCY_PLN` — `#,##0.00 "zł"` | `F.DATE_LONG` — `d mmmm yyyy` | `F.DATETIME_LONG` — `d mmmm yyyy hh:mm:ss` |
| `F.CURRENCY_EUR` — `#,##0.00 €` | `F.DATE_DAY_MONTH_YEAR` — `dd-mm-yyyy` | `F.TIME_HH_MM` — `hh:mm` |
| `F.PERCENTAGE` — `0%` | `F.DATE_ISO` — `yyyy-mm-dd` | `F.TIME_HH_MM_SS` — `hh:mm:ss` |
| `F.SCIENTIFIC` — `0.00E+00` | `F.DATE_MONTH_YEAR` — `mmmm yyyy` | `F.TIME_12H` — `h:mm AM/PM` |
| `F.TWO_DECIMALS` — `#,##0.00` | `F.DATE_WEEKDAY` — `dddd, d mmmm yyyy` | `F.DATETIME_24H` — `dd.mm.yyyy hh:mm:ss` |
| `F.TEXT` — `@` | `F.DATE_DAY_MONTH` — `d mmmm` | `F.DATETIME_ISO` — `yyyy-mm-dd"T"hh:mm:ss` |
| `F.LEADING_ZEROS` — `000000000` | `F.DATE_YEAR_ONLY` — `yyyy` | `F.TIME_MS` — `hh:mm:ss.000` |

You can also use custom format strings directly:

```python
writer.write_sheet([
    ["Custom", (1234.56, '#,##0.00 "USD"')],
    ["Date",   (datetime.date(2026,6,1), 'dd.mm.yyyy')],
])
```

The formatting works transparently on both `XlsxWriter` and `XlsbWriter`.

### Updating an Existing Workbook

`XlsxUpdater` (also for macro-enabled `.xlsm` files) and `XlsbUpdater` replace
the data region of an existing workbook
while keeping the workbook structure, styles, drawings and unrelated sheets.
They are useful when a workbook contains pivot tables: the worksheet source
range and pivot-cache refresh metadata are updated together.

```python
from xlspy import XlsxUpdater, XlsbUpdater

rows = [
    ["Alice", 42],
    ["Bob", 37],
]

updater = XlsxUpdater("template.xlsx")
print(updater.get_sheet_names())
updater.replace_sheet_data("Data", rows, headers=["Name", "Amount"])
updater.save("result.xlsx")       # use save() to replace the input atomically

# The XLSB API is identical:
binary_updater = XlsbUpdater("template.xlsb")
binary_updater.replace_sheet_data("Data", rows, headers=["Name", "Amount"])
binary_updater.save("result.xlsb")

# XLSM uses the same XML updater.  The VBA project is preserved byte-for-byte.
macro_updater = XlsxUpdater("template.xlsm")
macro_updater.replace_sheet_data("Data", rows, headers=["Name", "Amount"])
macro_updater.save("result.xlsm")
```

`XlsmUpdater` is an explicit alias of `XlsxUpdater` for `.xlsm` inputs.  The
updater does not parse, execute or rewrite VBA; `xl/vbaProject.bin` and other
unrelated macro parts are copied as opaque ZIP members.

For a database cursor or another one-pass source, use the streaming variant.
Rows are written through bounded-memory temporary files, so the complete
result set is not kept in Python memory. Multiple sheets may be streamed on
the same updater before `save()`:

```python
def rows_from_cursor(cursor):
    while True:
        row = cursor.fetchone()
        if row is None:
            break
        yield list(row)

updater = XlsxUpdater("template.xlsx")  # also accepts .xlsm; use XlsbUpdater for .xlsb
updater.replace_sheet_data_stream(
    "Data",
    rows_from_cursor(cursor),
    headers=["Name", "Amount"],
)
updater.save("result.xlsx")              # or result.xlsb
```

The streaming API requires `save()` for its memory-bounded write path.
`to_bytes()` necessarily creates the complete archive in memory because it
returns the complete file as `bytes`. The original `replace_sheet_data()` API
is unchanged and remains available for in-memory inputs.

Rows may be any iterable, including a generator. Only trailing rows consisting
entirely of `None` are removed; blank rows in the middle are preserved. By
default, the updater inherits the dominant style of each existing column. Use
`style_fallback="general"` to write new cells with the General style. Existing
writer-style `(value, format_string)` tuples are accepted for compatibility,
but updating a workbook does not create new styles.

When a pivot cache is present, keep the source column schema (column order and
count) unchanged; changing the schema requires rebuilding the pivot cache and
is outside the updater's scope.

On Windows, the generated workbook can be checked with the installed Excel
COM server (the script opens files read-only using Excel's normal loader):

```powershell
powershell -ExecutionPolicy Bypass -File tools\validate_excel_com.ps1 result.xlsx result.xlsb
```

### Reading XLSB and XLSX Files

Reading files is done via the `ExcelReader` class, which automatically detects the format.

```python
from xlspy import ExcelReader

with ExcelReader("input.xlsx") as reader:  # or .xlsb
    names = reader.get_sheet_names()
    print(f"Sheets: {names}")

    for sheet_name in names:
        rows = reader.read_all(sheet_name)
        for row in rows:
            print(row)

# Generator usage (memory efficient for large files):
with ExcelReader("large_file.xlsb") as reader:
    for row in reader.get_rows("Sheet1"):
        print(row)
```

### Streaming from a Database (Netezza)

This example shows how to stream data directly from a database query into an XLSB file using `nzpy-extended`. This is highly memory-efficient as it doesn't load the entire dataset into memory.

First, ensure you have `nzpy-extended` installed:
```bash
pip install nzpy-extended
```

Then, you can use a generator function to feed data to `XlsbWriter`.

```python
import os
from typing import Generator
from xlspy import XlsbWriter

# --- Configuration ---
NZ_CONFIG = {
    "host": os.environ.get("NZ_DEV_HOST", "your_host"),
    "port": int(os.environ.get("NZ_DEV_PORT", "5480")),
    "database": os.environ.get("NZ_DEV_DB", "your_db"),
    "user": os.environ.get("NZ_DEV_USER", "your_user"),
    "password": os.environ.get("NZ_DEV_PASSWORD", "your_password"),
}
QUERY = "SELECT * FROM YourTable"
OUTPUT_FILENAME = "db_output.xlsb"


def row_generator(cursor) -> Generator[list, None, None]:
    """Yields column headers first, then each data row."""
    headers = [column[0] for column in cursor.description]
    yield headers
    while row := cursor.fetchone():
        yield list(row)


# --- Main Execution ---
try:
    import nzpy_extended.sync as nzpy

    with nzpy.connect(**NZ_CONFIG) as conn:
        cursor = conn.cursor()
        cursor.execute(QUERY)

        with XlsbWriter(OUTPUT_FILENAME) as writer:
            writer.add_sheet("Database Export")
            writer.write_sheet(row_generator(cursor))
            writer.add_sheet("SQL Query", hidden=True)
            writer.write_sheet([["SQL"], [QUERY]])

    print(f"Successfully created '{OUTPUT_FILENAME}'")

except Exception as e:
    print(f"An unexpected error occurred: {e}")
```

## Performance

`xlspy` is designed for high performance. Since version 0.1.0, the library includes a **C extension** (`_c_core`) that accelerates XLSB read and write. The C extension is **enabled by default** (compiled automatically on install). Set `XLSPY_DISABLE_C_EXT=1` to force the pure Python fallback.

All benchmarks: **50000 × 50** dataset (2.5M cells). Tests performed on **Windows 11** (Python 3.14, AMD64).

### Write

| Library | Format | Time | Size |
|---------|--------|------|------|
| **xlspy (C_EXT)** | XLSB | **1.02 s** | 7.25 MB |
| xlspy (Python) | XLSB | 2.54 s | 7.25 MB |
| xlspy | XLSX | 5.35 s | 6.34 MB |
| [xlsxwriter](https://pypi.org/project/xlsxwriter/) | XLSX | 9.80 s | 11.57 MB |

### Read

| Library | Format | Time | Notes |
|---------|--------|------|-------|
| **xlspy (C_EXT)** | XLSB | **1.39 s** | default, compiled C |
| xlspy | XLSX | 4.72 s | uses expat XML parser (C) |
| xlspy (Python) | XLSB | 6.41 s | pure Python fallback |
| [openpyxl](https://pypi.org/project/openpyxl/) | XLSX | 7.85 s | read-only mode |


### Analysis

The **4.6× read speedup** comes from two factors:
- **~60–70%** — native C compilation, no interpreter overhead per record
- **~30–40%** — algorithm simplification: flat array indexed by `col − first_col` instead of `Dict[int, Any]`, no `isinstance` per cell, no `BiffReader.read_worksheet()` method call per record

Run the benchmarks yourself with `examples/performance_test.py`.

## Repository

<https://github.com/KrzysztofDusko/xlspy/>
