# Python XLSB Reader & Writer

A Python library for reading and writing XLSB and XLSX files efficiently.

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
| **xlspy (C_EXT)** | XLSB | **1.02 s** | 7.20 MB |
| xlspy (Python) | XLSB | 2.54 s | 7.20 MB |
| xlspy | XLSX | 3.66 s | 6.32 MB |
| [xlsxwriter](https://pypi.org/project/xlsxwriter/) | XLSX | 8.67 s | 11.44 MB |

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
