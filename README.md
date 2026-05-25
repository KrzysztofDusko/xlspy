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

### Streaming from a Database (ODBC)

This example shows how to stream data directly from a database query into an XLSB file. This is highly memory-efficient as it doesn't load the entire dataset into memory.

First, ensure you have `pyodbc` installed:
```bash
pip install pyodbc
```

Then, you can use a generator function to feed data to `XlsbWriter`.

```python
import os
import pyodbc
from typing import Generator
from xlspy import XlsbWriter

# --- Configuration ---
# Make sure you have an ODBC driver and a configured DSN, or use a DSN-less connection string.
DSN = "DRIVER={Your ODBC Driver};SERVER=your_server;DATABASE=your_db;UID=your_user;PWD=your_password"
QUERY = "SELECT * FROM YourTable"
OUTPUT_FILENAME = "db_output.xlsb"

def row_generator(cursor: pyodbc.Cursor) -> Generator[list[any], None, None]:
    """
    Generates rows from a pyodbc cursor, yielding headers first, followed by data rows.
    """
    # Extract column headers from cursor description
    headers = [column[0] for column in cursor.description]
    yield headers

    # Yield each row until the cursor is exhausted
    while row := cursor.fetchone():
        yield list(row)

# --- Main Execution ---
try:
    # Connect to the database
    with pyodbc.connect(DSN) as conn:
        cursor = conn.cursor()
        cursor.execute(QUERY)

        # Use XlsbWriter to write the data stream
        with XlsbWriter(OUTPUT_FILENAME) as writer:
            writer.add_sheet("Database Export")
            writer.write_sheet(row_generator(cursor))
            
            # You can also add hidden sheets with metadata, like the query itself
            writer.add_sheet("SQL Query", hidden=True)
            writer.write_sheet([["SQL"], [QUERY]])


    print(f"Successfully created '{OUTPUT_FILENAME}'")

except pyodbc.Error as ex:
    sqlstate = ex.args[0]
    print(f"Database connection or query execution error: {sqlstate}\n{ex}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
```

## Performance

`xlspy` is designed for high performance. Since version 0.1.0, the library includes a **C extension** (`_c_core`) that accelerates XLSB read and write. The C extension is **enabled by default** (compiled automatically on install). Set `XLSPY_DISABLE_C_EXT=1` to force the pure Python fallback.

All benchmarks: **50000 × 50** dataset (2.5M cells). Tests performed on **Windows 11** (Python 3.14, AMD64).

### Write

| Library | Format | Time | Size |
|---------|--------|------|------|
| **xlspy (C_EXT)** | XLSB | **0.70 s** | 8.05 MB |
| xlspy (Python) | XLSB | 1.61 s | 8.05 MB |
| xlspy | XLSX | 2.07 s | 5.37 MB |
| [xlsxwriter](https://pypi.org/project/xlsxwriter/) | XLSX | 9.00 s | 11.14 MB |

### Read

| Library | Format | Time | Notes |
|---------|--------|------|-------|
| **xlspy (C_EXT)** | XLSB | **0.65 s** | default, compiled C |
| xlspy | XLSX | 4.48 s | uses expat XML parser (C) |
| xlspy (Python) | XLSB | 5.93 s | pure Python fallback |
| [openpyxl](https://pypi.org/project/openpyxl/) | XLSX | 5.68 s | read-only mode |


### Analysis

The **8.96× read speedup** comes from two factors:
- **~60–70%** — native C compilation, no interpreter overhead per record
- **~30–40%** — algorithm simplification: flat array indexed by `col − first_col` instead of `Dict[int, Any]`, no `isinstance` per cell, no `BiffReader.read_worksheet()` method call per record

Run the benchmarks yourself with `examples/performance_test.py`.

## Repository

<https://github.com/KrzysztofDusko/xlspy/>
