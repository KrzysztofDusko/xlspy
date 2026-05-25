import os
import sys
import tempfile
import datetime

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from xlspy import XlsbWriter, XlsxWriter, ExcelReader


def test_read_xlsb():
    """Write data with XlsbWriter, then read it back with ExcelReader."""
    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        data = [
            ['Name', 'Age', 'Active', 'Score', 'Birth'],
            ['Alice', 30, True, 95.5, datetime.date(1990, 5, 15)],
            ['Bob', 25, False, 88.0, datetime.date(1995, 8, 20)],
            ['Charlie', 35, True, 100.0, datetime.date(1985, 3, 10)],
            [None, None, None, None, None],
            ['Eve', 28, True, 72.5, datetime.date(1992, 12, 1)],
        ]

        with XlsbWriter(tmpfile) as writer:
            writer.add_sheet("TestSheet1")
            writer.write_sheet(data)

        with ExcelReader() as reader:
            reader.open(tmpfile)
            names = reader.get_sheet_names()
            assert 'TestSheet1' in names, f"Sheet not found: {names}"
            print(f"  Sheets: {names}")

            rows = reader.read_all('TestSheet1')
            assert len(rows) == len(data), f"Expected {len(data)} rows, got {len(rows)}"
            for i, (expected, actual) in enumerate(zip(data, rows)):
                for j, (e, a) in enumerate(zip(expected, actual)):
                    if e is None:
                        assert a is None or a == '', f"Row {i} col {j}: expected None, got {a!r}"
                    elif isinstance(e, datetime.date) and not isinstance(e, datetime.datetime):
                        assert isinstance(a, (datetime.date, datetime.datetime)), f"Row {i} col {j}: expected date, got {type(a)} {a!r}"
                    elif isinstance(e, bool):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r}"
                    elif isinstance(e, int):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r} (type={type(a)})"
                    elif isinstance(e, float):
                        assert abs(a - e) < 0.001, f"Row {i} col {j}: expected {e}, got {a!r}"
                    else:
                        assert a == e, f"Row {i} col {j}: expected {e!r}, got {a!r}"
            print(f"  Read {len(rows)} rows, all values match!")
    finally:
        os.remove(tmpfile)


def test_read_xlsx():
    """Write data with XlsxWriter, then read it back with ExcelReader."""
    tmpfile = tempfile.mktemp(suffix='.xlsx')
    try:
        data = [
            ['Name', 'Age', 'Active', 'Score', 'Birth'],
            ['Alice', 30, True, 95.5, datetime.date(1990, 5, 15)],
            ['Bob', 25, False, 88.0, datetime.date(1995, 8, 20)],
            ['Charlie', 35, True, 100.0, datetime.date(1985, 3, 10)],
            ['Eve', 28, True, 72.5, datetime.date(1992, 12, 1)],
        ]

        with XlsxWriter(tmpfile) as writer:
            writer.add_sheet("TestSheet1")
            writer.write_sheet(data)

        with ExcelReader() as reader:
            reader.open(tmpfile)
            names = reader.get_sheet_names()
            assert 'TestSheet1' in names, f"Sheet not found: {names}"
            print(f"  Sheets: {names}")

            rows = reader.read_all('TestSheet1')
            assert len(rows) == len(data), f"Expected {len(data)} rows, got {len(rows)}"
            for i, (expected, actual) in enumerate(zip(data, rows)):
                for j, (e, a) in enumerate(zip(expected, actual)):
                    if e is None:
                        assert a is None or a == '', f"Row {i} col {j}: expected None, got {a!r}"
                    elif isinstance(e, datetime.date) and not isinstance(e, datetime.datetime):
                        assert isinstance(a, (datetime.date, datetime.datetime)), f"Row {i} col {j}: expected date, got {type(a)} {a!r}"
                    elif isinstance(e, bool):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r}"
                    elif isinstance(e, int):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r} (type={type(a)})"
                    elif isinstance(e, float):
                        assert abs(a - e) < 0.001, f"Row {i} col {j}: expected {e}, got {a!r}"
                    else:
                        assert a == e, f"Row {i} col {j}: expected {e!r}, got {a!r}"
            print(f"  Read {len(rows)} rows, all values match!")
    finally:
        os.remove(tmpfile)


def test_multi_sheet_xlsb():
    """Test reading multiple sheets from XLSB."""
    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        with XlsbWriter(tmpfile) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([['A', 'B'], [1, 2]])
            writer.add_sheet("Sheet2")
            writer.write_sheet([['X', 'Y', 'Z'], ['a', 'b', 'c']])

        with ExcelReader() as reader:
            reader.open(tmpfile)
            names = reader.get_sheet_names()
            print(f"  Sheets: {names}")
            assert len(names) == 2
            assert 'Sheet1' in names
            assert 'Sheet2' in names

            rows1 = reader.read_all('Sheet1')
            assert len(rows1) == 2
            assert rows1[0] == ['A', 'B']

            rows2 = reader.read_all('Sheet2')
            assert len(rows2) == 2
            assert rows2[1] == ['a', 'b', 'c']
            print(f"  Multi-sheet XLSB OK!")
    finally:
        os.remove(tmpfile)


def test_multi_sheet_xlsx():
    """Test reading multiple sheets from XLSX."""
    tmpfile = tempfile.mktemp(suffix='.xlsx')
    try:
        with XlsxWriter(tmpfile) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([['A', 'B'], [1, 2]])
            writer.add_sheet("Sheet2")
            writer.write_sheet([['X', 'Y', 'Z'], ['a', 'b', 'c']])

        with ExcelReader() as reader:
            reader.open(tmpfile)
            names = reader.get_sheet_names()
            print(f"  Sheets: {names}")
            assert len(names) == 2
            assert 'Sheet1' in names
            assert 'Sheet2' in names

            rows1 = reader.read_all('Sheet1')
            assert len(rows1) == 2
            assert rows1[0] == ['A', 'B']

            rows2 = reader.read_all('Sheet2')
            assert len(rows2) == 2
            assert rows2[1] == ['a', 'b', 'c']
            print(f"  Multi-sheet XLSX OK!")
    finally:
        os.remove(tmpfile)


def test_row_generator_xlsb():
    """Test reading XLSB using generator."""
    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        with XlsbWriter(tmpfile) as writer:
            writer.add_sheet("Data")
            writer.write_sheet([['h1', 'h2'], ['v1', 'v2'], ['v3', 'v4']])

        with ExcelReader() as reader:
            reader.open(tmpfile)
            count = 0
            for row in reader.get_rows("Data"):
                count += 1
                assert len(row) == 2
            assert count == 3, f"Expected 3 rows (header + 2 data), got {count}"
            print(f"  Generator XLSB OK!")
    finally:
        os.remove(tmpfile)


def test_sheet_not_found():
    """Test error on reading non-existent sheet."""
    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        with XlsbWriter(tmpfile) as writer:
            writer.add_sheet("OnlySheet")
            writer.write_sheet([['a']])

        with ExcelReader() as reader:
            reader.open(tmpfile)
            try:
                reader.read_all("NonExistent")
                assert False, "Should have raised an exception"
            except (KeyError, ValueError):
                print(f"  Sheet not found error correctly raised!")
    finally:
        os.remove(tmpfile)


def test_read_complex_xlsx_values():
    """Test reading various data types from XLSX."""
    tmpfile = tempfile.mktemp(suffix='.xlsx')
    try:
        data = [
            ['Text', 'Integer', 'Float', 'Boolean', 'Date'],
            ['Hello', 42, 3.14, True, '2024-01-15'],
            ['World', -7, -0.001, False, '2024-06-30'],
            ['', 0, 1e10, True, ''],
            ['Spaces  ', 1000000, 0.5, False, None],
        ]

        with XlsxWriter(tmpfile) as writer:
            writer.add_sheet("Types")
            writer.write_sheet(data)

        with ExcelReader() as reader:
            reader.open(tmpfile)
            rows = reader.read_all('Types')
            assert len(rows) == len(data)
            for i, (expected, actual) in enumerate(zip(data, rows)):
                for j, (e, a) in enumerate(zip(expected, actual)):
                    if e is None:
                        assert a is None or a == '', f"Row {i} col {j}: expected None, got {a!r}"
                    elif isinstance(e, str):
                        assert a == e, f"Row {i} col {j}: expected {e!r}, got {a!r}"
                    elif isinstance(e, bool):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r}"
                    elif isinstance(e, int):
                        assert a == e, f"Row {i} col {j}: expected {e}, got {a!r} (type={type(a)})"
                    elif isinstance(e, float):
                        assert abs(a - e) < 0.001, f"Row {i} col {j}: expected {e}, got {a!r}"
            print(f"  Complex XLSX values OK!")
    finally:
        os.remove(tmpfile)


def test_reader_with_path_in_constructor():
    """Regression: ``with ExcelReader(path)`` must auto-open the file."""
    tmpfile = tempfile.mktemp(suffix='.xlsb')
    try:
        data = [["Name", "Value"], ["foo", 42]]
        with XlsbWriter(tmpfile) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet(data)

        # This pattern is documented on PyPI and must work:
        with ExcelReader(tmpfile) as reader:
            names = reader.get_sheet_names()
            assert 'Sheet1' in names, \
                f"ExcelReader(path) with-statement: sheets={names!r} (empty = bug!)"

            rows = reader.read_all('Sheet1')
            assert len(rows) == 2
            assert rows[0] == ['Name', 'Value']
            assert rows[1] == ['foo', 42]

        # Also test ExcelReader() + .open(path) still works (backward compat)
        with ExcelReader() as reader:
            reader.open(tmpfile)
            assert 'Sheet1' in reader.get_sheet_names()

        print("  Reader with path in constructor: OK")
    finally:
        os.remove(tmpfile)


if __name__ == '__main__':
    print("=== ExcelReader Tests ===")
    print("\n1. XLSB read/write roundtrip...")
    test_read_xlsb()
    print("\n2. XLSX read/write roundtrip...")
    test_read_xlsx()
    print("\n3. Multi-sheet XLSB...")
    test_multi_sheet_xlsb()
    print("\n4. Multi-sheet XLSX...")
    test_multi_sheet_xlsx()
    print("\n5. Row generator XLSB...")
    test_row_generator_xlsb()
    print("\n6. Sheet not found error handling...")
    test_sheet_not_found()
    print("\n7. Complex XLSX values...")
    test_read_complex_xlsx_values()
    print("\n8. Reader with path in constructor...")
    test_reader_with_path_in_constructor()
    print("\n=== All tests passed! ===")
