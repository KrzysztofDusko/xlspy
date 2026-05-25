import os
import sys
import time
from typing import Generator

import xlsxwriter

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from xlspy import XlsbWriter

NZ_CONFIG = {
    "host": os.environ.get("NZ_DEV_HOST", "linux.local"),
    "port": int(os.environ.get("NZ_DEV_PORT", "5480")),
    "database": os.environ.get("NZ_DEV_DB", "JUST_DATA"),
    "user": os.environ.get("NZ_DEV_USER", "admin"),
    "password": os.environ.get("NZ_DEV_PASSWORD", "password"),
}
QUERY = "SELECT * FROM JUST_DATA..FACTRESELLERSALES"
XLSB_FILENAME = "benchmark_output.xlsb"
XLSX_FILENAME = "benchmark_output.xlsx"


def _ensure_import(package_name: str, import_name: str | None = None):
    import importlib, subprocess, sys
    try:
        return importlib.import_module(import_name or package_name)
    except ImportError:
        print(f"{package_name} not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        return importlib.import_module(import_name or package_name)


def row_generator(cursor) -> Generator[list, None, None]:
    headers = [column[0] for column in cursor.description]
    yield headers
    while row := cursor.fetchone():
        yield list(row)


def run_xlspy_benchmark(cursor):
    print("\n--- Running xlspy Benchmark ---")
    if os.path.exists(XLSB_FILENAME):
        os.remove(XLSB_FILENAME)

    start_time = time.time()
    cursor.execute(QUERY)
    with XlsbWriter(XLSB_FILENAME) as writer:
        writer.add_sheet("Benchmark Data")
        writer.write_sheet(row_generator(cursor))

    end_time = time.time()
    duration = end_time - start_time
    file_size = os.path.getsize(XLSB_FILENAME) / (1024 * 1024)
    print(f"Time taken: {duration:.2f} seconds")
    print(f"File size: {file_size:.2f} MB")
    return duration, file_size


def run_xlsxwriter_benchmark(cursor):
    print("\n--- Running xlsxwriter Benchmark ---")
    if os.path.exists(XLSX_FILENAME):
        os.remove(XLSX_FILENAME)

    start_time = time.time()
    cursor.execute(QUERY)
    workbook = xlsxwriter.Workbook(XLSX_FILENAME, {'constant_memory': True})
    worksheet = workbook.add_worksheet("Benchmark Data")
    for row_num, row in enumerate(row_generator(cursor)):
        worksheet.write_row(row_num, 0, row)
    workbook.close()

    end_time = time.time()
    duration = end_time - start_time
    file_size = os.path.getsize(XLSX_FILENAME) / (1024 * 1024)
    print(f"Time taken: {duration:.2f} seconds")
    print(f"File size: {file_size:.2f} MB")
    return duration, file_size


def main():
    nzpy = _ensure_import("nzpy-extended", "nzpy_extended.sync")
    _ensure_import("xlsxwriter")

    try:
        print(f"Connecting to {NZ_CONFIG['host']}:{NZ_CONFIG['port']}/{NZ_CONFIG['database']}...")
        with nzpy.connect(**NZ_CONFIG) as conn:
            cursor = conn.cursor()

            try:
                cnt_cursor = conn.cursor()
                cnt_cursor.execute(f"SELECT COUNT(*) FROM ({QUERY}) AS subq")
                row_count = cnt_cursor.fetchone()[0]
                print(f"Benchmark will process approximately {row_count:,} rows.")
            except Exception as e:
                print(f"Could not get row count: {e}. Continuing without it.")

            xlsb_duration, xlsb_size = run_xlspy_benchmark(cursor)
            xlsx_duration, xlsx_size = run_xlsxwriter_benchmark(cursor)

            print("\n--- Benchmark Summary ---")
            print(f"xlspy: {xlsb_duration:.2f}s, {xlsb_size:.2f} MB")
            print(f"xlsxwriter:   {xlsx_duration:.2f}s, {xlsx_size:.2f} MB")

            time_diff = abs(xlsb_duration - xlsx_duration)
            size_diff = abs(xlsb_size - xlsx_size)

            if xlsb_duration < xlsx_duration:
                print(f"\nxlspy was {time_diff:.2f}s faster ({xlsx_duration / xlsb_duration:.2f}x).")
            else:
                print(f"\nxlsxwriter was {time_diff:.2f}s faster ({xlsb_duration / xlsx_duration:.2f}x).")

            if xlsb_size < xlsx_size:
                print(f"xlspy produced a {size_diff:.2f} MB smaller file ({xlsx_size / xlsb_size:.2f}x).")
            else:
                print(f"xlsxwriter produced a {size_diff:.2f} MB smaller file ({xlsb_size / xlsx_size:.2f}x).")

    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
    finally:
        print("\nCleaning up generated files...")
        for f in (XLSB_FILENAME, XLSX_FILENAME):
            if os.path.exists(f):
                os.remove(f)
        print("Done.")


if __name__ == '__main__':
    main()
