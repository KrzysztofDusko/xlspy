import os
import sys
import time
from typing import Generator
import pyodbc
import xlsxwriter

# Add src to path for local testing of xlspy
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from xlspy import XlsbWriter

# --- CONFIGURATION ---
# Ensure you have pyodbc and xlsxwriter installed:
# pip install pyodbc xlsxwriter

# CHANGE THE VALUES BELOW TO MATCH YOUR DATABASE
password = os.environ.get("NZ_DEV_PASSWORD", "password")
DSN = f"DRIVER={{NetezzaSQL}};SERVER=linux.local;PORT=5480;DATABASE=JUST_DATA;UID=admin;PWD={password};"
# A query that returns a significant amount of data for a meaningful benchmark
QUERY = "SELECT * FROM JUST_DATA..FACTRESELLERSALES"
XLSB_FILENAME = "benchmark_output.xlsb"
XLSX_FILENAME = "benchmark_output.xlsx"


def install_and_import(package):
    """Tries to import a package, installing it if not found."""
    try:
        __import__(package)
    except ImportError:
        print(f"{package} not found. Installing...")
        import subprocess
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except subprocess.CalledProcessError:
            print(f"Failed to install {package}. Please install it manually: pip install {package}")
            sys.exit(1)
    finally:
        globals()[package] = __import__(package)


def row_generator(cursor) -> Generator[list[any], None, None]:
    """
    Generates rows from a pyodbc cursor, yielding headers first, then data rows.
    """
    try:
        headers = [column[0] for column in cursor.description]
        yield headers
        while row := cursor.fetchone():
            yield list(row)
    except pyodbc.Error as e:
        print(f"Error processing cursor data: {e}")
        raise


def run_xlspy_benchmark(cursor):
    """Runs the benchmark for xlspy."""
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
    file_size = os.path.getsize(XLSB_FILENAME) / (1024 * 1024)  # in MB

    print(f"Time taken: {duration:.2f} seconds")
    print(f"File size: {file_size:.2f} MB")
    
    return duration, file_size


def run_xlsxwriter_benchmark(cursor):
    """Runs the benchmark for xlsxwriter."""
    print("\n--- Running xlsxwriter Benchmark ---")
    if os.path.exists(XLSX_FILENAME):
        os.remove(XLSX_FILENAME)

    start_time = time.time()

    cursor.execute(QUERY)
    # XlsxWriter in optimized mode for writing large files
    workbook = xlsxwriter.Workbook(XLSX_FILENAME, {'constant_memory': True})
    worksheet = workbook.add_worksheet("Benchmark Data")
    
    row_num = 0
    for row in row_generator(cursor):
        worksheet.write_row(row_num, 0, row)
        row_num += 1

    workbook.close()

    end_time = time.time()
    duration = end_time - start_time
    file_size = os.path.getsize(XLSX_FILENAME) / (1024 * 1024)  # in MB

    print(f"Time taken: {duration:.2f} seconds")
    print(f"File size: {file_size:.2f} MB")

    return duration, file_size



def main():
    """Main function to run the benchmarks."""
    # Ensure required packages are installed
    install_and_import("pyodbc")
    install_and_import("xlsxwriter")

    try:
        print(f"Connecting to the database...")
        with pyodbc.connect(DSN, autocommit=True) as conn:
            cursor = conn.cursor()
            print(f"Executing query to get row count for benchmark info...")
            # Note: COUNT(*) can be slow on large tables without proper indexing.
            # This is just for context.
            try:
                row_count_cursor = conn.cursor()
                row_count_cursor.execute(f"SELECT COUNT(*) FROM ({QUERY}) as subquery")
                row_count = row_count_cursor.fetchone()[0]
                print(f"Benchmark will process approximately {row_count:,} rows.")
            except pyodbc.Error as e:
                print(f"Could not get row count: {e}. Continuing without it.")


            # Run benchmarks
            xlsb_duration, xlsb_size = run_xlspy_benchmark(cursor)
            xlsx_duration, xlsx_size = run_xlsxwriter_benchmark(cursor)

            # --- Summary ---
            print("\n--- Benchmark Summary ---")
            print(f"xlspy: {xlsb_duration:.2f}s, {xlsb_size:.2f} MB")
            print(f"xlsxwriter:   {xlsx_duration:.2f}s, {xlsx_size:.2f} MB")
            
            time_diff = abs(xlsb_duration - xlsx_duration)
            size_diff = abs(xlsb_size - xlsx_size)
            
            if xlsb_duration < xlsx_duration:
                print(f"\nxlspy was {time_diff:.2f}s faster ({xlsx_duration/xlsb_duration:.2f}x).")
            else:
                print(f"\nxlsxwriter was {time_diff:.2f}s faster ({xlsb_duration/xlsx_duration:.2f}x).")

            if xlsb_size < xlsx_size:
                print(f"xlspy produced a {size_diff:.2f} MB smaller file ({xlsx_size/xlsb_size:.2f}x).")
            else:
                print(f"xlsxwriter produced a {size_diff:.2f} MB smaller file ({xlsb_size/xlsx_size:.2f}x).")


    except pyodbc.Error as ex:
        sqlstate = ex.args[0]
        print(f"\nDatabase connection or query execution error: {sqlstate}\n{ex}")
        print("Please ensure your DSN and query are configured correctly at the top of the script.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
    finally:
        # Clean up generated files
        print("\nCleaning up generated files...")
        if os.path.exists(XLSB_FILENAME):
            os.remove(XLSB_FILENAME)
        if os.path.exists(XLSX_FILENAME):
            os.remove(XLSX_FILENAME)
        print("Done.")


if __name__ == '__main__':
    main()
