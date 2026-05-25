import sys
import os
import time
import xlsxwriter
from memory_profiler import profile
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))
from xlspy import XlsbWriter
from xlspy import XlsxWriter

import datetime
from decimal import Decimal

def generate_large_data(rows, cols):
    """Generates a large list of lists for testing."""
    print(f"Generating {rows}x{cols} dataset...")

    headers = [f"Header_{i}" for i in range(cols)]
    yield headers
    for i in range(1, rows):
        row_data = []
        for j in range(cols):
            col_type = j % 8
            if col_type == 0:
                row_data.append(f"String_{i}_{j}")
            elif col_type == 1:
                row_data.append(i * j)
            elif col_type == 2:
                row_data.append(i / (j + 1))
            elif col_type == 3:
                row_data.append(True if (i + j) % 2 == 0 else False)
            elif col_type == 4:
                row_data.append(datetime.date(2023, 1, 1) + datetime.timedelta(days=i))
            elif col_type == 5:
                row_data.append(datetime.datetime(2023, 1, 1, 12, 30, 0) + datetime.timedelta(seconds=i))
            elif col_type == 6:
                row_data.append(Decimal(f"{i}.{j}"))
            elif col_type == 7:
                row_data.append(2**60 + i)  # Large integer
        yield row_data

# @profile
def run_my_xlsb_writer(data, filename="test_xlspy.xlsb"):
    """Tests the performance of XlsbWriter."""
    print("\n--- Testing (my) XlsbWriter ---")
    start_time = time.time()
    
    with XlsbWriter(filename) as writer:
        writer.add_sheet("LargeSheet")
        writer.write_sheet(data)

    
    end_time = time.time()
    elapsed = end_time - start_time
    
    filesize = os.path.getsize(filename) / (1024 * 1024)
    print(f"XlsbWriter finished in: {elapsed:.2f} seconds")
    print(f"File size: {filesize:.2f} MB")
    print("--------------------------")

# @profile
def run_my_xlsx_writer(data, filename="test_xlspy.xlsx"):
    """Tests the performance of XlsxWriter."""
    print("\n--- Testing (my)  XlsxWriter ---")
    start_time = time.time()
    
    with XlsxWriter(filename) as writer:
        writer.add_sheet("LargeSheet")
        writer.write_sheet(data)

    
    end_time = time.time()
    elapsed = end_time - start_time
    
    filesize = os.path.getsize(filename) / (1024 * 1024)
    print(f"XlsbWriter finished in: {elapsed:.2f} seconds")
    print(f"File size: {filesize:.2f} MB")
    print("--------------------------")

# @profile
def run_xlsx_writer(data, filename="test_xlsxwriter.xlsx"):
    """Tests the performance of xlsxwriter."""
    print("\n--- Testing xlsxwriter ---")
    start_time = time.time()
    
    # Use constant_memory mode for performance with large datasets
    workbook = xlsxwriter.Workbook(filename, {'constant_memory': True})
    worksheet = workbook.add_worksheet("LargeSheet")
    
    for row_idx, row_data in enumerate(data):
        worksheet.write_row(row_idx, 0, row_data)
        
    workbook.close()
    
    end_time = time.time()
    elapsed = end_time - start_time
    
    filesize = os.path.getsize(filename) / (1024 * 1024)
    print(f"xlsxwriter finished in: {elapsed:.2f} seconds")
    print(f"File size: {filesize:.2f} MB")
    print("--------------------------")

if __name__ == "__main__":
    # Define the size of the dataset
    NUM_ROWS = 50000
    NUM_COLS = 50
    
    files_to_clean = []
    
    # Run tests
    f1 = "test_xlspy.xlsb"
    run_my_xlsb_writer(generate_large_data(NUM_ROWS, NUM_COLS), f1)
    files_to_clean.append(f1)
    
    f2 = "test_xlspy.xlsx"
    run_my_xlsx_writer(generate_large_data(NUM_ROWS, NUM_COLS), f2)
    files_to_clean.append(f2)
    
    f3 = "test_xlsxwriter.xlsx"
    run_xlsx_writer(generate_large_data(NUM_ROWS, NUM_COLS), f3)
    files_to_clean.append(f3)
    
    # Cleanup
    for f in files_to_clean:
        try:
            os.remove(f)
        except OSError:
            pass
    print("Temporary benchmark files removed.")
    
    print("\nPerformance test complete.")
    print("To see detailed memory usage, run this script with memory-profiler:")
    print("mprof run performance_test.py")
    print("Then, to see the plot:")
    print("mprof plot")
