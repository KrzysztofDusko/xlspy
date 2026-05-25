import os
import sys
from typing import Generator

import pytest
nzpy = pytest.importorskip("nzpy_extended.sync")

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from xlspy import XlsbWriter, XlsxWriter

NZ_CONFIG = {
    "host": os.environ.get("NZ_DEV_HOST", "linux.local"),
    "port": int(os.environ.get("NZ_DEV_PORT", "5480")),
    "database": os.environ.get("NZ_DEV_DB", "JUST_DATA"),
    "user": os.environ.get("NZ_DEV_USER", "admin"),
    "password": os.environ.get("NZ_DEV_PASSWORD", "password"),
}
QUERY1 = "SELECT * FROM JUST_DATA..DIMACCOUNT"
QUERY2 = "SELECT * FROM JUST_DATA..DIMCURRENCY"
OUTPUT_FILENAME = "real_netezza_output.xlsb"


def row_generator(cursor) -> Generator[list, None, None]:
    headers = [column[0] for column in cursor.description]
    yield headers
    while row := cursor.fetchone():
        yield list(row)


def main():
    if os.path.exists(OUTPUT_FILENAME):
        os.remove(OUTPUT_FILENAME)

    try:
        print(f"Connecting to {NZ_CONFIG['host']}:{NZ_CONFIG['port']}/{NZ_CONFIG['database']}...")
        with nzpy.connect(**NZ_CONFIG) as conn:
            cursor = conn.cursor()
            print(f"Executing query: '{QUERY1}'")
            cursor.execute(QUERY1)

            print(f"Writing data to file '{OUTPUT_FILENAME}'...")
            with XlsbWriter(OUTPUT_FILENAME) as writer:
                writer.add_sheet("Netezza Export 1")
                writer.write_sheet(row_generator(cursor))
                writer.add_sheet("SQL 1", hidden=True)
                writer.write_sheet([["code"], [QUERY1]])
                writer.add_sheet("Netezza Export 2")
                print(f"Executing query: '{QUERY2}'")
                cursor.execute(QUERY2)
                writer.write_sheet(row_generator(cursor))
                writer.add_sheet("SQL 2", hidden=True)
                writer.write_sheet([["code"], [QUERY2]])
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

    if os.path.exists(OUTPUT_FILENAME) and os.path.getsize(OUTPUT_FILENAME) > 0:
        print(f"\nTest completed successfully, file '{OUTPUT_FILENAME}' was created and contains data.")
    else:
        print(f"\nTest failed, file '{OUTPUT_FILENAME}' was not created or is empty.")
    if os.path.exists(OUTPUT_FILENAME):
        os.remove(OUTPUT_FILENAME)


def test_xlsx_writer():
    OUTPUT_XLSX_FILENAME = "real_netezza_output.xlsx"
    if os.path.exists(OUTPUT_XLSX_FILENAME):
        os.remove(OUTPUT_XLSX_FILENAME)

    try:
        print(f"Connecting to {NZ_CONFIG['host']}:{NZ_CONFIG['port']}/{NZ_CONFIG['database']}...")
        with nzpy.connect(**NZ_CONFIG) as conn:
            cursor = conn.cursor()
            print(f"Executing query: '{QUERY1}'")
            cursor.execute(QUERY1)

            print(f"Writing data to file '{OUTPUT_XLSX_FILENAME}'...")
            with XlsxWriter(OUTPUT_XLSX_FILENAME) as writer:
                writer.add_sheet("Netezza Export 1")
                writer.write_sheet(row_generator(cursor))
                writer.add_sheet("SQL 1", hidden=True)
                writer.write_sheet([["code"], [QUERY1]])
                writer.add_sheet("Netezza Export 2")
                print(f"Executing query: '{QUERY2}'")
                cursor.execute(QUERY2)
                writer.write_sheet(row_generator(cursor))
                writer.add_sheet("SQL 2", hidden=True)
                writer.write_sheet([["code"], [QUERY2]])
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

    if os.path.exists(OUTPUT_XLSX_FILENAME) and os.path.getsize(OUTPUT_XLSX_FILENAME) > 0:
        print(f"\nTest completed successfully, file '{OUTPUT_XLSX_FILENAME}' was created and contains data.")
    else:
        print(f"\nTest failed, file '{OUTPUT_XLSX_FILENAME}' was not created or is empty.")
    if os.path.exists(OUTPUT_XLSX_FILENAME):
        os.remove(OUTPUT_XLSX_FILENAME)


if __name__ == '__main__':
    try:
        import nzpy_extended.sync as nzpy
        main()
        print("\n" + "=" * 50)
        print("Testing XlsxWriter")
        print("=" * 50)
        test_xlsx_writer()
    except ImportError:
        print("nzpy-extended is not installed. Installing...")
        import subprocess
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "nzpy-extended"])
            import nzpy_extended.sync as nzpy
            main()
            print("\n" + "=" * 50)
            print("Testing XlsxWriter")
            print("=" * 50)
            test_xlsx_writer()
        except subprocess.CalledProcessError:
            print("Failed to install nzpy-extended. Install manually: pip install nzpy-extended")
            sys.exit(1)
