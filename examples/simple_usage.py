import datetime
from decimal import Decimal
import os
import sys
# Add src to path for local testing
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from xlspy import XlsbWriter
from xlspy import XlsxWriter

if __name__ == "__main__":
    data = [
        ["Name", "Age", "City", "info"],
        [-123,2147483647,2147483648,2147483999],
        ["x", "y", "z", datetime.datetime.today()],
        ["Alice", 25, "New York",datetime.date.today()],
        ["Bob", 30, "London",Decimal(3.14)],
        ["Charlie", 35, "Paris",datetime.datetime.now()],
        [True, False, None,datetime.datetime.now(datetime.timezone.utc)]
    ]
    data2 = [
        ["Name", "Age", "City", "info"],
        [-2,2147483647,2147483648,2147483999],
        ["x2", "y", "z", datetime.datetime.today()],
        ["2Alice", 25, "New York",datetime.date.today()],
        ["2Bob", 30, "London",Decimal(5.14)],
        ["2Charlie", 35, "Paris",datetime.datetime.now()],
        [True, False, None,datetime.datetime.now(datetime.timezone.utc)],
        [True, False, None,datetime.datetime.now(datetime.timezone.utc)]
    ]
    with XlsbWriter("output.xlsb") as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(data)
        writer.add_sheet("Sheet2")
        writer.write_sheet(data2)
    print("XLSB file created successfully!")      

    with XlsxWriter("output.xlsx") as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(data)
        writer.add_sheet("Sheet2")
        writer.write_sheet(data2)   
    print("XLSX file created successfully!")
