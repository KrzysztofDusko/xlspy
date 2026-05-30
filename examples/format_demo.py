# -*- coding: utf-8 -*-
"""
Example demonstrating cell number/date/datetime formatting in xlspy.
Creates both .xlsx and .xlsb files with all supported formats.

Usage:
    python examples/format_demo.py
"""
import datetime
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from xlspy import XlsbWriter, XlsxWriter, F


VALUE = 100000
DATE_VALUE = datetime.date(2026, 6, 1)
DATETIME_VALUE = datetime.datetime(2026, 6, 1, 14, 34, 20)


def number_formats_sheet():
    return [
        ["Styl formatowania", "Wartość"],
        ["Bez formatowania", VALUE],
        ["Z separatorem tysięcy", (VALUE, F.THOUSANDS_SEP)],
        ["Waluta PLN", (VALUE, F.CURRENCY_PLN)],
        ["Waluta EUR", (VALUE, F.CURRENCY_EUR)],
        ["Procent (jako 100000%)", (VALUE, F.PERCENTAGE)],
        ["Notacja naukowa", (VALUE, F.SCIENTIFIC)],
        ["Z dwoma miejscami po przecinku", (VALUE, F.TWO_DECIMALS)],
        ["Tekst", (VALUE, F.TEXT)],
        ["Z zerami wiodącymi", (VALUE, F.LEADING_ZEROS)],
    ]


def date_formats_sheet():
    return [
        ["Styl formatowania daty", "Wartość"],
        ["Bez formatowania", DATE_VALUE],
        ["Data krótka (dd.MM.yyyy)", (DATE_VALUE, F.DATE_SHORT)],
        ["Data długa", (DATE_VALUE, F.DATE_LONG)],
        ["Dzień-miesiąc-rok", (DATE_VALUE, F.DATE_DAY_MONTH_YEAR)],
        ["Rok-miesiąc-dzień (ISO)", (DATE_VALUE, F.DATE_ISO)],
        ["Miesiąc słownie", (DATE_VALUE, F.DATE_MONTH_YEAR)],
        ["Dzień tygodnia + data", (DATE_VALUE, F.DATE_WEEKDAY)],
        ["Tylko dzień i miesiąc", (DATE_VALUE, F.DATE_DAY_MONTH)],
        ["Tylko rok", (DATE_VALUE, F.DATE_YEAR_ONLY)],
        [None, None],
        [None, None],
        ["Styl formatowania daty i czasu", "Wartość"],
        ["Bez formatowania", DATETIME_VALUE],
        ["Data i czas krótki", (DATETIME_VALUE, F.DATETIME_SHORT)],
        ["Data i czas długi", (DATETIME_VALUE, F.DATETIME_LONG)],
        ["Tylko czas (hh:mm)", (DATETIME_VALUE, F.TIME_HH_MM)],
        ["Czas z sekundami", (DATETIME_VALUE, F.TIME_HH_MM_SS)],
        ["Czas 12h (AM/PM)", (DATETIME_VALUE, F.TIME_12H)],
        ["Data + czas 24h", (DATETIME_VALUE, F.DATETIME_24H)],
        ["ISO format", (DATETIME_VALUE, F.DATETIME_ISO)],
        ["Czas z milisekundami", (DATETIME_VALUE, F.TIME_MS)],
    ]


if __name__ == "__main__":
    out_dir = os.path.dirname(__file__)

    print("Creating XLSX file with formatted cells...")
    with XlsxWriter(os.path.join(out_dir, "format_demo.xlsx")) as writer:
        writer.add_sheet("Formaty Liczb")
        writer.write_sheet(number_formats_sheet())
        writer.add_sheet("Formaty Dat")
        writer.write_sheet(date_formats_sheet())
    print("  -> format_demo.xlsx created")

    print("Creating XLSB file with formatted cells...")
    with XlsbWriter(os.path.join(out_dir, "format_demo.xlsb")) as writer:
        writer.add_sheet("Formaty Liczb")
        writer.write_sheet(number_formats_sheet())
        writer.add_sheet("Formaty Dat")
        writer.write_sheet(date_formats_sheet())
    print("  -> format_demo.xlsb created")

    print("\nDone! Open the files in Excel to see the formatting.")
    print("Compare with data_formats.xlsx / data_formats.xlsb in the project root.")
