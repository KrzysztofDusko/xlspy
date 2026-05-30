# -*- coding: utf-8 -*-


class F:
    """Format string constants for cell number formatting.
    
    Use with writers by passing (value, format_string) tuples:
        writer.write_sheet([
            ["Name", "Value"],
            ["Revenue", (100000, F.THOUSANDS_SEP)],
            ["Date", (date(2026, 6, 1), F.DATE_SHORT)],
        ])
    """
    THOUSANDS_SEP = '#,##0'
    CURRENCY_PLN = '#,##0.00 "zł"'
    CURRENCY_EUR = '#,##0.00 €'
    PERCENTAGE = '0%'
    SCIENTIFIC = '0.00E+00'
    TWO_DECIMALS = '#,##0.00'
    TEXT = '@'
    LEADING_ZEROS = '000000000'

    DATE_SHORT = 'dd.mm.yyyy'
    DATE_LONG = 'd mmmm yyyy'
    DATE_DAY_MONTH_YEAR = 'dd-mm-yyyy'
    DATE_ISO = 'yyyy-mm-dd'
    DATE_MONTH_YEAR = 'mmmm yyyy'
    DATE_WEEKDAY = 'dddd, d mmmm yyyy'
    DATE_DAY_MONTH = 'd mmmm'
    DATE_YEAR_ONLY = 'yyyy'

    DATETIME_SHORT = 'dd.mm.yyyy hh:mm'
    DATETIME_LONG = 'd mmmm yyyy hh:mm:ss'
    TIME_HH_MM = 'hh:mm'
    TIME_HH_MM_SS = 'hh:mm:ss'
    TIME_12H = 'h:mm AM/PM'
    DATETIME_24H = 'dd.mm.yyyy hh:mm:ss'
    DATETIME_ISO = 'yyyy-mm-dd"T"hh:mm:ss'
    TIME_MS = 'hh:mm:ss.000'

    SHORT_DATE = 'dd.mm.yyyy'
    LONG_DATE = 'd mmmm yyyy'
    ISO_DATE = 'yyyy-mm-dd'
    SHORT_DATETIME = 'dd.mm.yyyy hh:mm'
    LONG_DATETIME = 'd mmmm yyyy hh:mm:ss'
    ISO_DATETIME = 'yyyy-mm-dd"T"hh:mm:ss'


def _is_formatted_cell(cell):
    return isinstance(cell, tuple) and len(cell) == 2 and isinstance(cell[1], str)


def _unwrap_cell(cell):
    if _is_formatted_cell(cell):
        return cell[0]
    return cell


def _get_format(cell):
    if _is_formatted_cell(cell):
        return cell[1]
    return None
