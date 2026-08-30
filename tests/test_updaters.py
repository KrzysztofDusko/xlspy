import datetime
import os
import struct
import zipfile

import pytest

from xlspy import ExcelReader, F, XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter
from xlspy.biff_utils import build_record, iter_records, read_utf16


def _write_xlsx(path, sheets, *, use_shared_strings=True):
    with XlsxWriter(path, useSharedStrings=use_shared_strings) as writer:
        for name, rows in sheets:
            writer.add_sheet(name)
            writer.write_sheet(rows)


def _write_xlsb(path, sheets):
    with XlsbWriter(path) as writer:
        for name, rows in sheets:
            writer.add_sheet(name)
            writer.write_sheet(rows)


def _read_sheet(path, name):
    reader = ExcelReader()
    reader.open(str(path))
    try:
        return list(reader.get_rows(name))
    finally:
        reader.close()


def _add_zip_members(path, members):
    with zipfile.ZipFile(path, "a") as archive:
        for name, data in members.items():
            archive.writestr(name, data)


def _make_xlsx_pivot(path):
    _add_zip_members(
        path,
        {
            "xl/pivotCache/pivotCacheDefinition1.xml": (
                '<pivotCacheDefinition recordCount="1">'
                '<cacheSource type="worksheet">'
                '<worksheetSource ref="A1:B2" sheet="Data"/>'
                "</cacheSource></pivotCacheDefinition>"
            ).encode(),
            "xl/pivotTables/pivotTable1.xml": b'<pivotTableDefinition refreshOnLoad="0"/>',
        },
    )


def _make_xlsb_pivot(path):
    sheet_name = "Data"
    encoded_name = sheet_name.encode("utf-16le")
    source_payload = (
        b"\x00\x00\x00"
        + struct.pack("<I", len(sheet_name))
        + encoded_name
        + struct.pack("<iiii", 0, 1, 0, 1)
    )
    cache = build_record(0xB3, b"\x00\x00\x00\x00") + build_record(0xBB, source_payload)
    _add_zip_members(path, {"xl/pivotCache/pivotCacheDefinition1.bin": cache})


def _xlsx_sheet_xml(path):
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/worksheets/sheet1.xml").decode("utf-8")


def _xlsb_sheet_bytes(path):
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/worksheets/sheet1.bin")


def _set_xlsx_date_system(path, enabled=True):
    with zipfile.ZipFile(path) as archive:
        workbook = archive.read("xl/workbook.xml").decode("utf-8")
    value = "1" if enabled else "0"
    workbook = workbook.replace(
        "<workbookPr ", f'<workbookPr date1904="{value}" '
    )
    if "<workbookPr" not in workbook:
        workbook = workbook.replace(
            "<workbook ", f'<workbookPr date1904="{value}"/><workbook '
        )
    _add_zip_members(path, {"xl/workbook.xml": workbook.encode("utf-8")})


def _set_xlsb_date_system(path, enabled=True):
    with zipfile.ZipFile(path) as archive:
        workbook = bytearray(archive.read("xl/workbook.bin"))
    for record in iter_records(bytes(workbook)):
        if record.record_id == 0x99 and record.length >= 4:
            flags = struct.unpack_from("<I", workbook, record.data_start)[0]
            struct.pack_into(
                "<I",
                workbook,
                record.data_start,
                flags | 0x01 if enabled else flags & ~0x01,
            )
            break
    _add_zip_members(path, {"xl/workbook.bin": bytes(workbook)})


def _set_xlsb_trailing_row_header(path):
    with zipfile.ZipFile(path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.bin")
    insertion = next(
        record.header_start for record in iter_records(sheet) if record.record_id == 0x92
    )
    payload = bytearray(25)
    struct.pack_into("<i", payload, 0, 4)
    struct.pack_into("<i", payload, 8, 300)
    payload[13] = 1
    struct.pack_into("<i", payload, 17, 0)
    struct.pack_into("<i", payload, 21, 0)
    sheet = sheet[:insertion] + build_record(0x00, bytes(payload)) + sheet[insertion:]
    _add_zip_members(path, {"xl/worksheets/sheet1.bin": sheet})


def test_xlsx_updater_replaces_data_and_preserves_other_sheet(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    _write_xlsx(
        source,
        [("Data", [["old", 1], ["remove", 2]]), ("Keep", [["untouched", 9]])],
    )

    updater = XlsxUpdater(source)
    assert updater.get_sheet_names() == ["Data", "Keep"]
    updater.replace_sheet_data(
        "Data",
        ((row[0], row[1]) for row in [["new & <value>", 42], [None, True], ["middle", None], [None, None]]),
        headers=["Name", "Value"],
    )
    output.write_bytes(updater.to_bytes())
    updater.save(output)

    assert _read_sheet(output, "Data") == [
        ["Name", "Value"],
        ["new & <value>", 42],
        [None, True],
        ["middle", None],
    ]
    assert _read_sheet(output, "Keep") == [["untouched", 9]]
    sheet_xml = _xlsx_sheet_xml(output)
    assert 'ref="A1:B4"' in sheet_xml
    assert "old" not in sheet_xml


def test_xlsx_updater_supports_inline_strings_and_general_fallback(tmp_path):
    path = tmp_path / "inline.xlsx"
    _write_xlsx(path, [("Sheet1", [["old", 10], ["old2", 20]])], use_shared_strings=False)

    updater = XlsxUpdater(path)
    updater.replace_sheet_data("Sheet1", [["<&\"'\x01", 99]], style_fallback="general")
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert "xl/sharedStrings.xml" not in names
        assert 't="inlineStr"' in xml
        assert "&lt;&amp;&quot;&apos;" in xml
        assert ' s="' not in xml
    assert _read_sheet(path, "Sheet1")[0][0] == "<&\"'"


def test_xlsx_updater_inherits_date_and_header_styles(tmp_path):
    path = tmp_path / "styles.xlsx"
    _write_xlsx(
        path,
        [
            (
                "Sheet1",
                [["Date", "Amount"], [(datetime.date(2024, 1, 2), F.DATE_SHORT), (1.25, F.TWO_DECIMALS)]],
            )
        ],
    )
    updater = XlsxUpdater(path)
    updater.replace_sheet_data(
        "Sheet1",
        [[datetime.date(2025, 2, 3), 2.5]],
        headers=["Date", "Amount"],
    )
    path.write_bytes(updater.to_bytes())
    xml = _xlsx_sheet_xml(path)
    assert xml.count('s="3"') >= 2
    assert 's="1"' in xml or 's="4"' in xml
    assert isinstance(_read_sheet(path, "Sheet1")[1][0], (datetime.date, datetime.datetime))


def test_xlsx_updater_updates_pivot_metadata_and_refresh_flag(tmp_path):
    path = tmp_path / "pivot.xlsx"
    _write_xlsx(path, [("Data", [["old", 1], ["old2", 2]])])
    _make_xlsx_pivot(path)

    updater = XlsxUpdater(path)
    updater.replace_sheet_data("Data", [["one", 1], ["two", 2], ["three", 3]])
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        cache = archive.read("xl/pivotCache/pivotCacheDefinition1.xml").decode()
        table = archive.read("xl/pivotTables/pivotTable1.xml").decode()
        assert 'recordCount="3"' in cache
        assert 'ref="A1:B3"' in cache
        assert 'refreshOnLoad="1"' in cache
        assert 'refreshOnLoad="1"' in table


@pytest.mark.parametrize("binary", [False, True])
def test_updater_honors_1904_date_system(binary, tmp_path):
    path = tmp_path / ("date1904.xlsb" if binary else "date1904.xlsx")
    if binary:
        _write_xlsb(path, [("Sheet1", [[1]])])
        _set_xlsb_date_system(path)
        updater = XlsbUpdater(path)
    else:
        _write_xlsx(path, [("Sheet1", [[1]])])
        _set_xlsx_date_system(path)
        updater = XlsxUpdater(path)

    updater.replace_sheet_data("Sheet1", [[datetime.date(2024, 1, 1)]])
    path.write_bytes(updater.to_bytes())

    expected_serial = 43830.0
    if binary:
        sheet = _xlsb_sheet_bytes(path)
        date_values = [
            struct.unpack_from("<d", sheet, record.data_start + 8)[0]
            for record in iter_records(sheet)
            if record.record_id == 0x05
        ]
        assert date_values == [expected_serial]
    else:
        assert f"<v>{expected_serial:.15g}</v>" in _xlsx_sheet_xml(path)


def test_xlsb_updater_detects_custom_date_format_for_unstyled_column(tmp_path):
    path = tmp_path / "custom-date.xlsb"
    _write_xlsb(
        path,
        [("Sheet1", [[None, None, (datetime.date(2024, 1, 1), F.DATE_SHORT)]])],
    )

    updater = XlsbUpdater(path)
    updater.replace_sheet_data("Sheet1", [[datetime.date(2025, 2, 3)]])
    path.write_bytes(updater.to_bytes())

    sheet = _xlsb_sheet_bytes(path)
    styles = [
        struct.unpack_from("<I", sheet, record.data_start + 4)[0] & 0xFFFFFF
        for record in iter_records(sheet)
        if record.record_id == 0x05
    ]
    assert styles and styles[0] > 0


def test_xlsx_updater_preserves_worksheet_and_sst_prefixes(tmp_path):
    path = tmp_path / "prefix.xlsx"
    _write_xlsx(path, [("Sheet1", [["old"]])])
    with zipfile.ZipFile(path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        strings = archive.read("xl/sharedStrings.xml").decode("utf-8")
    sheet = sheet.replace(
        "<worksheet ", '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    ).replace("</worksheet>", "</x:worksheet>")
    sheet = sheet.replace("<sheetData", "<x:sheetData").replace(
        "</sheetData>", "</x:sheetData>"
    )
    strings = strings.replace(
        "<sst ", '<x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    ).replace("</sst>", "</x:sst>")
    strings = strings.replace("<si>", "<x:si>").replace("</si>", "</x:si>")
    strings = strings.replace("<t", "<x:t").replace("</t>", "</x:t>")
    _add_zip_members(
        path,
        {
            "xl/worksheets/sheet1.xml": sheet.encode("utf-8"),
            "xl/sharedStrings.xml": strings.encode("utf-8"),
        },
    )

    updater = XlsxUpdater(path)
    updater.replace_sheet_data("Sheet1", [["new"]])
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        updated_sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        updated_strings = archive.read("xl/sharedStrings.xml").decode("utf-8")
    assert "<x:row" in updated_sheet
    assert "<x:c" in updated_sheet
    assert "<x:v" in updated_sheet
    assert "<x:si>" in updated_strings
    assert "<x:t>new</x:t>" in updated_strings


@pytest.mark.parametrize("binary", [False, True])
def test_updater_recalculates_shared_string_reference_count(binary, tmp_path):
    path = tmp_path / ("count.xlsb" if binary else "count.xlsx")
    if binary:
        _write_xlsb(path, [("Data", [["old"]]), ("Keep", [["keep"]])])
        updater = XlsbUpdater(path)
    else:
        _write_xlsx(path, [("Data", [["old"]]), ("Keep", [["keep"]])])
        updater = XlsxUpdater(path)

    updater.replace_sheet_data("Data", [])
    path.write_bytes(updater.to_bytes())

    member = "xl/sharedStrings.bin" if binary else "xl/sharedStrings.xml"
    with zipfile.ZipFile(path) as archive:
        data = archive.read(member)
    if binary:
        header = next(record for record in iter_records(data) if record.record_id == 0x9F)
        assert struct.unpack_from("<I", data, header.data_start)[0] == 1
    else:
        assert 'count="1"' in data.decode("utf-8")


@pytest.mark.parametrize("binary", [False, True])
def test_updater_uses_first_row_style_as_data_without_headers(binary, tmp_path):
    path = tmp_path / ("stylefirst.xlsb" if binary else "stylefirst.xlsx")
    if binary:
        _write_xlsb(path, [("Sheet1", [[(1, F.TWO_DECIMALS)]])])
        updater = XlsbUpdater(path)
    else:
        _write_xlsx(path, [("Sheet1", [[(1, F.TWO_DECIMALS)]])])
        updater = XlsxUpdater(path)

    updater.replace_sheet_data("Sheet1", [[2.5]])
    path.write_bytes(updater.to_bytes())

    if binary:
        sheet = _xlsb_sheet_bytes(path)
        styles = [
            struct.unpack_from("<I", sheet, record.data_start + 4)[0] & 0xFFFFFF
            for record in iter_records(sheet)
            if record.record_id == 0x05
        ]
        assert styles == [4]
    else:
        assert 's="' in _xlsx_sheet_xml(path)


def test_xlsb_updater_replaces_trailing_row_headers(tmp_path):
    path = tmp_path / "trailing.xlsb"
    _write_xlsb(path, [("Sheet1", [[1]])])
    _set_xlsb_trailing_row_header(path)

    updater = XlsbUpdater(path)
    updater.replace_sheet_data("Sheet1", [[2]])
    path.write_bytes(updater.to_bytes())

    sheet = _xlsb_sheet_bytes(path)
    assert sum(record.record_id == 0x00 for record in iter_records(sheet)) == 1


def test_xlsb_updater_only_refreshes_matching_pivot_cache(tmp_path):
    path = tmp_path / "multiple-pivots.xlsb"
    _write_xlsb(path, [("Data", [["old"]]), ("Other", [["old"]])])
    _make_xlsb_pivot(path)
    other_name = "Other".encode("utf-16le")
    other_payload = (
        b"\x00\x00\x00"
        + struct.pack("<I", len("Other"))
        + other_name
        + struct.pack("<iiii", 0, 1, 0, 0)
    )
    _add_zip_members(
        path,
        {
            "xl/pivotCache/pivotCacheDefinition2.bin": build_record(
                0xB3, b"\x00\x00\x00\x00"
            )
            + build_record(0xBB, other_payload)
        },
    )

    updater = XlsbUpdater(path)
    updater.replace_sheet_data("Data", [["new"]])
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        other_cache = archive.read("xl/pivotCache/pivotCacheDefinition2.bin")
    other_header = next(
        record for record in iter_records(other_cache) if record.record_id == 0xB3
    )
    assert other_cache[other_header.data_start + 3] & 0x04 == 0


def test_xlsb_updater_replaces_data_and_preserves_binary_wrappers(tmp_path):
    source = tmp_path / "source.xlsb"
    output = tmp_path / "output.xlsb"
    _write_xlsb(source, [("Data", [["old", 1], ["remove", 2]]), ("Keep", [["yes", 8]])])
    before = _xlsb_sheet_bytes(source)

    updater = XlsbUpdater(source)
    assert updater.get_sheet_names() == ["Data", "Keep"]
    updater.replace_sheet_data(
        "Data",
        [["new", 42], [None, True], ["middle", None], [None, None]],
        headers=["Name", "Value"],
    )
    output.write_bytes(updater.to_bytes())

    assert _read_sheet(output, "Data") == [
        ["Name", "Value"],
        ["new", 42],
        [None, True],
        ["middle", None],
    ]
    assert _read_sheet(output, "Keep") == [["yes", 8]]
    after = _xlsb_sheet_bytes(output)
    assert len(after) != len(before)
    assert any(record.record_id in (0x92, 0x217, 0x1DD, 0x1DC) for record in iter_records(after))


def test_xlsb_updater_supports_dates_styles_and_shared_strings(tmp_path):
    path = tmp_path / "styles.xlsb"
    _write_xlsb(
        path,
        [("Sheet1", [["Date", "Amount"], [(datetime.date(2024, 1, 2), F.DATE_SHORT), (1.25, F.TWO_DECIMALS)]])],
    )
    updater = XlsbUpdater(path)
    updater.replace_sheet_data(
        "Sheet1",
        [[datetime.date(2025, 2, 3), 2.5]],
        headers=["Date", "Amount"],
    )
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.bin")
        cell_styles = []
        for record in iter_records(sheet):
            if 0x01 <= record.record_id <= 0x0B and record.length >= 8:
                cell_styles.append(struct.unpack_from("<I", sheet, record.data_start + 4)[0] & 0xFFFFFF)
        assert 3 in cell_styles
        assert any(style in (1, 4) for style in cell_styles)
    assert isinstance(_read_sheet(path, "Sheet1")[1][0], (datetime.date, datetime.datetime))


def test_xlsb_updater_updates_pivot_cache_source_and_refresh_flag(tmp_path):
    path = tmp_path / "pivot.xlsb"
    _write_xlsb(path, [("Data", [["old", 1], ["old2", 2]])])
    _make_xlsb_pivot(path)

    updater = XlsbUpdater(path)
    updater.replace_sheet_data("Data", [["one", 1], ["two", 2], ["three", 3]])
    path.write_bytes(updater.to_bytes())

    with zipfile.ZipFile(path) as archive:
        cache = archive.read("xl/pivotCache/pivotCacheDefinition1.bin")
    refresh_flags = []
    source_ranges = []
    for record in iter_records(cache):
        payload = cache[record.data_start : record.data_end]
        if record.record_id == 0xB3:
            refresh_flags.append(payload[3] & 0x04)
        elif record.record_id == 0xBB:
            count = struct.unpack_from("<I", payload, 3)[0]
            name, name_end = read_utf16(payload, 7, count)
            if name == "Data":
                source_ranges.append(struct.unpack_from("<iiii", payload, name_end))
    assert refresh_flags == [4]
    assert source_ranges == [(0, 2, 0, 1)]


@pytest.mark.parametrize("binary", [False, True])
def test_updater_can_clear_a_sheet(binary, tmp_path):
    path = tmp_path / ("clear.xlsb" if binary else "clear.xlsx")
    if binary:
        _write_xlsb(path, [("Sheet1", [[1], [2]])])
        updater = XlsbUpdater(path)
    else:
        _write_xlsx(path, [("Sheet1", [[1], [2]])])
        updater = XlsxUpdater(path)

    updater.replace_sheet_data("Sheet1", [])
    updater.save()
    assert _read_sheet(path, "Sheet1") == []


@pytest.mark.parametrize("binary", [False, True])
def test_updater_preserves_non_target_zip_payloads(binary, tmp_path):
    path = tmp_path / ("preserve.xlsb" if binary else "preserve.xlsx")
    if binary:
        _write_xlsb(path, [("Data", [["old", 1]]), ("Keep", [["keep", 2]])])
        updater = XlsbUpdater(path)
    else:
        _write_xlsx(path, [("Data", [["old", 1]]), ("Keep", [["keep", 2]])])
        updater = XlsxUpdater(path)

    with zipfile.ZipFile(path) as archive:
        before = {name: archive.read(name) for name in archive.namelist()}
    updater.replace_sheet_data("Data", [["new", 10]])
    output = tmp_path / ("preserve-output.xlsb" if binary else "preserve-output.xlsx")
    updater.save(output)
    with zipfile.ZipFile(output) as archive:
        after = {name: archive.read(name) for name in archive.namelist()}

    changed_parts = {"xl/worksheets/sheet1.bin" if binary else "xl/worksheets/sheet1.xml"}
    changed_parts.add("xl/sharedStrings.bin" if binary else "xl/sharedStrings.xml")
    for name, payload in before.items():
        if name not in changed_parts:
            assert after[name] == payload, name


@pytest.mark.parametrize("updater_type", [XlsxUpdater, XlsbUpdater])
def test_updater_errors_are_explicit(tmp_path, updater_type):
    path = tmp_path / ("bad.xlsx" if updater_type is XlsxUpdater else "bad.xlsb")
    path.write_bytes(b"not a zip")
    with pytest.raises(zipfile.BadZipFile):
        updater_type(path)


def test_updater_rejects_wrong_sheet_and_style_option(tmp_path):
    path = tmp_path / "errors.xlsx"
    _write_xlsx(path, [("Sheet1", [[1]])])
    updater = XlsxUpdater(path)
    with pytest.raises(KeyError):
        updater.replace_sheet_data("Missing", [[1]])
    with pytest.raises(ValueError):
        updater.replace_sheet_data("Sheet1", [[1]], style_fallback="bad")
