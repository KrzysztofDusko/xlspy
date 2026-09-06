import datetime
import shutil
import struct
import zipfile
from pathlib import Path

import pytest

from xlspy import (
    ExcelReader,
    F,
    XlsbUpdater,
    XlsbWriter,
    XlsmUpdater,
    XlsxUpdater,
    XlsxWriter,
)
from xlspy.biff_utils import build_record
from xlspy.updater_utils import ZipPackage


ROWS = [
    ["new & <value>", 42, True, datetime.date(2025, 2, 3)],
    [None, False, 2.5, datetime.datetime(2025, 2, 3, 10, 30)],
    ["middle", None, None, None],
    [None, None, None, None],
]
HEADERS = ["Name", "Value", "Active", "Date"]


def _write_workbook(path: Path, binary: bool, *, shared_strings: bool = True) -> None:
    if binary:
        with XlsbWriter(path) as writer:
            writer.add_sheet("Data")
            writer.write_sheet(
                [
                    ["old", 1, True, (datetime.date(2024, 1, 2), F.DATE_SHORT)],
                    ["remove", 2, False, (1.25, F.TWO_DECIMALS)],
                ]
            )
            writer.add_sheet("Keep")
            writer.write_sheet([["untouched", 9]])
    else:
        with XlsxWriter(path, useSharedStrings=shared_strings) as writer:
            writer.add_sheet("Data")
            writer.write_sheet(
                [
                    ["old", 1, True, (datetime.date(2024, 1, 2), F.DATE_SHORT)],
                    ["remove", 2, False, (1.25, F.TWO_DECIMALS)],
                ]
            )
            writer.add_sheet("Keep")
            writer.write_sheet([["untouched", 9]])


def _archive_payloads(path: Path) -> list[tuple[str, bytes]]:
    with zipfile.ZipFile(path) as archive:
        return [(info.filename, archive.read(info)) for info in archive.infolist()]


def _read_sheets(path: Path) -> dict[str, list[list[object]]]:
    with ExcelReader() as reader:
        reader.open(str(path))
        return {name: list(reader.get_rows(name)) for name in reader.get_sheet_names()}


def _add_complex_parts(path: Path, binary: bool) -> None:
    with zipfile.ZipFile(path, "a") as archive:
        if binary:
            sheet = archive.read("xl/worksheets/sheet1.bin")
            archive.writestr(
                "xl/worksheets/sheet1.bin",
                sheet + build_record(0xA1, b"\x00" * 16),
            )
            sheet_name = "Data".encode("utf-16le")
            source_payload = (
                b"\x00\x00\x00"
                + struct.pack("<I", len("Data"))
                + sheet_name
                + struct.pack("<iiii", 0, 1, 0, 1)
            )
            archive.writestr(
                "xl/pivotCache/pivotCacheDefinition1.bin",
                build_record(0xB3, b"\x00\x00\x00\x00")
                + build_record(0xBB, source_payload),
            )
        else:
            sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
            sheet = sheet.replace(
                "</worksheet>", '<autoFilter ref="A1:D2"/></worksheet>'
            )
            archive.writestr("xl/worksheets/sheet1.xml", sheet.encode("utf-8"))
            archive.writestr(
                "xl/pivotCache/pivotCacheDefinition1.xml",
                b'<pivotCacheDefinition recordCount="2">'
                b'<cacheSource type="worksheet">'
                b'<worksheetSource ref="A1:D2" sheet="Data"/>'
                b"</cacheSource></pivotCacheDefinition>",
            )
            archive.writestr(
                "xl/pivotTables/pivotTable1.xml",
                b'<pivotTableDefinition refreshOnLoad="0"/>',
            )


def _make_xlsm(path: Path) -> bytes:
    """Add a minimal macro project without interpreting its binary payload."""

    macro_payload = b"synthetic-vba-project\x00\xff\x10"
    entries = []
    with zipfile.ZipFile(path) as source:
        for info in source.infolist():
            payload = source.read(info)
            if info.filename == "[Content_Types].xml":
                payload = payload.replace(
                    b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    b"application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                )
                payload = payload.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/vbaProject.bin" '
                    b'ContentType="application/vnd.ms-office.vbaProject"/></Types>',
                )
            elif info.filename == "xl/_rels/workbook.xml.rels":
                payload = payload.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rIdVba" '
                    b'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                    b'Target="vbaProject.bin"/></Relationships>',
                )
            entries.append((info, payload))

    temporary = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(temporary, "w") as target:
        for info, payload in entries:
            target.writestr(info, payload)
        target.writestr("xl/vbaProject.bin", macro_payload)
    temporary.replace(path)
    return macro_payload


class _OneShotRows:
    def __init__(self, rows):
        self._rows = rows
        self._iterated = False

    def __iter__(self):
        if self._iterated:
            raise AssertionError("The database row source was iterated twice")
        self._iterated = True
        yield from self._rows


@pytest.mark.parametrize(
    ("binary", "shared_strings"),
    [(False, True), (False, False), (True, True)],
)
def test_streaming_output_matches_regular_output_byte_for_byte(
    tmp_path, binary, shared_strings
):
    extension = ".xlsb" if binary else ".xlsx"
    source = tmp_path / f"source{extension}"
    regular = tmp_path / f"regular{extension}"
    streamed = tmp_path / f"streamed{extension}"
    _write_workbook(source, binary, shared_strings=shared_strings)

    # An unrelated package member must also survive unchanged.
    with zipfile.ZipFile(source, "a") as archive:
        archive.writestr("custom/untouched.bin", b"payload\x00\xff")
    shutil.copyfile(source, regular)
    shutil.copyfile(source, streamed)

    updater_type = XlsbUpdater if binary else XlsxUpdater
    regular_updater = updater_type(regular)
    regular_updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    regular.write_bytes(regular_updater.to_bytes())

    streaming_updater = updater_type(streamed)
    streaming_updater.replace_sheet_data_stream(
        "Data", _OneShotRows(ROWS), headers=HEADERS
    )
    streaming_updater.save()

    assert _archive_payloads(streamed) == _archive_payloads(regular)
    assert _read_sheets(streamed) == _read_sheets(regular)


@pytest.mark.parametrize("binary", [False, True])
def test_streaming_matches_regular_with_filters_and_pivot_metadata(tmp_path, binary):
    extension = ".xlsb" if binary else ".xlsx"
    source = tmp_path / f"source{extension}"
    regular = tmp_path / f"regular{extension}"
    streamed = tmp_path / f"streamed{extension}"
    _write_workbook(source, binary)
    _add_complex_parts(source, binary)
    shutil.copyfile(source, regular)
    shutil.copyfile(source, streamed)

    updater_type = XlsbUpdater if binary else XlsxUpdater
    regular_updater = updater_type(regular)
    regular_updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    regular.write_bytes(regular_updater.to_bytes())

    streaming_updater = updater_type(streamed)
    streaming_updater.replace_sheet_data_stream(
        "Data", _OneShotRows(ROWS), headers=HEADERS
    )
    streaming_updater.save()

    assert _archive_payloads(streamed) == _archive_payloads(regular)


@pytest.mark.parametrize("binary", [False, True])
def test_streaming_supports_multiple_sheet_replacements(tmp_path, binary):
    extension = ".xlsb" if binary else ".xlsx"
    source = tmp_path / f"source{extension}"
    regular = tmp_path / f"regular{extension}"
    streamed = tmp_path / f"streamed{extension}"
    _write_workbook(source, binary)
    shutil.copyfile(source, regular)
    shutil.copyfile(source, streamed)

    updater_type = XlsbUpdater if binary else XlsxUpdater
    regular_updater = updater_type(regular)
    regular_updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    regular_updater.replace_sheet_data("Keep", [["changed", 10]])
    regular.write_bytes(regular_updater.to_bytes())

    streaming_updater = updater_type(streamed)
    streaming_updater.replace_sheet_data_stream(
        "Data", _OneShotRows(ROWS), headers=HEADERS
    )
    streaming_updater.replace_sheet_data_stream(
        "Keep", _OneShotRows([["changed", 10]])
    )
    streaming_updater.save()

    assert _archive_payloads(streamed) == _archive_payloads(regular)


@pytest.mark.parametrize("binary", [False, True])
def test_streaming_replacing_same_sheet_again_matches_regular(tmp_path, binary):
    extension = ".xlsb" if binary else ".xlsx"
    source = tmp_path / f"source{extension}"
    regular = tmp_path / f"regular{extension}"
    streamed = tmp_path / f"streamed{extension}"
    _write_workbook(source, binary)
    shutil.copyfile(source, regular)
    shutil.copyfile(source, streamed)

    updater_type = XlsbUpdater if binary else XlsxUpdater
    regular_updater = updater_type(regular)
    regular_updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    regular_updater.replace_sheet_data("Data", [["final", 7]], headers=HEADERS)
    regular.write_bytes(regular_updater.to_bytes())

    streaming_updater = updater_type(streamed)
    streaming_updater.replace_sheet_data_stream(
        "Data", _OneShotRows(ROWS), headers=HEADERS
    )
    streaming_updater.replace_sheet_data_stream(
        "Data", _OneShotRows([["final", 7]]), headers=HEADERS
    )
    streaming_updater.save()

    assert _archive_payloads(streamed) == _archive_payloads(regular)


@pytest.mark.parametrize("binary", [False, True])
def test_streaming_save_does_not_materialize_rows_or_archive(
    tmp_path, binary, monkeypatch
):
    extension = ".xlsb" if binary else ".xlsx"
    path = tmp_path / f"streaming{extension}"
    _write_workbook(path, binary)
    updater_type = XlsbUpdater if binary else XlsxUpdater
    module_name = "xlspy.xlsb_updater" if binary else "xlspy.xlsx_updater"
    updater_module = __import__(module_name, fromlist=["materialize_rows"])

    def fail_materialization(_):
        raise AssertionError("streaming updater materialized all rows")

    def fail_to_bytes(_):
        raise AssertionError("streaming save materialized the complete ZIP")

    monkeypatch.setattr(updater_module, "materialize_rows", fail_materialization)
    monkeypatch.setattr(ZipPackage, "to_bytes", fail_to_bytes)

    updater = updater_type(path)
    original_get = updater._package.get
    shared_strings_member = "xl/sharedStrings.bin" if binary else "xl/sharedStrings.xml"

    def fail_shared_strings_get(name):
        if name == shared_strings_member:
            raise AssertionError("streaming updater loaded the full shared-string part")
        return original_get(name)

    monkeypatch.setattr(updater._package, "get", fail_shared_strings_get)
    updater.replace_sheet_data_stream("Data", _OneShotRows(ROWS), headers=HEADERS)
    updater.save()


@pytest.mark.parametrize("binary", [False, True])
def test_streaming_source_is_unchanged_when_generator_fails(tmp_path, binary):
    extension = ".xlsb" if binary else ".xlsx"
    path = tmp_path / f"source{extension}"
    _write_workbook(path, binary)
    before = path.read_bytes()

    def broken_rows():
        yield ["written-before-error", 1]
        raise RuntimeError("database failure")

    updater_type = XlsbUpdater if binary else XlsxUpdater
    updater = updater_type(path)
    with pytest.raises(RuntimeError, match="database failure"):
        updater.replace_sheet_data_stream("Data", broken_rows())

    assert path.read_bytes() == before


@pytest.mark.parametrize("binary", [False, True])
def test_streaming_failure_rolls_back_shared_strings_for_reuse(tmp_path, binary):
    extension = ".xlsb" if binary else ".xlsx"
    source = tmp_path / f"source{extension}"
    expected = tmp_path / f"expected{extension}"
    _write_workbook(source, binary)
    shutil.copyfile(source, expected)

    def broken_rows():
        yield ["failed-value-must-not-survive", 123]
        raise RuntimeError("database failure")

    updater_type = XlsbUpdater if binary else XlsxUpdater
    updater = updater_type(source)
    with pytest.raises(RuntimeError, match="database failure"):
        updater.replace_sheet_data_stream("Data", broken_rows())
    updater.replace_sheet_data_stream("Data", _OneShotRows(ROWS), headers=HEADERS)
    updater.save()

    expected_updater = updater_type(expected)
    expected_updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    expected.write_bytes(expected_updater.to_bytes())

    assert _archive_payloads(source) == _archive_payloads(expected)


@pytest.mark.parametrize("binary", [False, True])
def test_failed_shared_string_staging_removes_temporary_output(tmp_path, binary, monkeypatch):
    extension = ".xlsb" if binary else ".xlsx"
    path = tmp_path / f"failed-staging{extension}"
    _write_workbook(path, binary)
    updater_type = XlsbUpdater if binary else XlsxUpdater
    updater = updater_type(path)
    original_set_file = updater._package.set_file
    shared_strings_path = "xl/sharedStrings.bin" if binary else "xl/sharedStrings.xml"

    def fail_shared_string_staging(name, replacement):
        if name == shared_strings_path:
            raise RuntimeError("shared-string staging failure")
        original_set_file(name, replacement)

    monkeypatch.setattr(updater._package, "set_file", fail_shared_string_staging)

    with pytest.raises(RuntimeError, match="shared-string staging failure"):
        updater.replace_sheet_data_stream("Data", _OneShotRows(ROWS), headers=HEADERS)

    assert not list(updater._package.temporary_directory.glob("part-*.xml"))
    assert not list(updater._package.temporary_directory.glob("part-*.bin"))


@pytest.mark.parametrize("binary", [False, True])
def test_streaming_reuse_after_in_place_save_refreshes_zip_offsets(tmp_path, binary):
    extension = ".xlsb" if binary else ".xlsx"
    path = tmp_path / f"reused{extension}"
    _write_workbook(path, binary)

    def large_rows():
        for index in range(80):
            yield [f"resized-{index}-" + ("x" * 300), index]

    updater_type = XlsbUpdater if binary else XlsxUpdater
    updater = updater_type(path)
    updater.replace_sheet_data_stream("Data", large_rows(), headers=HEADERS)
    updater.save()
    updater.replace_sheet_data_stream("Keep", _OneShotRows([["after-save", 11]]))
    updater.save()

    sheets = _read_sheets(path)
    assert sheets["Keep"] == [["after-save", 11]]
    assert sheets["Data"][0] == HEADERS
    assert sheets["Data"][1][0].startswith("resized-0-")


@pytest.mark.parametrize("binary", [False, True])
def test_repeated_streaming_replacements_remove_superseded_staging(tmp_path, binary):
    extension = ".xlsb" if binary else ".xlsx"
    path = tmp_path / f"staging{extension}"
    _write_workbook(path, binary)
    updater_type = XlsbUpdater if binary else XlsxUpdater
    updater = updater_type(path)

    updater.replace_sheet_data_stream("Data", _OneShotRows(ROWS), headers=HEADERS)
    sheet_path = "xl/worksheets/sheet1.bin" if binary else "xl/worksheets/sheet1.xml"
    strings_path = "xl/sharedStrings.bin" if binary else "xl/sharedStrings.xml"
    sheet_index = updater._package._indices[sheet_path]
    strings_index = updater._package._indices[strings_path]
    first_sheet = updater._package._overrides[sheet_index]
    first_strings = updater._package._overrides[strings_index]

    updater.replace_sheet_data_stream(
        "Data", _OneShotRows([["second replacement", 2]]), headers=HEADERS
    )

    assert isinstance(first_sheet, Path)
    assert isinstance(first_strings, Path)
    assert not first_sheet.exists()
    assert not first_strings.exists()
    staged_parts = [
        item
        for item in updater._package.temporary_directory.iterdir()
        if item.name.startswith("part-")
    ]
    assert len(staged_parts) == 2


def test_xlsm_updates_preserve_vba_project_bytes(tmp_path):
    source = tmp_path / "source.xlsm"
    regular = tmp_path / "regular.xlsm"
    streamed = tmp_path / "streamed.xlsm"
    _write_workbook(source, False)
    macro_payload = _make_xlsm(source)
    shutil.copyfile(source, regular)
    shutil.copyfile(source, streamed)

    regular_updater = XlsmUpdater(regular)
    regular_updater.replace_sheet_data("Data", ROWS, headers=HEADERS)
    regular_updater.save()

    streaming_updater = XlsmUpdater(streamed)
    streaming_updater.replace_sheet_data_stream(
        "Data", _OneShotRows(ROWS), headers=HEADERS
    )
    streaming_updater.save()

    assert _archive_payloads(streamed) == _archive_payloads(regular)
    for path in (regular, streamed):
        with zipfile.ZipFile(path) as archive:
            assert archive.read("xl/vbaProject.bin") == macro_payload
            assert (
                archive.read("[Content_Types].xml").count(
                    b"application/vnd.ms-excel.sheet.macroEnabled.main+xml"
                )
                == 1
            )
