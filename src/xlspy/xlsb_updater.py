"""Update worksheet data in an existing XLSB package."""

from __future__ import annotations

import datetime as _datetime
import math
import os
import re
import struct
from decimal import Decimal
from typing import Any, Iterable, Literal, Sequence

from .biff_utils import (
    Biff12Record,
    build_record,
    iter_records,
    iter_records_stream,
    parse_shared_strings_bin,
    read_utf16,
)
from .updater_utils import (
    DiskStringIndex,
    ZipPackage,
    column_index_to_letter,
    copy_file_range,
    copy_stream,
    local_xml_name,
    materialize_rows,
    parse_relationships_xml,
    resolve_zip_target,
    strip_invalid_xml_characters,
    trim_trailing_empty_rows,
    unwrap_cell,
)


StyleFallback = Literal["inherit", "general"]
_CELL_RECORDS = set(range(0x01, 0x0C))
_RK_INT_LOWER = -536870912
_RK_INT_UPPER = 536870911


def _date_like_format(format_code: str) -> bool:
    """Return whether an Excel number format is date-like."""

    simplified = re.sub(r'"(?:[^"]|"")*"', "", format_code)
    simplified = re.sub(r"\\.", "", simplified).lower()
    simplified = re.sub(r"\[[^\]]*\]", "", simplified)
    return bool(re.search(r"[dy]", simplified) or re.search(r"m", simplified))


def _xml_attribute(attrs: dict[str, str], wanted: str) -> str | None:
    for name, value in attrs.items():
        if local_xml_name(name) == wanted:
            return value
    return None


class XlsbUpdater:
    """Replace the data region of sheets in an existing XLSB file."""

    def __init__(self, path: str | os.PathLike[str]):
        self._package = ZipPackage(path)
        if self._package.has("xl/workbook.xml"):
            raise ValueError("The input is XLSX; use XlsxUpdater")
        if not self._package.has("xl/workbook.bin"):
            raise ValueError("XLSB workbook.bin was not found")

        self._date_1904 = self._load_date_system()
        self._sheets = self._load_sheet_structure()
        self._shared_strings_loaded = False
        self._shared_strings_data: bytes | None = None
        self._shared_strings_values: list[str] = []
        self._shared_string_indices: dict[str, int] = {}
        self._shared_strings_total = 0
        self._shared_strings_pending: list[str] = []
        self._shared_strings_dirty = False
        self._stream_shared_strings: DiskStringIndex | None = None

    def get_sheet_names(self) -> list[str]:
        """Return worksheet names in workbook order."""

        return list(self._sheets)

    def replace_sheet_data(
        self,
        sheet_name: str,
        rows: Iterable[Sequence[Any]],
        *,
        headers: Sequence[str] | None = None,
        style_fallback: StyleFallback = "inherit",
    ) -> None:
        """Replace one worksheet's used data region and preserve sheet wrappers."""

        if style_fallback not in ("inherit", "general"):
            raise ValueError("style_fallback must be 'inherit' or 'general'")
        try:
            sheet_path = self._sheets[sheet_name]
        except KeyError as exc:
            raise KeyError(f"Worksheet not found: {sheet_name}") from exc

        if self._stream_shared_strings is not None:
            self.replace_sheet_data_stream(
                sheet_name,
                rows,
                headers=headers,
                style_fallback=style_fallback,
            )
            return

        materialized = trim_trailing_empty_rows(materialize_rows(rows))
        header_values = list(headers) if headers is not None else None
        sheet_data = self._package.get(sheet_path)
        self._ensure_shared_strings_loaded()
        if self._shared_strings_data is not None:
            self._shared_strings_total = max(
                0,
                self._shared_strings_total
                - self._shared_string_ref_count(sheet_data),
            )
            self._shared_strings_dirty = True

        header_styles, data_styles = self._collect_styles(
            sheet_data, headers_provided=headers is not None
        )
        date_xf = self._find_date_xf() if style_fallback == "inherit" else None
        width = max(
            [len(row) for row in materialized]
            + ([len(header_values)] if header_values is not None else [])
            + [0]
        )
        total_rows = len(materialized) + (1 if header_values is not None else 0)
        last_row = max(0, total_rows - 1)
        last_col = max(0, width - 1)
        new_rows = self._build_rows(
            materialized,
            header_values,
            data_styles,
            header_styles,
            date_xf,
            style_fallback,
        )

        region = self._find_rows_region(sheet_data)
        old_rows_length = region.rows_end - region.rows_start
        updated = bytearray(
            sheet_data[: region.rows_start]
            + new_rows
            + sheet_data[region.rows_end :]
        )
        delta = len(new_rows) - old_rows_length

        for filter_offset in region.auto_filter_ranges:
            new_offset = filter_offset
            if filter_offset >= region.rows_end:
                new_offset += delta
            if new_offset + 16 <= len(updated):
                struct.pack_into("<i", updated, new_offset + 4, last_row)
                struct.pack_into("<i", updated, new_offset + 12, last_col)

        self._patch_dimension(updated, last_row, last_col)
        self._package.set(sheet_path, bytes(updated))
        self._commit_shared_strings()
        self._patch_pivot_caches(sheet_name, last_row, last_col)

    def replace_sheet_data_stream(
        self,
        sheet_name: str,
        rows: Iterable[Sequence[Any]],
        *,
        headers: Sequence[str] | None = None,
        style_fallback: StyleFallback = "inherit",
    ) -> None:
        """Replace worksheet data atomically while streaming rows.

        The package and disk-backed shared-string index are committed only
        after the one-pass row source and all metadata updates succeed.  A
        failure therefore leaves this updater reusable for a later attempt.
        """

        if style_fallback not in ("inherit", "general"):
            raise ValueError("style_fallback must be 'inherit' or 'general'")
        if sheet_name not in self._sheets:
            raise KeyError(f"Worksheet not found: {sheet_name}")

        self._ensure_stream_shared_strings()
        store = self._stream_shared_strings
        self._package.begin_transaction()
        try:
            if store is not None:
                store.begin_transaction()
            self._replace_sheet_data_stream_impl(
                sheet_name,
                rows,
                headers=headers,
                style_fallback=style_fallback,
            )
            self._package.commit_transaction()
            if store is not None:
                store.commit_transaction()
        except BaseException:
            if store is not None:
                store.rollback_transaction()
            self._package.rollback_transaction()
            raise

    def _replace_sheet_data_stream_impl(
        self,
        sheet_name: str,
        rows: Iterable[Sequence[Any]],
        *,
        headers: Sequence[str] | None = None,
        style_fallback: StyleFallback = "inherit",
    ) -> None:
        """Replace worksheet data while streaming rows through temporary files."""

        if style_fallback not in ("inherit", "general"):
            raise ValueError("style_fallback must be 'inherit' or 'general'")
        try:
            sheet_path = self._sheets[sheet_name]
        except KeyError as exc:
            raise KeyError(f"Worksheet not found: {sheet_name}") from exc

        header_values = list(headers) if headers is not None else None
        original_sheet = self._package.temporary_path(suffix=".bin")
        with self._package.open_entry(sheet_path) as source, original_sheet.open("wb") as target:
            copy_stream(source, target)

        self._ensure_stream_shared_strings()
        header_styles, data_styles = self._collect_styles_stream(
            original_sheet, headers_provided=headers is not None
        )
        if self._stream_shared_strings is not None:
            self._shared_strings_total = max(
                0,
                self._stream_shared_strings.total_count
                - self._shared_string_ref_count_stream(original_sheet),
            )
            self._stream_shared_strings.total_count = self._shared_strings_total
            self._stream_shared_strings.dirty = True
        date_xf = self._find_date_xf() if style_fallback == "inherit" else None

        rows_file = self._package.temporary_path(suffix=".rows")
        width = len(header_values) if header_values is not None else 0
        data_rows_seen = 0
        kept_data_rows = 0
        last_kept_end = 0
        pending_empty_width = 0

        try:
            with rows_file.open("wb") as output:
                next_row = 0

                def write_row(row_number: int, row_cells: list[bytes]) -> None:
                    payload = bytearray(25)
                    struct.pack_into("<i", payload, 0, row_number)
                    struct.pack_into("<i", payload, 8, 300)
                    payload[13] = 1
                    struct.pack_into("<i", payload, 17, 0)
                    # The final column is not known until the generator ends.
                    struct.pack_into("<i", payload, 21, 0)
                    output.write(build_record(0x00, bytes(payload)))
                    for cell in row_cells:
                        output.write(cell)

                if header_values is not None:
                    header_texts = [
                        "" if unwrap_cell(value) is None else str(unwrap_cell(value))
                        for value in header_values
                    ]
                    write_row(
                        next_row,
                        [
                            self._string_cell_bytes(
                                column,
                                0 if style_fallback == "general" else header_styles.get(column, 0),
                                header_text,
                            )
                            for column, header_text in enumerate(header_texts)
                        ],
                    )
                    last_kept_end = output.tell()
                    next_row += 1

                for source_row in rows:
                    if isinstance(source_row, (str, bytes, bytearray)):
                        raise TypeError("Each row must be a sequence of cells, not text")
                    try:
                        row = list(source_row)
                    except TypeError as exc:
                        raise TypeError("Each row must be an iterable of cells") from exc

                    cells = [
                        self._value_cell_bytes(
                            raw,
                            column,
                            data_styles,
                            date_xf,
                            style_fallback,
                        )
                        for column, raw in enumerate(row)
                        if raw is not None
                    ]
                    write_row(next_row, cells)
                    data_rows_seen += 1
                    next_row += 1

                    if all(cell is None for cell in row):
                        pending_empty_width = max(pending_empty_width, len(row))
                    else:
                        width = max(width, pending_empty_width, len(row))
                        pending_empty_width = 0
                        kept_data_rows = data_rows_seen
                        last_kept_end = output.tell()

                output.truncate(last_kept_end)

            last_col = max(0, width - 1)
            self._patch_row_headers_file(rows_file, last_col)
            region = self._find_rows_region_stream(original_sheet)
            updated_sheet = self._package.temporary_path(suffix=".bin")
            with updated_sheet.open("wb") as output:
                copy_file_range(original_sheet, output, 0, region.rows_start)
                with rows_file.open("rb") as generated_rows:
                    copy_stream(generated_rows, output)
                copy_file_range(
                    original_sheet,
                    output,
                    region.rows_end,
                    original_sheet.stat().st_size,
                )

            total_rows = kept_data_rows + (1 if header_values is not None else 0)
            last_row = max(0, total_rows - 1)
            self._patch_sheet_file(updated_sheet, last_row, last_col)
            self._package.set_file(sheet_path, updated_sheet)
            self._commit_stream_shared_strings()
            self._patch_pivot_caches(sheet_name, last_row, last_col)
        finally:
            for temporary_path in (original_sheet, rows_file):
                try:
                    temporary_path.unlink()
                except FileNotFoundError:
                    pass

    def save(self, output_path: str | os.PathLike[str] | None = None) -> None:
        """Save to a new path or atomically replace the input file."""

        self._package.save(output_path)

    def to_bytes(self) -> bytes:
        """Return the current package as XLSB bytes."""

        return self._package.to_bytes()

    def _load_sheet_structure(self) -> dict[str, str]:
        rels_path = "xl/_rels/workbook.bin.rels"
        if not self._package.has(rels_path):
            raise ValueError("XLSB workbook relationships were not found")
        relationships = parse_relationships_xml(self._package.get(rels_path))
        workbook_data = self._package.get("xl/workbook.bin")
        sheets: dict[str, str] = {}

        for record in iter_records(workbook_data):
            if record.record_id != 0x9C:
                continue
            payload = workbook_data[record.data_start : record.data_end]
            if len(payload) < 12:
                raise ValueError("Malformed BrtBundleSh record")
            relationship_length = struct.unpack_from("<I", payload, 8)[0]
            relationship_start = 12
            relationship_id, name_length_offset = read_utf16(
                payload, relationship_start, relationship_length
            )
            if name_length_offset + 4 > len(payload):
                raise ValueError("Malformed BrtBundleSh sheet name")
            name_length = struct.unpack_from("<I", payload, name_length_offset)[0]
            sheet_name, _ = read_utf16(
                payload, name_length_offset + 4, name_length
            )
            target = relationships.get(relationship_id)
            if target is None:
                raise ValueError(f"Relationship {relationship_id!r} for sheet is missing")
            sheet_path = resolve_zip_target(target, "xl")
            if not self._package.has(sheet_path):
                raise ValueError(f"Worksheet part was not found: {sheet_path}")
            sheets[sheet_name] = sheet_path

        if not sheets:
            raise ValueError("No worksheets were found in workbook.bin")
        return sheets

    def _load_date_system(self) -> bool:
        for record in iter_records(self._package.get("xl/workbook.bin")):
            if record.record_id == 0x99 and record.length >= 4:
                payload = self._package.get("xl/workbook.bin")[
                    record.data_start : record.data_end
                ]
                return bool(struct.unpack_from("<I", payload, 0)[0] & 0x01)
        return False

    def _ensure_shared_strings_loaded(self) -> None:
        if self._shared_strings_loaded:
            return
        self._shared_strings_loaded = True
        if not self._package.has("xl/sharedStrings.bin"):
            return
        data = self._package.get("xl/sharedStrings.bin")
        parsed = parse_shared_strings_bin(data)
        self._shared_strings_data = data
        self._shared_strings_values = parsed.values
        self._shared_string_indices = {
            value: index for index, value in enumerate(parsed.values)
        }
        self._shared_strings_total = parsed.total_count

    def _ensure_stream_shared_strings(self) -> None:
        if self._stream_shared_strings is not None:
            return
        if not self._package.has("xl/sharedStrings.bin"):
            return

        store = DiskStringIndex(self._package.temporary_directory)
        total_count: int | None = None
        value_count = 0
        with self._package.open_entry("xl/sharedStrings.bin") as source:
            for record, payload in iter_records_stream(source):
                if record.record_id == 0x9F:
                    if len(payload) < 8:
                        raise ValueError("Malformed XLSB shared-string header")
                    total_count = struct.unpack_from("<I", payload, 0)[0]
                elif record.record_id == 0x13:
                    if len(payload) < 5:
                        raise ValueError("Malformed XLSB shared-string item")
                    character_count = struct.unpack_from("<I", payload, 1)[0]
                    value, string_end = read_utf16(payload, 5, character_count)
                    if string_end > len(payload):
                        raise ValueError("Malformed XLSB shared-string item length")
                    store.add_existing(value, value_count)
                    value_count += 1

        store.total_count = total_count if total_count is not None else value_count
        self._stream_shared_strings = store

    def _commit_stream_shared_strings(self) -> None:
        store = self._stream_shared_strings
        if store is None or not store.dirty:
            return

        source_path = self._package.temporary_path(suffix=".bin")
        with self._package.open_entry("xl/sharedStrings.bin") as source, source_path.open("wb") as target:
            copy_stream(source, target)
        output_path = self._package.temporary_path(suffix=".bin")
        inserted = False
        output_staged = False
        try:
            with source_path.open("rb") as source, output_path.open("wb") as target:
                for record, payload in iter_records_stream(source):
                    if record.record_id == 0xA0 and not inserted:
                        for value in store.pending_values():
                            encoded = value.encode("utf-16le")
                            target.write(
                                build_record(
                                    0x13,
                                    b"\x00" + struct.pack("<I", len(encoded) // 2) + encoded,
                                )
                            )
                        inserted = True

                    if record.record_id == 0x9F and len(payload) >= 8:
                        header_length = record.data_start - record.header_start
                        source.seek(record.header_start)
                        target.write(source.read(header_length))
                        updated_payload = bytearray(payload)
                        struct.pack_into(
                            "<II",
                            updated_payload,
                            0,
                            store.total_count,
                            store.unique_count,
                        )
                        target.write(updated_payload)
                        source.seek(record.data_end)
                    else:
                        copy_file_range(
                            source_path,
                            target,
                            record.header_start,
                            record.data_end,
                        )

                if not inserted:
                    for value in store.pending_values():
                        encoded = value.encode("utf-16le")
                        target.write(
                            build_record(
                                0x13,
                                b"\x00" + struct.pack("<I", len(encoded) // 2) + encoded,
                            )
                        )
            self._package.set_file("xl/sharedStrings.bin", output_path)
            output_staged = True
            store.clear_pending()
        finally:
            try:
                source_path.unlink()
            except FileNotFoundError:
                pass
            if not output_staged:
                try:
                    output_path.unlink()
                except FileNotFoundError:
                    pass

    def _collect_styles_stream(
        self, sheet_path: str | os.PathLike[str], *, headers_provided: bool
    ) -> tuple[dict[int, int], dict[int, int]]:
        header_styles: dict[int, int] = {}
        data_counts: dict[int, dict[int, int]] = {}
        row = -1
        with open(sheet_path, "rb") as stream:
            for record, payload in iter_records_stream(stream):
                if record.record_id == 0x00 and record.length >= 4:
                    row = struct.unpack_from("<i", payload, 0)[0]
                    continue
                if record.record_id not in _CELL_RECORDS or record.length < 8:
                    continue
                column, style = struct.unpack_from("<II", payload, 0)
                style &= 0xFFFFFF
                if style <= 0:
                    continue
                if headers_provided and row == 0:
                    header_styles.setdefault(column, style)
                elif row >= 0:
                    counts = data_counts.setdefault(column, {})
                    counts[style] = counts.get(style, 0) + 1

        data_styles = {
            column: max(counts.items(), key=lambda item: (item[1], -item[0]))[0]
            for column, counts in data_counts.items()
        }
        return header_styles, data_styles

    @staticmethod
    def _shared_string_ref_count_stream(sheet_path: str | os.PathLike[str]) -> int:
        with open(sheet_path, "rb") as stream:
            return sum(
                1
                for record, _ in iter_records_stream(stream)
                if record.record_id == 0x07
            )

    def _commit_shared_strings(self) -> None:
        if self._shared_strings_data is None or (
            not self._shared_strings_pending and not self._shared_strings_dirty
        ):
            return

        new_items = bytearray()
        for value in self._shared_strings_pending:
            encoded = value.encode("utf-16le")
            payload = b"\x00" + struct.pack("<I", len(encoded) // 2) + encoded
            new_items.extend(build_record(0x13, payload))

        parsed = parse_shared_strings_bin(self._shared_strings_data)
        insertion = parsed.end_offset if parsed.end_offset is not None else len(self._shared_strings_data)
        updated = bytearray(
            self._shared_strings_data[:insertion]
            + bytes(new_items)
            + self._shared_strings_data[insertion:]
        )
        for record in iter_records(bytes(updated)):
            if record.record_id == 0x9F and record.length >= 8:
                struct.pack_into("<II", updated, record.data_start, self._shared_strings_total, len(self._shared_strings_values))
                break

        self._shared_strings_data = bytes(updated)
        self._shared_strings_pending.clear()
        self._shared_strings_dirty = False
        self._package.set("xl/sharedStrings.bin", self._shared_strings_data)

    def _shared_string_index(self, value: str) -> int:
        if self._stream_shared_strings is not None:
            return self._stream_shared_strings.get_or_add(value)
        index = self._shared_string_indices.get(value)
        if index is None:
            index = len(self._shared_strings_values)
            self._shared_strings_values.append(value)
            self._shared_string_indices[value] = index
            self._shared_strings_pending.append(value)
        self._shared_strings_total += 1
        self._shared_strings_dirty = True
        return index

    @staticmethod
    def _shared_string_ref_count(sheet_data: bytes) -> int:
        return sum(1 for record in iter_records(sheet_data) if record.record_id == 0x07)

    def _collect_styles(
        self, sheet_data: bytes, *, headers_provided: bool
    ) -> tuple[dict[int, int], dict[int, int]]:
        header_styles: dict[int, int] = {}
        data_counts: dict[int, dict[int, int]] = {}
        row = -1
        for record in iter_records(sheet_data):
            payload = sheet_data[record.data_start : record.data_end]
            if record.record_id == 0x00 and record.length >= 4:
                row = struct.unpack_from("<i", payload, 0)[0]
                continue
            if record.record_id not in _CELL_RECORDS or record.length < 8:
                continue
            column, style = struct.unpack_from("<II", payload, 0)
            style &= 0xFFFFFF
            if style <= 0:
                continue
            if headers_provided and row == 0:
                header_styles.setdefault(column, style)
            elif row >= 0:
                counts = data_counts.setdefault(column, {})
                counts[style] = counts.get(style, 0) + 1

        data_styles: dict[int, int] = {}
        for column, counts in data_counts.items():
            data_styles[column] = max(counts.items(), key=lambda item: (item[1], -item[0]))[0]
        return header_styles, data_styles

    def _find_date_xf(self) -> int | None:
        if not self._package.has("xl/styles.bin"):
            return None
        styles_data = self._package.get("xl/styles.bin")
        in_cell_xfs = False
        xf_index = 0
        custom_formats: dict[int, str] = {}
        in_number_formats = False
        for record in iter_records(styles_data):
            payload = styles_data[record.data_start : record.data_end]
            if record.record_id == 0x267:
                in_number_formats = True
                continue
            if record.record_id == 0x268:
                in_number_formats = False
                continue
            if in_number_formats and record.record_id == 0x2C and record.length >= 6:
                number_format_id = struct.unpack_from("<H", payload, 0)[0]
                character_count = struct.unpack_from("<I", payload, 2)[0]
                format_end = 6 + character_count * 2
                if format_end <= record.length:
                    custom_formats[number_format_id] = payload[6:format_end].decode(
                        "utf-16le"
                    )
                continue
            if record.record_id == 0x269:
                in_cell_xfs = True
                xf_index = 0
                continue
            if record.record_id == 0x26A:
                in_cell_xfs = False
                continue
            if in_cell_xfs and record.record_id == 0x2F and record.length >= 4:
                number_format_id = struct.unpack_from("<H", payload, 2)[0]
                if (14 <= number_format_id <= 22) or (45 <= number_format_id <= 47):
                    return xf_index
                if _date_like_format(custom_formats.get(number_format_id, "")):
                    return xf_index
                xf_index += 1
        return None

    def _build_rows(
        self,
        rows: list[list[Any]],
        headers: list[Any] | None,
        data_styles: dict[int, int],
        header_styles: dict[int, int],
        date_xf: int | None,
        style_fallback: StyleFallback,
    ) -> bytes:
        chunks: list[bytes] = []
        max_col = max(
            [len(row) for row in rows] + ([len(headers)] if headers is not None else []) + [0]
        )
        col_last = max(0, max_col - 1)

        def push_row(row_number: int, cells: list[bytes]) -> None:
            payload = bytearray(25)
            struct.pack_into("<i", payload, 0, row_number)
            struct.pack_into("<i", payload, 8, 300)
            payload[13] = 1
            struct.pack_into("<i", payload, 17, 0)
            struct.pack_into("<i", payload, 21, col_last)
            chunks.append(build_record(0x00, bytes(payload)))
            chunks.extend(cells)

        if headers is not None:
            header_texts = [
                "" if unwrap_cell(value) is None else str(unwrap_cell(value))
                for value in headers
            ]
            cells = [
                self._string_cell_bytes(
                    column,
                    0 if style_fallback == "general" else header_styles.get(column, 0),
                    header_text,
                )
                for column, header_text in enumerate(header_texts)
            ]
            push_row(0, cells)

        row_number = 1 if headers is not None else 0
        for row in rows:
            cells: list[bytes] = []
            for column, raw in enumerate(row):
                if raw is None:
                    continue
                cells.append(
                    self._value_cell_bytes(
                        raw,
                        column,
                        data_styles,
                        date_xf,
                        style_fallback,
                    )
                )
            push_row(row_number, cells)
            row_number += 1
        return b"".join(chunks)

    def _value_cell_bytes(
        self,
        raw: Any,
        column: int,
        data_styles: dict[int, int],
        date_xf: int | None,
        style_fallback: StyleFallback,
    ) -> bytes:
        value = unwrap_cell(raw)
        style = 0 if style_fallback == "general" else data_styles.get(column, 0)
        if isinstance(value, bool):
            payload = struct.pack("<IIB", column, style, 1 if value else 0)
            return build_record(0x04, payload)
        if isinstance(value, int) and not isinstance(value, bool):
            if _RK_INT_LOWER <= value <= _RK_INT_UPPER:
                payload = struct.pack("<IIi", column, style, (value << 2) | 2)
                return build_record(0x02, payload)
            payload = struct.pack("<IId", column, style, float(value))
            return build_record(0x05, payload)
        if isinstance(value, (float, Decimal)):
            numeric_value = float(value)
            if math.isfinite(numeric_value):
                payload = struct.pack("<IId", column, style, numeric_value)
                return build_record(0x05, payload)
            return self._string_cell_bytes(column, style, str(value))
        if isinstance(value, (_datetime.datetime, _datetime.date)):
            serial = self._date_to_serial(value)
            if math.isfinite(serial):
                date_style = style if style > 0 else (date_xf or 0)
                payload = struct.pack("<IId", column, date_style, serial)
                return build_record(0x05, payload)
        return self._string_cell_bytes(column, style, "" if value is None else str(value))

    def _string_cell_bytes(self, column: int, style: int, text: str) -> bytes:
        text = strip_invalid_xml_characters(text)
        if self._shared_strings_data is None and self._stream_shared_strings is None:
            encoded = text.encode("utf-16le")
            payload = struct.pack("<III", column, style, len(encoded) // 2) + encoded
            return build_record(0x06, payload)
        index = self._shared_string_index(text)
        return build_record(0x07, struct.pack("<III", column, style, index))

    @staticmethod
    def _patch_row_headers_file(path: str | os.PathLike[str], last_col: int) -> None:
        with open(path, "r+b") as stream:
            for record, _ in iter_records_stream(stream):
                if record.record_id == 0x00 and record.length >= 25:
                    stream.seek(record.data_start + 21)
                    stream.write(struct.pack("<i", last_col))
                    stream.seek(record.data_end)

    @staticmethod
    def _find_rows_region_stream(sheet_path: str | os.PathLike[str]) -> "_RowsRegion":
        rows_start = -1
        last_cell_end = -1
        last_row_header_end = -1
        last_wrapper = -1
        with open(sheet_path, "rb") as stream:
            for record, _ in iter_records_stream(stream):
                if record.record_id == 0x00:
                    if rows_start == -1:
                        rows_start = record.header_start
                    last_row_header_end = record.data_end
                elif 0x01 <= record.record_id <= 0x0B:
                    last_cell_end = record.data_end
                elif record.record_id == 0x25 and record.length == 6:
                    last_wrapper = record.header_start

        if rows_start == -1:
            size = os.path.getsize(sheet_path)
            insert_at = last_wrapper if last_wrapper >= 0 else size
            return _RowsRegion(insert_at, insert_at, [])
        rows_end = max(last_cell_end, last_row_header_end)
        if rows_end < 0:
            rows_end = last_wrapper if last_wrapper >= 0 else os.path.getsize(sheet_path)
        return _RowsRegion(rows_start, rows_end, [])

    @staticmethod
    def _patch_sheet_file(
        sheet_path: str | os.PathLike[str], last_row: int, last_col: int
    ) -> None:
        dimension_patched = False
        with open(sheet_path, "r+b") as stream:
            for record, _ in iter_records_stream(stream):
                if (
                    record.record_id == 0x98
                    and record.length >= 36
                    and not dimension_patched
                ):
                    stream.seek(record.data_start + 24)
                    stream.write(struct.pack("<i", last_row))
                    stream.seek(record.data_start + 32)
                    stream.write(struct.pack("<i", last_col))
                    stream.seek(record.data_end)
                    dimension_patched = True
                elif record.record_id == 0xA1 and record.length >= 16:
                    stream.seek(record.data_start + 4)
                    stream.write(struct.pack("<i", last_row))
                    stream.seek(record.data_start + 12)
                    stream.write(struct.pack("<i", last_col))
                    stream.seek(record.data_end)

    def _date_to_serial(self, value: _datetime.date | _datetime.datetime) -> float:
        if isinstance(value, _datetime.datetime):
            date_value = value.replace(tzinfo=None)
        else:
            date_value = _datetime.datetime.combine(value, _datetime.time())
        epoch = _datetime.datetime(1904, 1, 1) if self._date_1904 else _datetime.datetime(1899, 12, 30)
        return (date_value - epoch).total_seconds() / 86400

    def _find_rows_region(self, sheet_data: bytes) -> "_RowsRegion":
        rows_start = -1
        last_cell_end = -1
        last_row_header_end = -1
        last_wrapper = -1
        auto_filter_ranges: list[int] = []
        for record in iter_records(sheet_data):
            if record.record_id == 0x00:
                if rows_start == -1:
                    rows_start = record.header_start
                last_row_header_end = record.data_end
            elif 0x01 <= record.record_id <= 0x0B:
                last_cell_end = record.data_end
            elif record.record_id == 0x25 and record.length == 6:
                last_wrapper = record.header_start
            elif record.record_id == 0xA1 and record.length >= 16:
                auto_filter_ranges.append(record.data_start)

        if rows_start == -1:
            insert_at = last_wrapper if last_wrapper >= 0 else len(sheet_data)
            return _RowsRegion(insert_at, insert_at, auto_filter_ranges)
        rows_end = max(last_cell_end, last_row_header_end)
        if rows_end < 0:
            rows_end = last_wrapper if last_wrapper >= 0 else len(sheet_data)
        return _RowsRegion(rows_start, rows_end, auto_filter_ranges)

    @staticmethod
    def _patch_dimension(sheet_data: bytearray, last_row: int, last_col: int) -> None:
        for record in iter_records(bytes(sheet_data)):
            if record.record_id == 0x98 and record.length >= 36:
                struct.pack_into("<i", sheet_data, record.data_start + 24, last_row)
                struct.pack_into("<i", sheet_data, record.data_start + 32, last_col)
                return

    def _patch_pivot_caches(self, sheet_name: str, last_row: int, last_col: int) -> None:
        for path in sorted(
            name
            for name in self._package_names()
            if re.fullmatch(r"xl/pivotCache/pivotCacheDefinition\d*\.bin", name)
        ):
            data = bytearray(self._package.get(path))
            patched = False
            source_matched = False
            for record in iter_records(bytes(data)):
                payload = data[record.data_start : record.data_end]
                if record.record_id == 0xBB and record.length >= 23:
                    if len(payload) < 7:
                        continue
                    character_count = struct.unpack_from("<I", payload, 3)[0]
                    name_start = record.data_start + 7
                    reference_start = name_start + character_count * 2
                    if reference_start + 16 > record.data_end:
                        continue
                    source_sheet, _ = read_utf16(data, name_start, character_count)
                    if source_sheet == sheet_name:
                        struct.pack_into("<i", data, reference_start + 4, last_row)
                        struct.pack_into("<i", data, reference_start + 12, last_col)
                        patched = True
                        source_matched = True
            if source_matched:
                for record in iter_records(bytes(data)):
                    if record.record_id == 0xB3 and record.length >= 4:
                        data[record.data_start + 3] |= 0x04
                        break
            if patched:
                self._package.set(path, bytes(data))

    def _package_names(self) -> list[str]:
        return self._package.names()


class _RowsRegion:
    def __init__(self, rows_start: int, rows_end: int, auto_filter_ranges: list[int]):
        self.rows_start = rows_start
        self.rows_end = rows_end
        self.auto_filter_ranges = auto_filter_ranges
