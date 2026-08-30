"""Update worksheet data in an existing XLSX package."""

from __future__ import annotations

import datetime as _datetime
import math
import numbers
import os
import re
from decimal import Decimal
from typing import Any, Iterable, Literal, Sequence
from xml.parsers import expat

from .updater_utils import (
    ZipPackage,
    column_index_to_letter,
    column_letter_to_index,
    escape_xml_text,
    local_xml_name,
    materialize_rows,
    parse_relationships_xml,
    parse_shared_strings_xml,
    replace_or_add_xml_attribute,
    resolve_zip_target,
    strip_invalid_xml_characters,
    trim_trailing_empty_rows,
    unescape_xml_text,
    unwrap_cell,
)


StyleFallback = Literal["inherit", "general"]


def _xml_attribute(attrs: dict[str, str], wanted: str) -> str | None:
    for name, value in attrs.items():
        if local_xml_name(name) == wanted:
            return value
    return None


def _tag_attribute(opening_tag: str, wanted: str) -> str | None:
    pattern = re.compile(
        rf"(?:^|\s)(?:[A-Za-z_][\w.-]*:)?{re.escape(wanted)}\s*=\s*([\"'])(.*?)\1",
        re.DOTALL,
    )
    match = pattern.search(opening_tag)
    return match.group(2) if match else None


def _date_like_format(format_code: str) -> bool:
    # Remove quoted literals and escaped characters before looking for date
    # tokens.  This is intentionally conservative: a false positive only
    # chooses an existing date style for a date value.
    simplified = re.sub(r'"(?:[^"]|"")*"', "", format_code)
    simplified = re.sub(r"\\.", "", simplified).lower()
    simplified = re.sub(r"\[[^\]]*\]", "", simplified)
    return bool(re.search(r"[dy]", simplified) or re.search(r"m", simplified))


class XlsxUpdater:
    """Replace the data region of sheets in an existing XLSX file.

    The workbook, styles, drawings and all unrelated ZIP members are kept in
    place.  Only the selected worksheet XML, shared strings and pivot-cache
    metadata are changed.
    """

    def __init__(self, path: str | os.PathLike[str]):
        self._package = ZipPackage(path)
        if self._package.has("xl/workbook.bin"):
            raise ValueError("The input is XLSB; use XlsbUpdater")
        if not self._package.has("xl/workbook.xml"):
            raise ValueError("XLSX workbook.xml was not found")

        self._date_1904 = self._load_date_system()
        self._sheets = self._load_sheet_structure()
        self._shared_strings_loaded = False
        self._shared_strings_values: list[str] = []
        self._shared_string_indices: dict[str, int] = {}
        self._shared_strings_total = 0
        self._shared_strings_pending: list[str] = []
        self._shared_strings_prefix = ""
        self._shared_strings_dirty = False
        self._use_shared_strings = False

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
        """Replace one worksheet's used data region.

        ``rows`` is materialised once, so generators are supported but the
        operation remains atomic with respect to the in-memory package.  A
        ``(value, format_string)`` cell is accepted for compatibility with
        the writers; its value is used and no new style is created.
        """

        if style_fallback not in ("inherit", "general"):
            raise ValueError("style_fallback must be 'inherit' or 'general'")
        try:
            sheet_path = self._sheets[sheet_name]
        except KeyError as exc:
            raise KeyError(f"Worksheet not found: {sheet_name}") from exc

        materialized = trim_trailing_empty_rows(materialize_rows(rows))
        header_values = list(headers) if headers is not None else None
        old_sheet_xml = self._package.get(sheet_path).decode("utf-8")
        self._ensure_shared_strings_loaded()
        if self._use_shared_strings:
            self._shared_strings_total = max(
                0,
                self._shared_strings_total
                - self._shared_string_ref_count(old_sheet_xml),
            )
            self._shared_strings_dirty = True

        header_styles, data_styles = self._collect_styles(
            old_sheet_xml, headers_provided=headers is not None
        )
        date_style = self._find_date_style_index() if style_fallback == "inherit" else None
        worksheet_prefix = self._tag_prefix(old_sheet_xml, "sheetData")
        width = max(
            [len(row) for row in materialized]
            + ([len(header_values)] if header_values is not None else [])
            + [0]
        )
        total_rows = len(materialized) + (1 if header_values is not None else 0)
        last_col = width - 1
        dimension_ref = self._dimension_ref(total_rows, last_col)

        rows_xml: list[str] = []
        next_row = 1
        if header_values is not None:
            cells = [
                self._value_cell(
                    value,
                    column,
                    next_row,
                    header_styles,
                    data_styles,
                    date_style,
                    style_fallback,
                    header=True,
                    prefix=worksheet_prefix,
                )
                for column, value in enumerate(header_values)
            ]
            rows_xml.append(self._row_xml(next_row, cells, worksheet_prefix))
            next_row += 1

        for row in materialized:
            cells = [
                self._value_cell(
                    value,
                    column,
                    next_row,
                    header_styles,
                    data_styles,
                    date_style,
                    style_fallback,
                    header=False,
                    prefix=worksheet_prefix,
                )
                for column, value in enumerate(row)
            ]
            rows_xml.append(self._row_xml(next_row, cells, worksheet_prefix))
            next_row += 1

        updated_sheet_xml = self._patch_sheet_xml(
            old_sheet_xml, "".join(rows_xml), dimension_ref
        )
        self._package.set(sheet_path, updated_sheet_xml.encode("utf-8"))
        self._commit_shared_strings()
        self._patch_pivot_metadata(sheet_name, len(materialized), dimension_ref)

    def save(self, output_path: str | os.PathLike[str] | None = None) -> None:
        """Save to a new path or atomically replace the input file."""

        self._package.save(output_path)

    def to_bytes(self) -> bytes:
        """Return the current package as XLSX bytes."""

        return self._package.to_bytes()

    def _load_sheet_structure(self) -> dict[str, str]:
        workbook_rels_path = "xl/_rels/workbook.xml.rels"
        if not self._package.has(workbook_rels_path):
            raise ValueError("XLSX workbook relationships were not found")
        relationships = parse_relationships_xml(self._package.get(workbook_rels_path))

        sheets: dict[str, str] = {}
        parser = expat.ParserCreate()

        def start_element(name: str, attrs: dict[str, str]) -> None:
            if local_xml_name(name) != "sheet":
                return
            sheet_name = _xml_attribute(attrs, "name")
            relationship_id = _xml_attribute(attrs, "id")
            if not sheet_name or not relationship_id:
                return
            target = relationships.get(relationship_id)
            if target is None:
                raise ValueError(f"Relationship {relationship_id!r} for sheet is missing")
            sheet_path = resolve_zip_target(target, "xl")
            if not self._package.has(sheet_path):
                raise ValueError(f"Worksheet part was not found: {sheet_path}")
            sheets[sheet_name] = sheet_path

        parser.StartElementHandler = start_element
        parser.Parse(self._package.get("xl/workbook.xml"), True)
        if not sheets:
            raise ValueError("No worksheets were found in workbook.xml")
        return sheets

    def _load_date_system(self) -> bool:
        date_1904 = False
        parser = expat.ParserCreate()

        def start_element(name: str, attrs: dict[str, str]) -> None:
            nonlocal date_1904
            if local_xml_name(name) != "workbookPr":
                return
            value = _xml_attribute(attrs, "date1904")
            date_1904 = value is not None and value.strip().lower() in {
                "1",
                "true",
                "on",
            }

        parser.StartElementHandler = start_element
        parser.Parse(self._package.get("xl/workbook.xml"), True)
        return date_1904

    def _ensure_shared_strings_loaded(self) -> None:
        if self._shared_strings_loaded:
            return
        self._shared_strings_loaded = True
        if not self._package.has("xl/sharedStrings.xml"):
            self._use_shared_strings = False
            return

        xml_data = self._package.get("xl/sharedStrings.xml")
        self._shared_strings_values = parse_shared_strings_xml(xml_data)
        self._shared_string_indices = {
            value: index for index, value in enumerate(self._shared_strings_values)
        }
        decoded_xml = xml_data.decode("utf-8")
        opening_match = re.search(
            r"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)sst\b[^>]*>", decoded_xml
        )
        if opening_match:
            count_value = _tag_attribute(opening_match.group(0), "count")
            self._shared_strings_total = int(count_value) if count_value else len(self._shared_strings_values)
            self._shared_strings_prefix = opening_match.group("prefix") or ""
        else:
            self._shared_strings_total = len(self._shared_strings_values)
        self._use_shared_strings = True

    def _commit_shared_strings(self) -> None:
        if not self._use_shared_strings or (
            not self._shared_strings_pending and not self._shared_strings_dirty
        ):
            return
        xml_data = self._package.get("xl/sharedStrings.xml").decode("utf-8")
        opening_match = re.search(r"<(?P<tag>(?:[A-Za-z_][\w.-]*:)?sst\b[^>]*)>", xml_data)
        closing_match = re.search(r"</(?:[A-Za-z_][\w.-]*:)?sst\s*>\s*$", xml_data)
        if not opening_match or not closing_match:
            raise ValueError("Malformed XLSX sharedStrings.xml")

        opening_tag = opening_match.group(0)
        opening_tag = replace_or_add_xml_attribute(
            opening_tag, "count", self._shared_strings_total
        )
        opening_tag = replace_or_add_xml_attribute(
            opening_tag, "uniqueCount", len(self._shared_strings_values)
        )
        items = "".join(
            f'<{self._shared_strings_prefix}si>'
            f'<{self._shared_strings_prefix}t{self._xml_space_attribute(value)}>'
            f'{escape_xml_text(value)}</{self._shared_strings_prefix}t>'
            f'</{self._shared_strings_prefix}si>'
            for value in self._shared_strings_pending
        )
        xml_data = (
            xml_data[: opening_match.start()]
            + opening_tag
            + xml_data[opening_match.end() : closing_match.start()]
            + items
            + xml_data[closing_match.start() :]
        )
        self._package.set("xl/sharedStrings.xml", xml_data.encode("utf-8"))
        self._shared_strings_pending.clear()
        self._shared_strings_dirty = False

    def _shared_string_index(self, value: str) -> int:
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
    def _shared_string_ref_count(sheet_xml: str) -> int:
        count = 0
        parser = expat.ParserCreate()

        def start_element(name: str, attrs: dict[str, str]) -> None:
            nonlocal count
            if local_xml_name(name) == "c" and _xml_attribute(attrs, "t") == "s":
                count += 1

        parser.StartElementHandler = start_element
        parser.Parse(sheet_xml.encode("utf-8"), True)
        return count

    @staticmethod
    def _xml_space_attribute(value: str) -> str:
        preserve = (
            value[:1].isspace()
            or value[-1:].isspace()
            or any(character in value for character in "\t\n\r")
        )
        return ' xml:space="preserve"' if preserve else ""

    @staticmethod
    def _dimension_ref(total_rows: int, last_col: int) -> str:
        if total_rows <= 0 or last_col < 0:
            return "A1"
        return f"A1:{column_index_to_letter(last_col)}{total_rows}"

    def _collect_styles(
        self, sheet_xml: str, *, headers_provided: bool
    ) -> tuple[dict[int, int], dict[int, int]]:
        header_counts: dict[int, dict[int, int]] = {}
        data_counts: dict[int, dict[int, int]] = {}
        current_row: int | None = None
        current_column = 0
        row_sequence = 0
        parser = expat.ParserCreate()

        def start_element(name: str, attrs: dict[str, str]) -> None:
            nonlocal current_row, current_column, row_sequence
            local = local_xml_name(name)
            if local == "row":
                row_value = _xml_attribute(attrs, "r")
                if row_value:
                    current_row = int(row_value)
                    row_sequence = current_row
                else:
                    row_sequence += 1
                    current_row = row_sequence
                current_column = 0
            elif local == "c" and current_row is not None:
                style_value = _xml_attribute(attrs, "s")
                reference = _xml_attribute(attrs, "r")
                column = (
                    column_letter_to_index(reference) if reference else current_column
                )
                current_column = column + 1
                if style_value:
                    style = int(style_value)
                    if style > 0:
                        target = (
                            header_counts
                            if headers_provided and current_row == 1
                            else data_counts
                        )
                        counts = target.setdefault(column, {})
                        counts[style] = counts.get(style, 0) + 1

        def end_element(name: str) -> None:
            nonlocal current_row, current_column
            if local_xml_name(name) == "row":
                current_row = None
                current_column = 0

        parser.StartElementHandler = start_element
        parser.EndElementHandler = end_element
        parser.Parse(sheet_xml.encode("utf-8"), True)

        def dominant(counts: dict[int, dict[int, int]]) -> dict[int, int]:
            return {
                column: max(styles.items(), key=lambda item: (item[1], -item[0]))[0]
                for column, styles in counts.items()
            }

        return dominant(header_counts), dominant(data_counts)

    def _find_date_style_index(self) -> int | None:
        if not self._package.has("xl/styles.xml"):
            return None
        custom_formats: dict[int, str] = {}
        cell_xfs: list[int] = []
        in_cell_xfs = False
        cell_xfs_depth = 0
        depth = 0
        parser = expat.ParserCreate()

        def start_element(name: str, attrs: dict[str, str]) -> None:
            nonlocal in_cell_xfs, cell_xfs_depth, depth
            depth += 1
            local = local_xml_name(name)
            if local == "numFmt":
                number_format_id = _xml_attribute(attrs, "numFmtId")
                format_code = _xml_attribute(attrs, "formatCode")
                if number_format_id and format_code:
                    custom_formats[int(number_format_id)] = format_code
            elif local == "cellXfs":
                in_cell_xfs = True
                cell_xfs_depth = depth
            elif local == "xf" and in_cell_xfs and depth == cell_xfs_depth + 1:
                number_format_id = _xml_attribute(attrs, "numFmtId")
                cell_xfs.append(int(number_format_id) if number_format_id else 0)

        def end_element(name: str) -> None:
            nonlocal in_cell_xfs, depth
            local = local_xml_name(name)
            if local == "cellXfs" and depth == cell_xfs_depth:
                in_cell_xfs = False
            depth -= 1

        parser.StartElementHandler = start_element
        parser.EndElementHandler = end_element
        parser.Parse(self._package.get("xl/styles.xml"), True)

        built_in_date_formats = set(range(14, 23)) | {45, 46, 47}
        for index, number_format_id in enumerate(cell_xfs):
            if number_format_id in built_in_date_formats:
                return index
            format_code = custom_formats.get(number_format_id)
            if format_code and _date_like_format(format_code):
                return index
        return None

    def _style_for_cell(
        self,
        column: int,
        value: Any,
        header_styles: dict[int, int],
        data_styles: dict[int, int],
        date_style: int | None,
        style_fallback: StyleFallback,
        header: bool,
    ) -> int:
        if style_fallback == "general":
            return 0
        style = (header_styles if header else data_styles).get(column, 0)
        value = unwrap_cell(value)
        if isinstance(value, (_datetime.date, _datetime.datetime)) and style == 0:
            return date_style or 0
        return style

    def _value_cell(
        self,
        value: Any,
        column: int,
        row_number: int,
        header_styles: dict[int, int],
        data_styles: dict[int, int],
        date_style: int | None,
        style_fallback: StyleFallback,
        *,
        header: bool,
        prefix: str,
    ) -> str:
        value = unwrap_cell(value)
        if value is None:
            return ""
        style = self._style_for_cell(
            column,
            value,
            header_styles,
            data_styles,
            date_style,
            style_fallback,
            header,
        )
        cell_reference = f"{column_index_to_letter(column)}{row_number}"
        style_attribute = f' s="{style}"' if style > 0 else ""
        cell_prefix = f'<{prefix}c r="{cell_reference}"{style_attribute}'

        if isinstance(value, bool):
            return f'{cell_prefix} t="b"><{prefix}v>{1 if value else 0}</{prefix}v></{prefix}c>'
        if isinstance(value, (_datetime.datetime, _datetime.date)):
            serial = self._date_to_serial(value)
            return f"{cell_prefix}><{prefix}v>{serial:.15g}</{prefix}v></{prefix}c>"
        if isinstance(value, numbers.Real) or isinstance(value, Decimal):
            numeric_value = float(value)
            if math.isfinite(numeric_value):
                return f"{cell_prefix}><{prefix}v>{numeric_value:.15g}</{prefix}v></{prefix}c>"

        text = strip_invalid_xml_characters(str(value))
        if self._use_shared_strings:
            index = self._shared_string_index(text)
            return f'{cell_prefix} t="s"><{prefix}v>{index}</{prefix}v></{prefix}c>'
        return (
            f'{cell_prefix} t="inlineStr"><{prefix}is><{prefix}t{self._xml_space_attribute(text)}>'
            f"{escape_xml_text(text)}</{prefix}t></{prefix}is></{prefix}c>"
        )

    def _date_to_serial(self, value: _datetime.date | _datetime.datetime) -> float:
        if isinstance(value, _datetime.datetime):
            date_value = value.replace(tzinfo=None)
        else:
            date_value = _datetime.datetime.combine(value, _datetime.time())
        epoch = _datetime.datetime(1904, 1, 1) if self._date_1904 else _datetime.datetime(1899, 12, 30)
        return (date_value - epoch).total_seconds() / 86400

    @staticmethod
    def _row_xml(row_number: int, cells: list[str], prefix: str = "") -> str:
        return f'<{prefix}row r="{row_number}">' + "".join(cells) + f"</{prefix}row>"

    @staticmethod
    def _tag_prefix(xml_data: str, tag_name: str) -> str:
        match = re.search(
            rf"<(?P<prefix>[A-Za-z_][\w.-]*:)?{re.escape(tag_name)}\b",
            xml_data,
        )
        return match.group("prefix") or "" if match else ""

    @staticmethod
    def _patch_sheet_xml(sheet_xml: str, rows_xml: str, dimension_ref: str) -> str:
        sheet_data_pattern = re.compile(
            r"<(?P<prefix>[A-Za-z_][\w.-]*:)?sheetData\b(?P<attrs>[^>]*)/>", re.DOTALL
        )
        empty_match = sheet_data_pattern.search(sheet_xml)
        if empty_match:
            opening = (
                f"<{empty_match.group('prefix') or ''}sheetData"
                f"{empty_match.group('attrs')}>"
            )
            closing = f"</{empty_match.group('prefix') or ''}sheetData>"
            replacement = opening + rows_xml + closing
            sheet_xml = (
                sheet_xml[: empty_match.start()]
                + replacement
                + sheet_xml[empty_match.end() :]
            )
        else:
            opening_match = re.search(
                r"<(?P<prefix>[A-Za-z_][\w.-]*:)?sheetData\b[^>]*>",
                sheet_xml,
            )
            if not opening_match:
                raise ValueError("Worksheet sheetData element was not found")
            prefix = opening_match.group("prefix") or ""
            closing_match = re.search(
                rf"</{re.escape(prefix)}sheetData\s*>",
                sheet_xml[opening_match.end() :],
            )
            if not closing_match:
                raise ValueError("Worksheet sheetData closing element was not found")
            closing_start = opening_match.end() + closing_match.start()
            closing_end = opening_match.end() + closing_match.end()
            sheet_xml = (
                sheet_xml[: opening_match.end()]
                + rows_xml
                + sheet_xml[closing_start:closing_end]
                + sheet_xml[closing_end:]
            )

        dimension_match = re.search(
            r"<(?P<prefix>[A-Za-z_][\w.-]*:)?dimension\b[^>]*>", sheet_xml
        )
        if dimension_match:
            opening = dimension_match.group(0)
            replacement = replace_or_add_xml_attribute(opening, "ref", dimension_ref)
            sheet_xml = (
                sheet_xml[: dimension_match.start()]
                + replacement
                + sheet_xml[dimension_match.end() :]
            )
        else:
            worksheet_match = re.search(r"<(?P<prefix>[A-Za-z_][\w.-]*:)?worksheet\b[^>]*>", sheet_xml)
            if not worksheet_match:
                raise ValueError("Worksheet root element was not found")
            prefix = XlsxUpdater._tag_prefix(sheet_xml, "worksheet")
            dimension = f'<{prefix}dimension ref="{escape_xml_text(dimension_ref)}"/>'
            sheet_xml = (
                sheet_xml[: worksheet_match.end()]
                + dimension
                + sheet_xml[worksheet_match.end() :]
            )

        auto_filter_match = re.search(
            r"<(?P<prefix>[A-Za-z_][\w.-]*:)?autoFilter\b[^>]*>", sheet_xml
        )
        if auto_filter_match:
            opening = auto_filter_match.group(0)
            replacement = replace_or_add_xml_attribute(opening, "ref", dimension_ref)
            sheet_xml = (
                sheet_xml[: auto_filter_match.start()]
                + replacement
                + sheet_xml[auto_filter_match.end() :]
            )
        return sheet_xml

    def _patch_pivot_metadata(
        self, sheet_name: str, record_count: int, dimension_ref: str
    ) -> None:
        pivot_cache_paths = sorted(
            name
            for name in self._package_names()
            if re.fullmatch(r"xl/pivotCache/pivotCacheDefinition\d*\.xml", name)
        )
        for path in pivot_cache_paths:
            xml_data = self._package.get(path).decode("utf-8")
            source_match = re.search(
                r"<(?P<prefix>[A-Za-z_][\w.-]*:)?worksheetSource\b[^>]*>", xml_data
            )
            if not source_match:
                continue
            source_tag = source_match.group(0)
            source_sheet = _tag_attribute(source_tag, "sheet")
            if source_sheet is not None:
                source_sheet = unescape_xml_text(source_sheet)
            if source_sheet != sheet_name:
                continue
            source_tag = replace_or_add_xml_attribute(source_tag, "ref", dimension_ref)
            xml_data = (
                xml_data[: source_match.start()]
                + source_tag
                + xml_data[source_match.end() :]
            )
            root_match = re.search(
                r"<(?P<prefix>[A-Za-z_][\w.-]*:)?pivotCacheDefinition\b[^>]*>",
                xml_data,
            )
            if root_match:
                root_tag = replace_or_add_xml_attribute(
                    root_match.group(0), "recordCount", record_count
                )
                root_tag = replace_or_add_xml_attribute(
                    root_tag, "refreshOnLoad", 1
                )
                xml_data = (
                    xml_data[: root_match.start()]
                    + root_tag
                    + xml_data[root_match.end() :]
                )
            self._package.set(path, xml_data.encode("utf-8"))

        for path in sorted(
            name
            for name in self._package_names()
            if re.fullmatch(r"xl/pivotTables/pivotTable\d*\.xml", name)
        ):
            xml_data = self._package.get(path).decode("utf-8")
            root_match = re.search(
                r"<(?P<prefix>[A-Za-z_][\w.-]*:)?pivotTableDefinition\b[^>]*>",
                xml_data,
            )
            if not root_match:
                continue
            root_tag = replace_or_add_xml_attribute(root_match.group(0), "refreshOnLoad", 1)
            xml_data = (
                xml_data[: root_match.start()]
                + root_tag
                + xml_data[root_match.end() :]
            )
            self._package.set(path, xml_data.encode("utf-8"))

    def _package_names(self) -> list[str]:
        return self._package.names()
