from dataclasses import dataclass
import datetime
from decimal import Decimal
import zipfile
import xml.sax.saxutils as saxutils
from typing import Tuple, Iterable
import io
import itertools
from .formats import _is_formatted_cell, _unwrap_cell, _get_format


@dataclass
class FilterData:
    sheet_index: int = 0
    start_column: int = 0
    end_column: int = 0  
    start_row: int = 0
    end_row: int = 0


class XlsxWriter:
    def __init__(self, filename: str, compressionLevel: int = 4, useSharedStrings: bool = True):
        self.filename = filename
        self._worksheet_data: list[Tuple[str, Iterable[list[any]], bool]] = []
        self._shared_strings: list[str] = []
        self._shared_strings_dict: dict[str, int] = {}
        self._sheet_count = 0
        self._sst_unique_count = 0
        self._sst_all_count = 0
        self._filtered_data_list: list[FilterData] = []
        self._sheetCnt = 1
        self._compressionLevel = compressionLevel
        self._useSharedStrings = useSharedStrings
        self._zf: zipfile.ZipFile = None
        self._letters = self._generate_column_letters()

        self._format_registry: dict[str, int] = {}
        self._format_xf_map: dict[str, int] = {}
        self._next_numfmt_id = 164
        self._next_xf_index = 4

    def __enter__(self):
        if self._zf is None:
            self._zf = zipfile.ZipFile(self.filename, 'w', zipfile.ZIP_DEFLATED, compresslevel=self._compressionLevel)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self._zf is not None:
                self.save()
        finally:
            if self._zf is not None:
                self._zf.close()
                self._zf = None
        
        return False

    def _generate_column_letters(self) -> list[str]:
        letters = []
        for i in range(26):
            letters.append(chr(65 + i))
        for i in range(26):
            for j in range(26):
                letters.append(chr(65 + i) + chr(65 + j))
        for i in range(24):
            for j in range(26):
                for k in range(26):
                    if i == 23 and j > 24:
                        break
                    letters.append(chr(65 + i) + chr(65 + j) + chr(65 + k))
        return letters

    def add_sheet(self, sheet_name: str, hidden: bool = False):
        self._sheet_count += 1
        self._worksheet_data.append((sheet_name, iter([]), hidden))

    def _register_format(self, fmt_string: str) -> int:
        if fmt_string in self._format_xf_map:
            return self._format_xf_map[fmt_string]

        numfmt_id = self._next_numfmt_id
        self._next_numfmt_id += 1
        self._format_registry[fmt_string] = numfmt_id

        xf_index = self._next_xf_index
        self._next_xf_index += 1
        self._format_xf_map[fmt_string] = xf_index

        return xf_index

    def write_sheet(self, data: Iterable[list[any]]):
        if not self._worksheet_data:
            self.add_sheet("Sheet1")
        
        sheet_name, _, hidden = self._worksheet_data[self._sheet_count - 1]
        self._worksheet_data[self._sheet_count - 1] = (sheet_name, data, hidden)

        sheet_id = self._sheet_count
        with self._zf.open(f"xl/worksheets/sheet{sheet_id}.xml", 'w') as sheet_file:
            self._write_worksheet_xml(sheet_file, data, self._sheet_count - 1)

    def save(self):
        if self._zf is None:
            raise RuntimeError("Zip file not initialized. Use context manager or initialize manually.")
            
        self._zf.writestr("[Content_Types].xml", self._create_content_types())
        self._zf.writestr("_rels/.rels", self._create_root_rels())
        self._zf.writestr("xl/workbook.xml", self._create_workbook_xml())
        self._zf.writestr("xl/styles.xml", self._create_styles_xml())
        self._zf.writestr("xl/_rels/workbook.xml.rels", self._create_workbook_rels())

        if self._useSharedStrings and self._shared_strings:
            with self._zf.open("xl/sharedStrings.xml", 'w') as sst_file:
                self._write_shared_strings_xml(sst_file)
        
    def close(self):
        if self._zf is not None:
            self._zf.close()
            self._zf = None

    def _create_content_types(self) -> str:
        parts = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(self._sheet_count)
        )

        shared_strings_part = ""
        if self._useSharedStrings and self._shared_strings:
            shared_strings_part = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{parts}
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
{shared_strings_part}
</Types>'''

    def _create_root_rels(self) -> str:
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    def _create_workbook_rels(self) -> str:
        relationships = [
            f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
            for i in range(self._sheet_count)
        ]
        
        style_rid = self._sheet_count + 1
        relationships.append(f'<Relationship Id="rId{style_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
        
        if self._useSharedStrings and self._shared_strings:
            shared_strings_rid = self._sheet_count + 2
            relationships.append(f'<Relationship Id="rId{shared_strings_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>')

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{"".join(relationships)}
</Relationships>'''

    def _create_workbook_xml(self) -> str:
        sheets = ""
        for i, (sheet_name, _, hidden) in enumerate(self._worksheet_data):
            sheet_id = i + 1
            hidden_attr = ' state="hidden"' if hidden else ''
            sheets += f'<sheet name="{saxutils.escape(sheet_name)}" sheetId="{sheet_id}"{hidden_attr} r:id="rId{sheet_id}"/>'

        defined_names = ""
        if self._filtered_data_list:
            defined_names = "<definedNames>"
            for filter_data in self._filtered_data_list:
                sheet_name = self._worksheet_data[filter_data.sheet_index][0]
                start_col = self._letters[filter_data.start_column]
                end_col = self._letters[filter_data.end_column]
                range_ref = f"{sheet_name}!${start_col}${filter_data.start_row + 1}:${end_col}${filter_data.end_row + 1}"
                defined_names += f'<definedName name="_xlnm._FilterDatabase" localSheetId="{filter_data.sheet_index}" hidden="1">{range_ref}</definedName>'
            defined_names += "</definedNames>"

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
<workbookPr defaultThemeVersion="124226"/>
<bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>
<sheets>{sheets}</sheets>
{defined_names}
<calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>'''

    def _create_styles_xml(self) -> str:
        numfmts = ""
        xf_entries = []

        base_xf_count = 4

        numfmt_count = len(self._format_registry)
        if numfmt_count > 0:
            numfmts_parts = []
            for fmt_string, numfmt_id in self._format_registry.items():
                escaped_fmt = saxutils.escape(fmt_string, {'"': '&quot;'})
                numfmts_parts.append(f'<numFmt numFmtId="{numfmt_id}" formatCode="{escaped_fmt}"/>')
            numfmts = f'<numFmts count="{numfmt_count}">{"".join(numfmts_parts)}</numFmts>'

        xf_entries.append('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')
        xf_entries.append('<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
        xf_entries.append('<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
        xf_entries.append('<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>')

        for fmt_string, xf_index in sorted(self._format_xf_map.items(), key=lambda x: x[1]):
            numfmt_id = self._format_registry[fmt_string]
            xf_entries.append(f'<xf numFmtId="{numfmt_id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')

        total_xf = len(xf_entries)

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{numfmts}
<fonts count="2">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/><b/></font>
</fonts>
<fills count="2">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1">
<border><left/><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="{total_xf}">
{"".join(xf_entries)}
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>'''

    def _write_shared_strings_xml(self, sst_file: io.BufferedWriter):
        content = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{self._sst_all_count}" uniqueCount="{self._sst_unique_count}">'''
        
        for s in self._shared_strings:
            escaped_string = saxutils.escape(s)
            if s and (s[0] == ' ' or s[0] == '\t'):
                content += f'<si><t xml:space="preserve">{escaped_string}</t></si>'
            else:
                content += f'<si><t>{escaped_string}</t></si>'
        
        content += '</sst>'
        sst_file.write(content.encode('utf-8'))

    def _calculate_column_widths(self, data_iterator, max_cols: int) -> list[float]:
        col_widths = [8.43] * max_cols
        
        rows_analyzed = 0
        for row in data_iterator:
            if rows_analyzed >= 100:
                break
            
            for col_idx, cell in enumerate(row):
                if col_idx >= max_cols:
                    break
                    
                cell_val = _unwrap_cell(cell) if cell is not None else None
                if cell_val is not None:
                    if isinstance(cell_val, datetime.datetime):
                        width = 18.0
                    elif isinstance(cell_val, datetime.date):
                        width = 10.14
                    else:
                        width = max(8.43, len(str(cell_val)) * 1.25 + 2)
                    
                    col_widths[col_idx] = max(col_widths[col_idx], min(width, 255))
            
            rows_analyzed += 1
        
        return col_widths

    def _write_worksheet_xml(self, sheet_file: io.BufferedWriter, worksheet_data: Iterable[list[any]], worksheet_index: int):
        buffer = io.StringIO()
        
        buffer.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        buffer.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        
        data_list = list(worksheet_data)
        max_cols = max(len(row) for row in data_list) if data_list else 1
        num_rows = len(data_list)
        
        if num_rows > 0 and max_cols > 0:
            end_cell = f"{self._letters[max_cols - 1]}{num_rows}"
            buffer.write(f'<dimension ref="A1:{end_cell}"/>')
        
        if worksheet_index == 0:
            buffer.write('<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
        else:
            buffer.write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
        
        col_widths = self._calculate_column_widths(iter(data_list), max_cols)
        
        if any(width != 8.43 for width in col_widths):
            buffer.write('<cols>')
            for i, width in enumerate(col_widths):
                if width != 8.43:
                    buffer.write(f'<col min="{i + 1}" max="{i + 1}" width="{width:.2f}" bestFit="1" customWidth="1"/>')
            buffer.write('</cols>')
        
        buffer.write('<sheetData>')
        
        last_row_idx = 0
        for row_idx, row in enumerate(data_list):
            last_row_idx = row_idx
            buffer.write('<row>')
            
            for col_idx, cell in enumerate(row):
                if cell is None:
                    buffer.write('<c/>')
                    continue

                fmt_string = _get_format(cell)
                if fmt_string is not None:
                    xf_index = self._register_format(fmt_string)
                    cell = _unwrap_cell(cell)
                    has_custom_fmt = True
                else:
                    has_custom_fmt = False

                if isinstance(cell, str):
                    if has_custom_fmt:
                        self._write_string_cell(buffer, cell, row_idx == 0, xf_index)
                    else:
                        self._write_string_cell(buffer, cell, row_idx == 0)
                elif isinstance(cell, bool):
                    s_attr = f' s="{xf_index}"' if has_custom_fmt else ''
                    buffer.write(f'<c t="b"{s_attr}><v>{"1" if cell else "0"}</v></c>')
                elif isinstance(cell, (int, float, Decimal)):
                    if isinstance(cell, float):
                        if cell != cell:
                            self._write_string_cell(buffer, "NaN", row_idx == 0)
                            continue
                        elif cell == float('inf'):
                            self._write_string_cell(buffer, "\u221e", row_idx == 0)
                            continue
                        elif cell == float('-inf'):
                            self._write_string_cell(buffer, "-\u221e", row_idx == 0)
                            continue
                    s_attr = f' s="{xf_index}"' if has_custom_fmt else ''
                    buffer.write(f'<c{s_attr}><v>{float(cell)}</v></c>')
                elif isinstance(cell, datetime.datetime):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_string_cell(buffer, str(cell), row_idx == 0)
                    else:
                        if cell.tzinfo is not None:
                            cell = cell.replace(tzinfo=None)
                        excel_date = (cell - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0
                        if has_custom_fmt:
                            buffer.write(f'<c s="{xf_index}"><v>{excel_date}</v></c>')
                        else:
                            buffer.write(f'<c s="2"><v>{excel_date}</v></c>')
                elif isinstance(cell, datetime.date):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_string_cell(buffer, str(cell), row_idx == 0)
                    else:
                        excel_date = (datetime.datetime.combine(cell, datetime.time()) - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0
                        if has_custom_fmt:
                            buffer.write(f'<c s="{xf_index}"><v>{excel_date}</v></c>')
                        else:
                            buffer.write(f'<c s="1"><v>{excel_date}</v></c>')
                else:
                    self._write_string_cell(buffer, str(cell), row_idx == 0)
            
            buffer.write('</row>')
        
        buffer.write('</sheetData>')
        
        if last_row_idx > 0:
            start_col = self._letters[0]
            end_col = self._letters[max_cols - 1]
            buffer.write(f'<autoFilter ref="{start_col}1:{end_col}{last_row_idx + 1}"/>')
            
            self._filtered_data_list.append(FilterData(
                sheet_index=worksheet_index,
                start_column=0,
                end_column=max_cols - 1,
                start_row=0,
                end_row=last_row_idx
            ))
        
        buffer.write('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        buffer.write('</worksheet>')
        
        sheet_file.write(buffer.getvalue().encode('utf-8'))

    def _write_string_cell(self, buffer: io.StringIO, cell_value: str, is_header: bool = False, xf_index: int = None):
        escaped_value = saxutils.escape(cell_value)
        
        if self._useSharedStrings:
            self._sst_all_count += 1
            if cell_value not in self._shared_strings_dict:
                string_index = self._sst_unique_count
                self._shared_strings_dict[cell_value] = string_index
                self._shared_strings.append(cell_value)
                self._sst_unique_count += 1
            else:
                string_index = self._shared_strings_dict[cell_value]
            
            if xf_index is not None:
                style_ref = f' s="{xf_index}"'
            else:
                style_ref = ' s="3"' if is_header else ''
            buffer.write(f'<c t="s"{style_ref}><v>{string_index}</v></c>')
        else:
            if xf_index is not None:
                style_ref = f' s="{xf_index}"'
            else:
                style_ref = ' s="3"' if is_header else ''
            if cell_value and (cell_value[0] == ' ' or cell_value[0] == '\t'):
                buffer.write(f'<c t="inlineStr"{style_ref}><is><t xml:space="preserve">{escaped_value}</t></is></c>')
            else:
                buffer.write(f'<c t="inlineStr"{style_ref}><is><t>{escaped_value}</t></is></c>')
