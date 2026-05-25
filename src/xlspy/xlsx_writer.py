from dataclasses import dataclass
import datetime
from decimal import Decimal
import zipfile
import xml.sax.saxutils as saxutils
from typing import Tuple, Iterable
import io
import itertools


@dataclass
class FilterData:
    sheet_index: int = 0
    start_column: int = 0
    end_column: int = 0  
    start_row: int = 0
    end_row: int = 0


class XlsxWriter:
    def __init__(self, filename: str, compressionLevel: int = 4, useSharedStrings: bool = True):
        """
        Initializes an XlsxWriter object to create an .xlsx file.

        Args:
            filename (str): The path to the output .xlsx file.
            compressionLevel (int): The compression level for the zip archive.
            useSharedStrings (bool): Whether to use shared strings table for optimization.
        """
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

    def __enter__(self):
        """Enter the runtime context for the XlsxWriter."""
        if self._zf is None:
            self._zf = zipfile.ZipFile(self.filename, 'w', zipfile.ZIP_DEFLATED, compresslevel=self._compressionLevel)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        Exit the runtime context for the XlsxWriter.
        
        Args:
            exc_type: Exception type
            exc_val: Exception value
            exc_tb: Exception traceback
            
        Returns:
            bool: False to propagate any exceptions, True to suppress them
        """
        try:
            if self._zf is not None:
                self.save()
        finally:
            if self._zf is not None:
                self._zf.close()
                self._zf = None
        
        return False

    def _generate_column_letters(self) -> list[str]:
        """Generate Excel column letters (A, B, ..., Z, AA, AB, ...)"""
        letters = []
        
        # Single letters A-Z
        for i in range(26):
            letters.append(chr(65 + i))
        
        # Double letters AA-ZZ
        for i in range(26):
            for j in range(26):
                letters.append(chr(65 + i) + chr(65 + j))
        
        # Triple letters AAA-XYZ (limited to reasonable range)
        for i in range(24):  # A-X
            for j in range(26):
                for k in range(26):
                    if i == 23 and j > 24:  # Stop at XY
                        break
                    letters.append(chr(65 + i) + chr(65 + j) + chr(65 + k))
        
        return letters

    def add_sheet(self, sheet_name: str, hidden: bool = False):
        """Add a new sheet to the workbook."""
        self._sheet_count += 1
        self._worksheet_data.append((sheet_name, iter([]), hidden))

    def write_sheet(self, data: Iterable[list[any]]):
        """
        Writes data to the current sheet.

        Args:
            data: An iterable of rows, where each row is a list of cell values.
        """
        if not self._worksheet_data:
            self.add_sheet("Sheet1")
        
        sheet_name, _, hidden = self._worksheet_data[self._sheet_count - 1]
        self._worksheet_data[self._sheet_count - 1] = (sheet_name, data, hidden)

        sheet_id = self._sheet_count
        with self._zf.open(f"xl/worksheets/sheet{sheet_id}.xml", 'w') as sheet_file:
            self._write_worksheet_xml(sheet_file, data, self._sheet_count - 1)

    def save(self):
        """Save all content to the zip file."""
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
        """Explicitly close the writer."""
        if self._zf is not None:
            self._zf.close()
            self._zf = None

    def _create_content_types(self) -> str:
        """Create the [Content_Types].xml file."""
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
        """Create the _rels/.rels file."""
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    def _create_workbook_rels(self) -> str:
        """Create the xl/_rels/workbook.xml.rels file."""
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
        """Create the xl/workbook.xml file."""
        sheets = ""
        for i, (sheet_name, _, hidden) in enumerate(self._worksheet_data):
            sheet_id = i + 1
            hidden_attr = ' state="hidden"' if hidden else ''
            sheets += f'<sheet name="{saxutils.escape(sheet_name)}" sheetId="{sheet_id}"{hidden_attr} r:id="rId{sheet_id}"/>'

        # Add autofilter defined names if needed
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
        """Create the xl/styles.xml file."""
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
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
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>'''

    def _write_shared_strings_xml(self, sst_file: io.BufferedWriter):
        """Write the shared strings XML file."""
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
        """Calculate optimal column widths based on content."""
        col_widths = [8.43] * max_cols  # Default width
        
        # Analyze first 100 rows for width estimation
        rows_analyzed = 0
        for row in data_iterator:
            if rows_analyzed >= 100:
                break
            
            for col_idx, cell in enumerate(row):
                if col_idx >= max_cols:
                    break
                    
                if cell is not None:
                    if isinstance(cell, datetime.datetime):
                        width = 18.0
                    elif isinstance(cell, datetime.date):
                        width = 10.14
                    else:
                        width = max(8.43, len(str(cell)) * 1.25 + 2)
                    
                    col_widths[col_idx] = max(col_widths[col_idx], min(width, 255))
            
            rows_analyzed += 1
        
        return col_widths

    def _write_worksheet_xml(self, sheet_file: io.BufferedWriter, worksheet_data: Iterable[list[any]], worksheet_index: int):
        """Write the worksheet XML content."""
        buffer = io.StringIO()
        
        # Start worksheet
        buffer.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        buffer.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        
        # Convert to list to allow multiple iterations
        data_list = list(worksheet_data)
        max_cols = max(len(row) for row in data_list) if data_list else 1
        num_rows = len(data_list)
        
        # Add dimension
        if num_rows > 0 and max_cols > 0:
            end_cell = f"{self._letters[max_cols - 1]}{num_rows}"
            buffer.write(f'<dimension ref="A1:{end_cell}"/>')
        
        # Add sheet views with frozen first row
        if worksheet_index == 0:
            buffer.write('<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
        else:
            buffer.write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
        
        # Calculate column widths
        col_widths = self._calculate_column_widths(iter(data_list), max_cols)
        
        # Write column definitions
        if any(width != 8.43 for width in col_widths):
            buffer.write('<cols>')
            for i, width in enumerate(col_widths):
                if width != 8.43:
                    buffer.write(f'<col min="{i + 1}" max="{i + 1}" width="{width:.2f}" bestFit="1" customWidth="1"/>')
            buffer.write('</cols>')
        
        # Write sheet data
        buffer.write('<sheetData>')
        
        last_row_idx = 0
        for row_idx, row in enumerate(data_list):
            last_row_idx = row_idx
            buffer.write('<row>')
            
            for col_idx, cell in enumerate(row):
                if cell is None:
                    buffer.write('<c/>')
                    continue
                
                if isinstance(cell, str):
                    self._write_string_cell(buffer, cell, row_idx == 0)
                elif isinstance(cell, bool):
                    buffer.write(f'<c t="b"><v>{"1" if cell else "0"}</v></c>')
                elif isinstance(cell, (int, float, Decimal)):
                    # Handle special float values
                    if isinstance(cell, float):
                        if cell != cell:  # NaN
                            self._write_string_cell(buffer, "NaN", row_idx == 0)
                            continue
                        elif cell == float('inf'):
                            self._write_string_cell(buffer, "∞", row_idx == 0)
                            continue
                        elif cell == float('-inf'):
                            self._write_string_cell(buffer, "-∞", row_idx == 0)
                            continue
                    buffer.write(f'<c><v>{float(cell)}</v></c>')
                elif isinstance(cell, datetime.datetime):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_string_cell(buffer, str(cell), row_idx == 0)
                    else:
                        # Remove timezone info for Excel date calculation
                        if cell.tzinfo is not None:
                            cell = cell.replace(tzinfo=None)
                        excel_date = (cell - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0
                        buffer.write(f'<c s="2"><v>{excel_date}</v></c>')
                elif isinstance(cell, datetime.date):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_string_cell(buffer, str(cell), row_idx == 0)
                    else:
                        excel_date = (datetime.datetime.combine(cell, datetime.time()) - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0
                        buffer.write(f'<c s="1"><v>{excel_date}</v></c>')
                else:
                    self._write_string_cell(buffer, str(cell), row_idx == 0)
            
            buffer.write('</row>')
        
        buffer.write('</sheetData>')
        
        # Add autofilter if there's data
        if last_row_idx > 0:
            start_col = self._letters[0]
            end_col = self._letters[max_cols - 1]
            buffer.write(f'<autoFilter ref="{start_col}1:{end_col}{last_row_idx + 1}"/>')
            
            # Store filter data
            self._filtered_data_list.append(FilterData(
                sheet_index=worksheet_index,
                start_column=0,
                end_column=max_cols - 1,
                start_row=0,
                end_row=last_row_idx
            ))
        
        buffer.write('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        buffer.write('</worksheet>')
        
        # Write to file
        sheet_file.write(buffer.getvalue().encode('utf-8'))

    def _write_string_cell(self, buffer: io.StringIO, cell_value: str, is_header: bool = False):
        """Write a string cell to the buffer."""
        escaped_value = saxutils.escape(cell_value)
        
        if self._useSharedStrings:
            # Use shared strings
            self._sst_all_count += 1
            if cell_value not in self._shared_strings_dict:
                string_index = self._sst_unique_count
                self._shared_strings_dict[cell_value] = string_index
                self._shared_strings.append(cell_value)
                self._sst_unique_count += 1
            else:
                string_index = self._shared_strings_dict[cell_value]
            
            # Style 3 is bold header style, style 0 is normal
            style_ref = ' s="3"' if is_header else ''
            buffer.write(f'<c t="s"{style_ref}><v>{string_index}</v></c>')
        else:
            # Inline strings
            style_ref = ' s="3"' if is_header else ''
            if cell_value and (cell_value[0] == ' ' or cell_value[0] == '	'):
                buffer.write(f'<c t="inlineStr"{style_ref}><is><t xml:space="preserve">{escaped_value}</t></is></c>')
            else:
                buffer.write(f'<c t="inlineStr"{style_ref}><is><t>{escaped_value}</t></is></c>')