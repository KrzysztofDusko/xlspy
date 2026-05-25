import datetime
import zipfile
from enum import IntEnum
from typing import Dict, List, Optional, Set, Generator, Any
from xml.parsers.expat import ParserCreate

from .biff_reader import BiffReader, CellType
from ._accel import HAVE_C_EXT, read_xlsb_worksheet as _c_read_xlsb_worksheet


class ExcelDataType(IntEnum):
    Null = 0
    Int32 = 1
    Int64 = 2
    Double = 3
    DateTime = 4
    String = 5
    Boolean = 6
    Error = 7


class FieldInfo:
    def __init__(self):
        self.type = ExcelDataType.Null
        self.int32_value: int = 0
        self.int64_value: int = 0
        self.double_value: float = 0.0
        self.dt_value: datetime.datetime = datetime.datetime(1899, 12, 30)
        self.bool_value: bool = False
        self.str_value: str = ""


_date_excel_masks: Set[str] = {
    '[$-F800]dddd,\\ mmmm\\ dd,\\ yyyy',
    'd\\-mm;@',
    'yy\\-mm\\-dd;@',
    '[$-415]d\\ mmm;@',
    '[$-415]d\\ mmm\\ yy;@',
    '[$-415]dd\\ mmm\\ yy;@',
    '[$-415]mmm\\ yy;@',
    '[$-415]mmmm\\ yy;@',
    '[$-415]d\\ mmmm\\ yyyy;@',
    'yyyy\\-mm\\-dd\\ hh:mm',
    'yyyy\\-mm\\-dd\\ hh:mm:ss',
    'yyyy\\-mm\\-dd;@',
    'yyyy\\-mm\\-dd',
    '[$-409]dd\\-mm\\-yy\\ h:mm\\ AM/PM;@',
    'dd\\-mm\\-yy\\ h:mm;@',
    '[$-415]mmmmm;@',
    '[$-415]mmmmm\\.yy;@',
    '\\-m\\-yyyy;@',
    '[$-415]d\\-mmm\\-yyyy;@',
    'd\\-m\\-yyyy;@',
}


def _looks_like_date_format(fmt_code: str) -> bool:
    if not fmt_code:
        return False
    fmt_lower = fmt_code.lower()
    has_y = 'y' in fmt_lower
    has_m = 'm' in fmt_lower
    has_d = 'd' in fmt_lower
    has_h = 'h' in fmt_lower
    has_s = 's' in fmt_lower
    date_indicators = sum([has_y, has_m and not has_h and not has_s, has_d])
    return date_indicators >= 2

_number_formats_type_dict: Dict[int, type] = {
    0: str,
    1: float,
    2: float,
    3: float,
    4: float,
    5: float,
    6: float,
    7: float,
    9: float,
    14: datetime.datetime,
    15: datetime.datetime,
    16: datetime.datetime,
    17: datetime.datetime,
    18: datetime.datetime,
    19: datetime.datetime,
    20: datetime.datetime,
    21: datetime.datetime,
    22: datetime.datetime,
    44: float,
}


def _get_date_num_fmt_ids() -> Set[int]:
    return {k for k, v in _number_formats_type_dict.items() if v is datetime.datetime}

_letters = [chr(65 + i) for i in range(26)]
for i in range(26):
    for j in range(26):
        _letters.append(chr(65 + i) + chr(65 + j))
for i in range(24):
    for j in range(26):
        for k in range(26):
            if i == 23 and j > 24:
                break
            _letters.append(chr(65 + i) + chr(65 + j) + chr(65 + k))

_letter_to_column_num: Dict[str, int] = {letter: idx for idx, letter in enumerate(_letters)}


class ExcelReader:
    _open_xml_info_string = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def __init__(self, path: Optional[str] = None, use_memory_stream_for_xlsb: bool = True):
        self._archive: Optional[zipfile.ZipFile] = None
        self._is_xlsb = False
        self._worksheet_id_to_name: Dict[str, str] = {}
        self._worksheet_name_to_id: Dict[str, str] = {}
        self._worksheet_id_to_location: Dict[str, str] = {}
        self._shared_string_array: List[str] = []
        self._styles_cell_xfs_array: List[dict] = []
        self._shared_strings_location: Optional[str] = None
        self._styles_location: Optional[str] = None
        self._path: Optional[str] = path
        self.use_memory_stream_for_xlsb = use_memory_stream_for_xlsb

    def open(self, path: str):
        ext = path.lower()
        if ext.endswith('xlsb'):
            self._is_xlsb = True
            self._open_xlsb(path)
        else:
            self._is_xlsb = False
            self._open_xlsx(path)

    def close(self):
        if self._archive is not None:
            self._archive.close()
            self._archive = None

    def __enter__(self):
        if self._path is not None:
            self.open(self._path)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False

    def get_sheet_names(self) -> List[str]:
        return list(self._worksheet_name_to_id.keys())

    def get_sheet_index(self, sheet_name: str) -> int:
        names = self.get_sheet_names()
        for i, n in enumerate(names):
            if n == sheet_name:
                return i
        raise ValueError(f"Sheet '{sheet_name}' not found")

    def get_rows(self, sheet_name: str) -> Generator[List[Any], None, None]:
        if self._is_xlsb:
            yield from self._get_rows_xlsb(sheet_name)
        else:
            yield from self._get_rows_xlsx(sheet_name)

    def read_all(self, sheet_name: str) -> List[List[Any]]:
        return list(self.get_rows(sheet_name))

    def _open_xlsx(self, path: str):
        self._archive = zipfile.ZipFile(path, 'r')

        with self._archive.open('xl/workbook.xml') as f:
            content = f.read().decode('utf-8')
            self._parse_workbook_xml(content)

        self._fill_rels('xml')

        if self._styles_location is not None:
            self._fill_styles()

        if self._shared_strings_location is not None:
            self._fill_shared_strings()

    def _open_xlsb(self, path: str):
        self._archive = zipfile.ZipFile(path, 'r')

        with self._archive.open('xl/workbook.bin') as f:
            raw = f.read()
            stream = _MemoryStream(raw)
            reader = BiffReader(stream)
            while reader.read_workbook():
                if reader.is_sheet:
                    self._worksheet_id_to_name[reader.rec_id] = reader.workbook_name
                    self._worksheet_name_to_id[reader.workbook_name] = reader.rec_id
            reader.close()

        self._fill_rels('bin')

        if self._styles_location is not None:
            self._fill_bin_styles()

        if self._shared_strings_location is not None:
            self._fill_bin_shared_strings()

    def _parse_workbook_xml(self, content: str):
        name = ""
        r_id = ""
        in_sheet = False

        parser = ParserCreate()
        parser.StartElementHandler = lambda tag, attrs: self._workbook_start_elem(tag, attrs)
        parser.Parse(content, True)

    def _workbook_start_elem(self, tag: str, attrs: Dict[str, str]):
        local = tag.split('}')[-1] if '}' in tag else tag
        if local == 'sheet':
            name = attrs.get('name', '')
            r_id = attrs.get('r:id') or attrs.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            self._worksheet_id_to_name[r_id] = name
            self._worksheet_name_to_id[name] = r_id
        elif local == 'pivotCache':
            pass

    def _fill_rels(self, ext: str):
        rels_path = f'xl/_rels/workbook.{ext}.rels'
        if rels_path not in self._archive.namelist():
            return

        with self._archive.open(rels_path) as f:
            content = f.read().decode('utf-8')

        parser = ParserCreate()
        parser.StartElementHandler = lambda tag, attrs: self._rels_start_elem(tag, attrs)
        parser.Parse(content, True)

    def _rels_start_elem(self, tag: str, attrs: Dict[str, str]):
        local = tag.split('}')[-1] if '}' in tag else tag
        if local == 'Relationship':
            target = attrs.get('Target', '')
            rel_type = attrs.get('Type', '')
            r_id = attrs.get('Id', '')

            if rel_type == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet':
                self._worksheet_id_to_location[r_id] = target
            elif rel_type == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings':
                self._shared_strings_location = target
            elif rel_type == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles':
                self._styles_location = target

    def _fill_shared_strings(self):
        if self._shared_strings_location is None:
            return
        entry_path = self._resolve_path(self._shared_strings_location)
        with self._archive.open(entry_path) as f:
            content = f.read().decode('utf-8')

        self._shared_string_array = []
        in_t = False
        char_buf = []

        parser = ParserCreate()
        parser.StartElementHandler = lambda tag, attrs: None
        parser.EndElementHandler = lambda tag: None
        parser.CharacterDataHandler = lambda data: None

        class SstHandler:
            def __init__(self, strings_list):
                self.strings = strings_list
                self.in_si = False
                self.in_t = False
                self.buf = ""

            def start(self, tag, attrs):
                local = tag.split('}')[-1] if '}' in tag else tag
                if local == 'si':
                    self.in_si = True
                    self.buf = ""
                elif local == 't' and self.in_si:
                    self.in_t = True

            def end(self, tag):
                local = tag.split('}')[-1] if '}' in tag else tag
                if local == 't' and self.in_t:
                    self.in_t = False
                elif local == 'si':
                    if self.in_si:
                        self.strings.append(self.buf)
                        self.in_si = False
                        self.buf = ""

            def chars(self, data):
                if self.in_t:
                    self.buf += data

        h = SstHandler(self._shared_string_array)
        parser.StartElementHandler = h.start
        parser.EndElementHandler = h.end
        parser.CharacterDataHandler = h.chars
        parser.Parse(content, True)

    def _fill_bin_shared_strings(self):
        if self._shared_strings_location is None:
            return
        entry_path = self._resolve_path(self._shared_strings_location)
        with self._archive.open(entry_path) as f:
            raw = f.read()
        stream = _MemoryStream(raw)
        reader = BiffReader(stream)
        reader.read_shared_strings()
        if reader.shared_string_unique_count != 0:
            self._shared_string_array = [""] * reader.shared_string_unique_count
        else:
            self._shared_string_array = []

        idx = 0
        while reader.read_shared_strings():
            val = reader.shared_string_value
            if val is not None:
                if idx >= len(self._shared_string_array):
                    self._shared_string_array.append(val)
                else:
                    self._shared_string_array[idx] = val
                idx += 1
        if idx < len(self._shared_string_array):
            self._shared_string_array = self._shared_string_array[:idx]
        reader.close()

    def _fill_styles(self):
        if self._styles_location is None:
            return
        entry_path = self._resolve_path(self._styles_location)
        with self._archive.open(entry_path) as f:
            content = f.read().decode('utf-8')

        self._styles_cell_xfs_array = []

        class StylesHandler:
            def __init__(self, xfs_list):
                self.xfs = xfs_list
                self.in_cell_xfs = False
                self.in_num_fmts = False

            def start(self, tag, attrs):
                local = tag.split('}')[-1] if '}' in tag else tag
                if local == 'cellXfs':
                    self.in_cell_xfs = True
                elif local == 'numFmts':
                    self.in_num_fmts = True
                elif local == 'xf' and self.in_cell_xfs:
                    xf_id = int(attrs.get('xfId', '0'))
                    num_fmt_id = int(attrs.get('numFmtId', '0'))
                    self.xfs.append({'xfId': xf_id, 'numFmtId': num_fmt_id})
                elif local == 'numFmt' and self.in_num_fmts:
                    num_fmt_id = int(attrs.get('numFmtId', '0'))
                    fmt_code = attrs.get('formatCode', '')
                    if num_fmt_id not in _number_formats_type_dict:
                        if fmt_code in _date_excel_masks or _looks_like_date_format(fmt_code):
                            _number_formats_type_dict[num_fmt_id] = datetime.datetime
                        else:
                            _number_formats_type_dict[num_fmt_id] = str

            def end(self, tag):
                local = tag.split('}')[-1] if '}' in tag else tag
                if local == 'cellXfs':
                    self.in_cell_xfs = False
                elif local == 'numFmts':
                    self.in_num_fmts = False

            def chars(self, data):
                pass

        h = StylesHandler(self._styles_cell_xfs_array)
        parser = ParserCreate()
        parser.StartElementHandler = h.start
        parser.EndElementHandler = h.end
        parser.CharacterDataHandler = h.chars
        parser.Parse(content, True)

    def _fill_bin_styles(self):
        if self._styles_location is None:
            return
        entry_path = self._resolve_path(self._styles_location)
        with self._archive.open(entry_path) as f:
            raw = f.read()
        stream = _MemoryStream(raw)
        reader = BiffReader(stream)

        self._styles_cell_xfs_array = []
        styles_first_time = True
        format_first_time = True

        while reader.read_styles():
            if reader.in_cell_xf:
                if styles_first_time:
                    reader.read_styles()
                    styles_first_time = False
                num_fmt_id = reader.number_format_index
                xf_id = reader.parent_cell_style_xf
                self._styles_cell_xfs_array.append({'xfId': xf_id, 'numFmtId': num_fmt_id})
            elif reader.in_number_format:
                if format_first_time:
                    reader.read_styles()
                    format_first_time = False
                fmt_code = reader.format_string
                num_fmt_id = reader.format
                if num_fmt_id not in _number_formats_type_dict:
                    if fmt_code in _date_excel_masks or _looks_like_date_format(fmt_code):
                        _number_formats_type_dict[num_fmt_id] = datetime.datetime
                    else:
                        _number_formats_type_dict[num_fmt_id] = str

        reader.close()

    @staticmethod
    def _resolve_path(target: str) -> str:
        if target.startswith('/'):
            return target.lstrip('/')
        if target.startswith('xl/'):
            return target
        return f'xl/{target}'

    def _get_rows_xlsx(self, sheet_name: str) -> Generator[List[Any], None, None]:
        r_id = self._worksheet_name_to_id[sheet_name]
        location = self._worksheet_id_to_location[r_id]
        entry_path = self._resolve_path(location)

        with self._archive.open(entry_path) as f:
            content = f.read().decode('utf-8')

        rows = []
        current_row: List[Any] = []
        current_cell = {}
        in_sheet_data = False
        in_row = False
        in_c = False
        in_v = False
        in_is = False
        in_t_elem = False
        char_buf = ""
        cell_index = 0

        class XlsxSheetHandler:
            def __init__(self, outer, strings):
                self.outer = outer
                self.shared_strings = strings
                self.rows = []
                self.current_row = []
                self.current_cell = {}
                self.in_sheet_data = False
                self.in_row = False
                self.in_c = False
                self.in_v = False
                self.in_is = False
                self.in_t_elem = False
                self.char_buf = ""
                self.cell_index = 0
                self.col_num_from_r = None
                self.first_row_cols = 0
                self.is_first_row = True

            def start(self, tag, attrs):
                local = tag.split('}')[-1] if '}' in tag else tag
                if local == 'sheetData':
                    self.in_sheet_data = True
                elif local == 'row' and self.in_sheet_data:
                    self.in_row = True
                    self.current_row = []
                    self.cell_index = 0
                elif local == 'c' and self.in_row:
                    self.in_c = True
                    r_attr = attrs.get('r', '')
                    t_attr = attrs.get('t', '')
                    s_attr = attrs.get('s', '')
                    self.col_num_from_r = None
                    if r_attr:
                        self.col_num_from_r = self.outer._parse_column_reference(r_attr)
                    self.current_cell = {
                        'r': self.col_num_from_r if self.col_num_from_r is not None else self.cell_index,
                        't': t_attr,
                        's': int(s_attr) if s_attr else None,
                        'is_inline': False,
                    }
                    self.in_is = False
                    self.in_v = False
                    self.in_t_elem = False
                    self.char_buf = ""
                elif local == 'v' and self.in_c:
                    self.in_v = True
                    self.char_buf = ""
                elif local == 'is' and self.in_c:
                    self.in_is = True
                    self.current_cell['is_inline'] = True
                elif local == 't' and self.in_c and self.in_is:
                    self.in_t_elem = True
                    self.char_buf = ""

            def end(self, tag):
                local = tag.split('}')[-1] if '}' in tag else tag
                if local == 'sheetData':
                    self.in_sheet_data = False
                elif local == 'row' and self.in_row:
                    self.in_row = False
                    self._finalize_row()
                elif local == 'c':
                    self.in_c = False
                    self._finalize_cell()
                    self.cell_index += 1
                elif local == 'v':
                    self.in_v = False
                elif local == 'is':
                    self.in_is = False
                elif local == 't' and self.in_t_elem:
                    self.in_t_elem = False

            def chars(self, data):
                if self.in_v:
                    self.char_buf += data
                elif self.in_t_elem:
                    self.char_buf += data

            def _finalize_cell(self):
                cell = self.current_cell
                t_attr = cell.get('t', '')
                val = self.char_buf
                col = cell['r']
                s_idx = cell['s']
                is_inline = cell.get('is_inline', False)

                col_num = col if col is not None else self.cell_index

                while len(self.current_row) <= col_num:
                    self.current_row.append(None)

                if is_inline:
                    self.current_row[col_num] = val
                    return

                if t_attr == 's':
                    idx = int(val) if val else 0
                    if 0 <= idx < len(self.shared_strings):
                        self.current_row[col_num] = self.shared_strings[idx]
                    else:
                        self.current_row[col_num] = val
                elif t_attr == 'b':
                    self.current_row[col_num] = (val == '1')
                elif t_attr == 'e':
                    self.current_row[col_num] = f"error:{val}"
                elif t_attr == 'str' or t_attr == 'inlineStr':
                    self.current_row[col_num] = val
                else:
                    if val == '' or val is None:
                        self.current_row[col_num] = None
                    else:
                        if s_idx is not None and s_idx < len(self.outer._styles_cell_xfs_array):
                            style = self.outer._styles_cell_xfs_array[s_idx]
                            num_fmt_id = style['numFmtId']
                            fmt_type = _number_formats_type_dict.get(num_fmt_id, str)
                            if fmt_type is datetime.datetime:
                                try:
                                    excel_date = float(val)
                                    self.current_row[col_num] = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=excel_date)
                                except (ValueError, OverflowError):
                                    self.current_row[col_num] = val if val else None
                                return

                        if '.' in val or 'E' in val or 'e' in val:
                            self.current_row[col_num] = float(val)
                        else:
                            try:
                                self.current_row[col_num] = int(val)
                            except ValueError:
                                self.current_row[col_num] = float(val)

            def _finalize_row(self):
                if self.is_first_row:
                    self.first_row_cols = len(self.current_row)
                    self.is_first_row = False
                while len(self.current_row) < self.first_row_cols:
                    self.current_row.append(None)
                self.rows.append(self.current_row[:self.first_row_cols])

        h = XlsxSheetHandler(self, self._shared_string_array)
        parser = ParserCreate()
        parser.StartElementHandler = h.start
        parser.EndElementHandler = h.end
        parser.CharacterDataHandler = h.chars
        parser.Parse(content, True)

        for row in h.rows:
            yield row

    def _get_rows_xlsb(self, sheet_name: str) -> Generator[List[Any], None, None]:
        r_id = self._worksheet_name_to_id[sheet_name]
        location = self._worksheet_id_to_location[r_id]
        entry_path = self._resolve_path(location)

        with self._archive.open(entry_path) as f:
            raw = f.read()

        if HAVE_C_EXT:
            rows = _c_read_xlsb_worksheet(
                raw,
                self._shared_string_array,
                self._styles_cell_xfs_array,
                _get_date_num_fmt_ids(),
            )
            if rows is not None:
                yield from rows
                return

        stream = _MemoryStream(raw)
        reader = BiffReader(stream)

        while not reader.read_cell:
            if not reader.read_worksheet():
                break
        if not reader.read_cell:
            reader.close()
            return

        prev_row = -1
        row_num = reader.row_index
        col_num = reader.column_num
        number_of_first_column_with_data = -1
        number_of_last_column_with_data = -1
        columns_cnt_from_first_row = -1
        is_first_row = True
        return_value = True
        prev_col_num = -1

        inner_row: Dict[int, Any] = {}

        while True:
            if not return_value:
                break

            if not is_first_row and col_num > number_of_first_column_with_data:
                for i in range(col_num - number_of_first_column_with_data):
                    inner_row[i] = None

            if prev_row != -1 and row_num > prev_row + 1 and columns_cnt_from_first_row > 0:
                for i in range(columns_cnt_from_first_row):
                    inner_row[i] = None
                how_many_empty = row_num - prev_row - 1
                prev_row += 1
                for _ in range(how_many_empty):
                    row_list = [inner_row.get(c, None) for c in range(columns_cnt_from_first_row)]
                    yield row_list
                continue

            prev_row = row_num

            for i in range(columns_cnt_from_first_row if columns_cnt_from_first_row > 0 else 0):
                inner_row[i] = None

            while row_num == prev_row and return_value:
                if reader.read_cell:
                    if is_first_row:
                        if number_of_first_column_with_data == -1:
                            number_of_first_column_with_data = col_num
                        number_of_last_column_with_data = col_num
                        columns_cnt_from_first_row = number_of_last_column_with_data - number_of_first_column_with_data + 1

                    tmp_len = col_num - number_of_first_column_with_data

                    if reader.cell_type == CellType.sharedString:
                        idx = reader.int_value
                        if 0 <= idx < len(self._shared_string_array):
                            inner_row[tmp_len] = self._shared_string_array[idx]
                        else:
                            inner_row[tmp_len] = None
                    elif reader.cell_type == CellType.stringVal:
                        inner_row[tmp_len] = reader.string_value
                    elif reader.cell_type == CellType.boolVal:
                        inner_row[tmp_len] = reader.bool_value
                    elif reader.cell_type == CellType.doubleVal:
                        double_val = reader.double_val
                        style_index = reader.xf_index
                        if style_index != 0 and style_index < len(self._styles_cell_xfs_array):
                            num_fmt_id = self._styles_cell_xfs_array[style_index]['numFmtId']
                            fmt_type = _number_formats_type_dict.get(num_fmt_id, str)
                            if fmt_type is datetime.datetime:
                                inner_row[tmp_len] = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=double_val)
                            else:
                                long_val = int(double_val)
                                if abs(long_val - double_val) < 1e-10:
                                    inner_row[tmp_len] = long_val
                                else:
                                    inner_row[tmp_len] = double_val
                        else:
                            long_val = int(double_val)
                            if abs(long_val - double_val) < 1e-10:
                                inner_row[tmp_len] = long_val
                            else:
                                inner_row[tmp_len] = double_val
                    else:
                        inner_row[tmp_len] = None

                reader.read_worksheet()
                while return_value and not reader.read_cell:
                    return_value = reader.read_worksheet()

                prev_col_num = col_num
                row_num = reader.row_index
                col_num = reader.column_num

                if not is_first_row:
                    if col_num > prev_col_num + 1 and row_num == prev_row:
                        for i in range(prev_col_num + 1, col_num):
                            idx = i - number_of_first_column_with_data
                            inner_row[idx] = None
                    elif row_num > prev_row and prev_col_num < number_of_last_column_with_data:
                        for i in range(1, number_of_last_column_with_data - prev_col_num + 1):
                            idx = prev_col_num + i - number_of_first_column_with_data
                            inner_row[idx] = None

            if is_first_row:
                is_first_row = False

            row_list = [inner_row.get(c, None) for c in range(columns_cnt_from_first_row)]
            yield row_list

        reader.close()

    @staticmethod
    def _parse_column_reference(ref: str) -> int:
        col = -1
        for ch in ref:
            if 'A' <= ch <= 'Z':
                col = (col + 1) * 26 + (ord(ch) - ord('A'))
            else:
                break
        return col


class _MemoryStream:
    def __init__(self, data: bytes):
        self._data = data
        self._pos = 0

    def read(self, size: int = -1) -> bytes:
        if size < 0:
            result = self._data[self._pos:]
            self._pos = len(self._data)
            return result
        if self._pos + size > len(self._data):
            result = self._data[self._pos:]
            self._pos = len(self._data)
            return result
        result = self._data[self._pos:self._pos + size]
        self._pos += size
        return result

    def readinto(self, buf: bytearray) -> int:
        size = len(buf)
        if self._pos + size > len(self._data):
            size = len(self._data) - self._pos
            if size <= 0:
                return 0
            buf[:size] = self._data[self._pos:self._pos + size]
            self._pos = len(self._data)
            return size
        buf[:size] = self._data[self._pos:self._pos + size]
        self._pos += size
        return size

    def close(self):
        pass

    @property
    def length(self) -> int:
        return len(self._data)
