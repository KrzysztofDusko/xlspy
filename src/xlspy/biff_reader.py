import struct
from enum import IntEnum
from typing import Optional


class CellType(IntEnum):
    doubleVal = 0
    boolVal = 1
    stringVal = 2
    sharedString = 3
    nullValue = 4


class BiffReader:
    _sheet = 0x9C

    _xf = 0x2F
    _cellXfStart = 0x269
    _cellXfEnd = 0x26A
    _cellStyleXfStart = 0x272
    _cellStyleXfEnd = 0x273
    _numberFormatStart = 0x267
    _numberFormat = 0x2C
    _numberFormatEnd = 0x268
    _sharedStringStart = 159
    _stringItem = 0x13

    _row = 0x00
    _blank = 0x01
    _number = 0x02
    _boolError = 0x03
    _bool = 0x04
    _float = 0x05
    _string = 0x06
    _sharedString = 0x07
    _formulaString = 0x08
    _formulaNumber = 0x09
    _formulaBool = 0x0A
    _formulaError = 0x0B

    def __init__(self, stream):
        self._stream = stream
        self._buffer = bytearray(128)

        self.workbook_id = 0
        self.rec_id = ""
        self.workbook_name = ""
        self.is_sheet = False

        self.in_cell_xf = False
        self.in_cell_style_xf = False
        self.in_number_format = False
        self.parent_cell_style_xf = 0
        self.number_format_index = 0
        self.format = 0
        self.format_string = ""

        self.shared_string_value: Optional[str] = None
        self.shared_string_unique_count = 0

        self.cell_type = CellType.nullValue
        self.int_value = 0
        self.double_val = 0.0
        self.bool_value = False
        self.string_value = ""
        self.column_num = -1
        self.xf_index = 0
        self.read_cell = False
        self.row_index = -1

    def _try_read_variable_value(self):
        value = 0
        b = self._stream.read(1)
        if len(b) == 0:
            return False, 0

        b1 = b[0]
        value = b1 & 0x7F

        if (b1 & 0x80) == 0:
            return True, value

        b = self._stream.read(1)
        if len(b) == 0:
            return False, 0
        b2 = b[0]
        value = ((b2 & 0x7F) << 7) | value

        if (b2 & 0x80) == 0:
            return True, value

        b = self._stream.read(1)
        if len(b) == 0:
            return False, 0
        b3 = b[0]
        value = ((b3 & 0x7F) << 14) | value

        if (b3 & 0x80) == 0:
            return True, value

        b = self._stream.read(1)
        if len(b) == 0:
            return False, 0
        b4 = b[0]
        value = ((b4 & 0x7F) << 21) | value

        return True, value

    def _read_record(self):
        ok1, record_id = self._try_read_variable_value()
        if not ok1:
            return False, 0, None
        ok2, record_length = self._try_read_variable_value()
        if not ok2:
            return False, 0, None
        if record_length < len(self._buffer):
            buf = self._buffer[:record_length]
        else:
            buf = bytearray(record_length)
        if record_length > 0:
            read = self._stream.readinto(buf)
            if read != record_length:
                return False, 0, None
        return True, record_id, bytes(buf)

    @staticmethod
    def _get_dword(buffer, offset):
        return (buffer[offset + 3] << 24) + (buffer[offset + 2] << 16) + (buffer[offset + 1] << 8) + buffer[offset]

    @staticmethod
    def _get_int32(buffer, offset):
        result = buffer[offset + 3] << 24
        result += buffer[offset + 2] << 16
        result += buffer[offset + 1] << 8
        result += buffer[offset]
        return result

    @staticmethod
    def _get_word(buffer, offset):
        return (buffer[offset + 1] << 8) + buffer[offset]

    @staticmethod
    def _get_string(buffer, offset, length):
        chars = []
        for i in range(offset, offset + 2 * length, 2):
            ch = (buffer[i + 1] << 8) + buffer[i]
            chars.append(chr(ch))
        return ''.join(chars)

    @staticmethod
    def _get_nullable_string(buffer, offset):
        length = BiffReader._get_dword(buffer, offset)
        offset += 4
        if length == 0xFFFFFFFF:
            return None, offset
        chars = []
        end = offset + length * 2
        while offset < end:
            ch = (buffer[offset + 1] << 8) + buffer[offset]
            chars.append(chr(ch))
            offset += 2
        return ''.join(chars), offset

    @staticmethod
    def _get_rk_number(buffer, offset):
        flags = buffer[offset]
        val = struct.unpack_from('<I', buffer, offset)[0]
        if (flags & 0x02) != 0:
            result = float(BiffReader._get_int32(buffer, offset) >> 2)
        else:
            masked = val & 0xFFFFFFFC
            packed = struct.pack('<Q', masked << 32)
            result = struct.unpack('<d', packed)[0]
        if (flags & 0x01) != 0:
            result /= 100
        return result

    @staticmethod
    def _get_double(buffer, offset):
        return struct.unpack_from('<d', buffer, offset)[0]

    def read_workbook(self):
        ok, record_id, buffer = self._read_record()
        if not ok:
            return False

        self.is_sheet = False
        if record_id == self._sheet:
            self.workbook_id = self._get_dword(buffer, 4)
            offset = 8
            self.rec_id, offset = self._get_nullable_string(buffer, offset)
            name_length = self._get_dword(buffer, offset)
            self.workbook_name = self._get_string(buffer, offset + 4, name_length)
            self.is_sheet = True
        return True

    def read_styles(self):
        ok, record_id, buffer = self._read_record()
        if not ok:
            return False

        if record_id == self._cellXfStart:
            self.in_cell_xf = True
        elif record_id == self._cellXfEnd:
            self.in_cell_xf = False
        elif record_id == self._cellStyleXfStart:
            self.in_cell_style_xf = True
        elif record_id == self._cellStyleXfEnd:
            self.in_cell_style_xf = False
        elif record_id == self._numberFormatStart:
            self.in_number_format = True
        elif record_id == self._numberFormatEnd:
            self.in_number_format = False
        elif record_id == self._xf and self.in_cell_style_xf:
            pass
        elif record_id == self._xf and self.in_cell_xf:
            self.parent_cell_style_xf = self._get_word(buffer, 0)
            self.number_format_index = self._get_word(buffer, 2)
        elif record_id == self._numberFormat and self.in_number_format:
            self.format = self._get_word(buffer, 0)
            length = self._get_dword(buffer, 2)
            self.format_string = self._get_string(buffer, 2 + 4, length)

        return True

    def read_shared_strings(self):
        ok, record_id, buffer = self._read_record()
        if not ok:
            return False

        if record_id == self._stringItem:
            length = self._get_dword(buffer, 1)
            self.shared_string_value = self._get_string(buffer, 1 + 4, length)
        elif record_id == self._sharedStringStart:
            self.shared_string_unique_count = self._get_dword(buffer, 4)
            self.shared_string_value = None
        else:
            self.shared_string_value = None

        return True

    def read_worksheet(self):
        ok, record_id, buffer = self._read_record()
        if not ok:
            return False

        self.read_cell = False
        self.column_num = -1

        if record_id == self._row:
            self.row_index = self._get_int32(buffer, 0)
        elif record_id in (self._blank, self._boolError, self._formulaError):
            self.read_cell = True
            self.cell_type = CellType.nullValue
        elif record_id == self._number:
            self.double_val = self._get_rk_number(buffer, 8)
            self.read_cell = True
            self.cell_type = CellType.doubleVal
        elif record_id in (self._bool, self._formulaBool):
            self.bool_value = (buffer[8] == 1)
            self.read_cell = True
            self.cell_type = CellType.boolVal
        elif record_id in (self._formulaNumber, self._float):
            self.double_val = self._get_double(buffer, 8)
            self.read_cell = True
            self.cell_type = CellType.doubleVal
        elif record_id in (self._string, self._formulaString):
            length = self._get_dword(buffer, 8)
            self.string_value = self._get_string(buffer, 8 + 4, length)
            self.read_cell = True
            self.cell_type = CellType.stringVal
        elif record_id == self._sharedString:
            self.int_value = self._get_dword(buffer, 8)
            self.read_cell = True
            self.cell_type = CellType.sharedString

        if self.read_cell:
            self.column_num = self._get_dword(buffer, 0)
            self.xf_index = self._get_dword(buffer, 4) & 0xFFFFFF

        return True

    def close(self):
        self._stream.close()
