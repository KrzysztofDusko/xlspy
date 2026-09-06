"""Minimal BIFF12 helpers for editing XLSB ZIP members."""

from __future__ import annotations

import struct
from dataclasses import dataclass
from typing import BinaryIO


@dataclass(frozen=True)
class Biff12Record:
    """A BIFF12 record and its offsets in the containing byte buffer."""

    header_start: int
    data_start: int
    data_end: int
    record_id: int
    length: int


def read_vlq(data: bytes, offset: int) -> tuple[int, int]:
    """Read a BIFF12 variable-length integer."""

    value = 0
    shift = 0
    position = offset
    while position < len(data):
        byte = data[position]
        position += 1
        value |= (byte & 0x7F) << shift
        if byte & 0x80 == 0:
            return value, position
        shift += 7
        if shift > 63:
            raise ValueError("BIFF12 variable-length integer is too long")
    raise ValueError("Truncated BIFF12 variable-length integer")


def vlq_bytes(value: int) -> bytes:
    """Encode a non-negative integer using BIFF12's 7-bit representation."""

    if value < 0:
        raise ValueError("BIFF12 variable-length integers cannot be negative")
    encoded = bytearray()
    while value >= 0x80:
        encoded.append((value & 0x7F) | 0x80)
        value >>= 7
    encoded.append(value)
    return bytes(encoded)


def iter_records(data: bytes):
    """Yield all BIFF12 records in a byte buffer."""

    offset = 0
    while offset < len(data):
        header_start = offset
        record_id, offset = read_vlq(data, offset)
        length, data_start = read_vlq(data, offset)
        data_end = data_start + length
        if data_end > len(data):
            raise ValueError(
                f"Truncated BIFF12 record 0x{record_id:x}: "
                f"expected {length} bytes, only {len(data) - data_start} available"
            )
        yield Biff12Record(header_start, data_start, data_end, record_id, length)
        offset = data_end


def _read_vlq_stream(stream: BinaryIO, offset: int) -> tuple[int | None, int]:
    """Read one BIFF12 variable-length integer from a binary stream."""

    value = 0
    shift = 0
    position = offset
    first = stream.read(1)
    if not first:
        return None, position
    while True:
        byte = first[0]
        position += 1
        value |= (byte & 0x7F) << shift
        if byte & 0x80 == 0:
            return value, position
        shift += 7
        if shift > 63:
            raise ValueError("BIFF12 variable-length integer is too long")
        first = stream.read(1)
        if not first:
            raise ValueError("Truncated BIFF12 variable-length integer")


def iter_records_stream(stream: BinaryIO):
    """Yield BIFF12 records from a seekable binary stream.

    The record payload is yielded as the second item.  Only one record is
    materialised at a time, while offsets remain relative to the stream.
    """

    offset = 0
    while True:
        header_start = offset
        record_id, offset = _read_vlq_stream(stream, offset)
        if record_id is None:
            return
        length, data_start = _read_vlq_stream(stream, offset)
        if length is None:
            raise ValueError("Truncated BIFF12 record length")
        payload = stream.read(length)
        if len(payload) != length:
            raise ValueError(
                f"Truncated BIFF12 record 0x{record_id:x}: "
                f"expected {length} bytes, only {len(payload)} available"
            )
        data_end = data_start + length
        yield (
            Biff12Record(header_start, data_start, data_end, record_id, length),
            payload,
        )
        offset = data_end


def build_record(record_id: int, payload: bytes = b"") -> bytes:
    """Build one BIFF12 record from its ID and payload."""

    return vlq_bytes(record_id) + vlq_bytes(len(payload)) + payload


def read_utf16(data: bytes, offset: int, character_count: int) -> tuple[str, int]:
    """Read a BIFF12 UTF-16LE string with a character count prefix supplied by the caller."""

    end = offset + character_count * 2
    if end > len(data):
        raise ValueError("Truncated BIFF12 UTF-16 string")
    return data[offset:end].decode("utf-16le"), end


@dataclass
class SharedStringsBin:
    values: list[str]
    total_count: int
    unique_count: int
    begin_offset: int | None
    end_offset: int | None


def parse_shared_strings_bin(data: bytes) -> SharedStringsBin:
    """Parse the plain-string SST emitted by Excel and :class:`XlsbWriter`."""

    values: list[str] = []
    total_count = 0
    unique_count = 0
    begin_offset: int | None = None
    end_offset: int | None = None

    for record in iter_records(data):
        payload = data[record.data_start : record.data_end]
        if record.record_id == 0x9F:
            if len(payload) < 8:
                raise ValueError("Malformed XLSB shared-string header")
            total_count, unique_count = struct.unpack_from("<II", payload, 0)
            begin_offset = record.header_start
        elif record.record_id == 0x13:
            if len(payload) < 5:
                raise ValueError("Malformed XLSB shared-string item")
            character_count = struct.unpack_from("<I", payload, 1)[0]
            string, string_end = read_utf16(payload, 5, character_count)
            if string_end > len(payload):
                raise ValueError("Malformed XLSB shared-string item length")
            values.append(string)
        elif record.record_id == 0xA0:
            end_offset = record.header_start

    if begin_offset is None:
        raise ValueError("XLSB shared-string header was not found")
    return SharedStringsBin(values, total_count, unique_count, begin_offset, end_offset)
