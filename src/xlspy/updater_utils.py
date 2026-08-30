"""Small, format-neutral helpers used by the existing-file updaters.

The readers and writers in :mod:`xlspy` deliberately have very different
implementations for XML and BIFF12.  The updater layer only needs a few
shared operations: materialising a row iterable, manipulating ZIP packages,
and generating safe XML text.
"""

from __future__ import annotations

import copy
import html
import os
import posixpath
import re
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable, Sequence
from xml.parsers import expat
from xml.sax import saxutils


def materialize_rows(rows: Iterable[Sequence[Any]]) -> list[list[Any]]:
    """Materialise an iterable of rows while keeping its cell boundaries."""

    materialized: list[list[Any]] = []
    for row in rows:
        if isinstance(row, (str, bytes, bytearray)):
            raise TypeError("Each row must be a sequence of cells, not text")
        try:
            materialized.append(list(row))
        except TypeError as exc:
            raise TypeError("Each row must be an iterable of cells") from exc
    return materialized


def trim_trailing_empty_rows(rows: list[list[Any]]) -> list[list[Any]]:
    """Drop only rows whose cells are all literally ``None``.

    Empty rows in the middle of a range remain in place.  This mirrors the
    behaviour of the reference implementation and keeps an explicit blank
    row meaningful for downstream spreadsheet consumers.
    """

    last = len(rows)
    while last > 0 and all(cell is None for cell in rows[last - 1]):
        last -= 1
    return rows[:last]


def is_formatted_cell(cell: Any) -> bool:
    """Return whether a writer-style ``(value, format_string)`` wrapper exists."""

    return isinstance(cell, tuple) and len(cell) == 2 and isinstance(cell[1], str)


def unwrap_cell(cell: Any) -> Any:
    """Return the value from the optional writer-style format wrapper."""

    return cell[0] if is_formatted_cell(cell) else cell


def strip_invalid_xml_characters(value: str) -> str:
    """Remove characters which XML 1.0 cannot represent."""

    return "".join(
        character
        for character in value
        if character in "\t\n\r"
        or 0x20 <= ord(character) <= 0xD7FF
        or 0xE000 <= ord(character) <= 0xFFFD
        or 0x10000 <= ord(character) <= 0x10FFFF
    )


def escape_xml_text(value: Any) -> str:
    """Escape a value for use as XML text or an XML attribute value."""

    text = strip_invalid_xml_characters(str(value))
    return saxutils.escape(text, {"\"": "&quot;", "'": "&apos;"})


def unescape_xml_text(value: str) -> str:
    """Decode XML/HTML character references in a value read from XML text."""

    return html.unescape(value)


def column_index_to_letter(index: int) -> str:
    """Convert a zero-based column index to an Excel column name."""

    if index < 0:
        raise ValueError("Column index cannot be negative")
    result: list[str] = []
    value = index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def column_letter_to_index(reference: str) -> int:
    """Extract a zero-based column index from an A1-style cell reference."""

    letters = "".join(character for character in reference if character.isalpha())
    if not letters:
        raise ValueError(f"Invalid cell reference: {reference!r}")
    result = 0
    for character in letters.upper():
        if not "A" <= character <= "Z":
            raise ValueError(f"Invalid cell reference: {reference!r}")
        result = result * 26 + ord(character) - ord("A") + 1
    return result - 1


def local_xml_name(name: str) -> str:
    """Return an XML local name for either prefixed or namespace-expanded input."""

    if "}" in name:
        name = name.rsplit("}", 1)[1]
    if ":" in name:
        name = name.rsplit(":", 1)[1]
    return name


def parse_shared_strings_xml(xml_data: bytes | str) -> list[str]:
    """Parse plain and rich-text ``<si>`` entries from an XLSX SST.

    Expat performs entity decoding for us and this parser intentionally joins
    all ``<t>`` runs inside one shared-string item, which is what Excel
    displays for rich text.
    """

    if isinstance(xml_data, bytes):
        source = xml_data
    else:
        source = xml_data.encode("utf-8")

    values: list[str] = []
    in_shared_item = False
    in_text = False
    text_parts: list[str] = []

    parser = expat.ParserCreate()

    def start_element(name: str, attrs: dict[str, str]) -> None:
        nonlocal in_shared_item, in_text, text_parts
        local = local_xml_name(name)
        if local == "si":
            in_shared_item = True
            text_parts = []
        elif local == "t" and in_shared_item:
            in_text = True

    def end_element(name: str) -> None:
        nonlocal in_shared_item, in_text
        local = local_xml_name(name)
        if local == "t":
            in_text = False
        elif local == "si" and in_shared_item:
            values.append("".join(text_parts))
            in_shared_item = False

    def character_data(data: str) -> None:
        if in_shared_item and in_text:
            text_parts.append(data)

    parser.StartElementHandler = start_element
    parser.EndElementHandler = end_element
    parser.CharacterDataHandler = character_data
    parser.Parse(source, True)
    return values


_XML_ATTRIBUTE_RE_TEMPLATE = r"(?P<prefix>(?:[A-Za-z_][\w.-]*:)?{name})\s*=\s*(?P<quote>[\"\'])(?P<value>.*?)(?P=quote)"


def replace_or_add_xml_attribute(opening_tag: str, name: str, value: Any) -> str:
    """Replace an attribute in an opening tag, or add it before its closing marker."""

    escaped = escape_xml_text(value)
    pattern = re.compile(_XML_ATTRIBUTE_RE_TEMPLATE.format(name=re.escape(name)), re.DOTALL)
    match = pattern.search(opening_tag)
    if match:
        start, end = match.span("value")
        return opening_tag[:start] + escaped + opening_tag[end:]

    insertion = f' {name}="{escaped}"'
    if opening_tag.endswith("/>"):
        return opening_tag[:-2] + insertion + "/>"
    return opening_tag[:-1] + insertion + ">"


def parse_relationships_xml(xml_data: bytes) -> dict[str, str]:
    """Read relationship IDs and targets from an OOXML relationships part."""

    relationships: dict[str, str] = {}
    parser = expat.ParserCreate()

    def start_element(name: str, attrs: dict[str, str]) -> None:
        if local_xml_name(name) != "Relationship":
            return
        relationship_id = attrs.get("Id")
        target = attrs.get("Target")
        if relationship_id and target:
            relationships[relationship_id] = target

    parser.StartElementHandler = start_element
    parser.Parse(xml_data, True)
    return relationships


def resolve_zip_target(target: str, base_directory: str = "xl") -> str:
    """Resolve an OOXML relationship target to a ZIP member name."""

    if target.startswith("/"):
        resolved = target[1:]
    else:
        resolved = posixpath.normpath(posixpath.join(base_directory, target))
    return resolved.replace("\\", "/")


class ZipPackage:
    """In-memory ZIP package that preserves all non-target payloads exactly."""

    def __init__(self, path: str | os.PathLike[str]):
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(self.path)
        if not self.path.is_file():
            raise ValueError(f"Expected a file: {self.path}")

        self._entries: list[tuple[zipfile.ZipInfo, bytes]] = []
        self._indices: dict[str, int] = {}
        with zipfile.ZipFile(self.path, "r") as archive:
            for info in archive.infolist():
                copied_info = copy.copy(info)
                self._entries.append((copied_info, archive.read(info)))
                self._indices[info.filename] = len(self._entries) - 1

    def has(self, name: str) -> bool:
        return name in self._indices

    def names(self) -> list[str]:
        """Return current ZIP member names in archive order."""

        return [info.filename for info, _ in self._entries]

    def get(self, name: str) -> bytes:
        try:
            return self._entries[self._indices[name]][1]
        except KeyError as exc:
            raise KeyError(f"ZIP member not found: {name}") from exc

    def set(self, name: str, data: bytes) -> None:
        if name in self._indices:
            index = self._indices[name]
            info, _ = self._entries[index]
            self._entries[index] = (info, bytes(data))
            return

        info = zipfile.ZipInfo(name)
        info.compress_type = zipfile.ZIP_DEFLATED
        self._indices[name] = len(self._entries)
        self._entries.append((info, bytes(data)))

    def to_bytes(self) -> bytes:
        result = BytesIO()
        with zipfile.ZipFile(result, "w") as archive:
            for original_info, data in self._entries:
                info = copy.copy(original_info)
                archive.writestr(info, data)
        return result.getvalue()

    def save(self, output_path: str | os.PathLike[str] | None = None) -> None:
        """Write the package, using a same-directory temporary file and replace."""

        target = Path(output_path) if output_path is not None else self.path
        target = target.resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary_path: str | None = None
        descriptor, temporary_path = tempfile.mkstemp(
            prefix=f".{target.name}.", suffix=".tmp", dir=target.parent
        )
        os.close(descriptor)
        try:
            with open(temporary_path, "wb") as stream:
                stream.write(self.to_bytes())
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary_path, target)
            temporary_path = None
        finally:
            if temporary_path is not None:
                try:
                    os.unlink(temporary_path)
                except FileNotFoundError:
                    pass
