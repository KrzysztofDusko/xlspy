"""Small, format-neutral helpers used by the existing-file updaters.

The readers and writers in :mod:`xlspy` deliberately have very different
implementations for XML and BIFF12.  The updater layer only needs a few
shared operations: materialising a row iterable, manipulating ZIP packages,
and generating safe XML text.
"""

from __future__ import annotations

import copy
import html
import mmap
import os
import posixpath
import re
import shutil
import sqlite3
import struct
import tempfile
import zipfile
from contextlib import contextmanager
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable, Sequence
from xml.parsers import expat
from xml.sax import saxutils


STREAM_BUFFER_SIZE = 1024 * 1024


def copy_stream(source, target, *, length: int = STREAM_BUFFER_SIZE) -> None:
    """Copy a binary stream with a bounded buffer."""

    shutil.copyfileobj(source, target, length=length)


def copy_file_range(source_path: str | os.PathLike[str], target, start: int, end: int) -> None:
    """Copy a byte range from a seekable file to an open binary stream."""

    with open(source_path, "rb") as source:
        source.seek(start)
        remaining = end - start
        while remaining > 0:
            chunk = source.read(min(STREAM_BUFFER_SIZE, remaining))
            if not chunk:
                raise ValueError("Unexpected end of temporary file")
            target.write(chunk)
            remaining -= len(chunk)


def xml_space_attribute(value: str) -> str:
    """Return the XML space-preservation attribute used for text nodes."""

    preserve = (
        value[:1].isspace()
        or value[-1:].isspace()
        or any(character in value for character in "\t\n\r")
    )
    return ' xml:space="preserve"' if preserve else ""


class DiskStringIndex:
    """Disk-backed shared-string lookup with an append-only pending file."""

    def __init__(self, directory: str | os.PathLike[str], *, total_count: int = 0):
        directory = Path(directory)
        self._database = sqlite3.connect(directory / "shared-strings.sqlite3")
        self._database.execute(
            "CREATE TABLE IF NOT EXISTS strings (value TEXT PRIMARY KEY, idx INTEGER NOT NULL)"
        )
        self._database.commit()
        self._pending_path = directory / "shared-strings.pending"
        self._pending = self._pending_path.open("ab+")
        self.total_count = total_count
        self.unique_count = 0
        self.dirty = False
        self._uncommitted_inserts = 0
        self._transaction: dict[str, Any] | None = None

    def seed(self, values: Iterable[str]) -> None:
        for index, value in enumerate(values):
            self.add_existing(value, index)
        self._database.commit()

    def add_existing(self, value: str, index: int) -> None:
        """Add one original SST value using its stable zero-based index."""

        # The in-memory reference implementation uses a dict comprehension,
        # so duplicate shared strings resolve to their last index.
        self._database.execute(
            "INSERT OR REPLACE INTO strings(value, idx) VALUES (?, ?)",
            (value, index),
        )
        self.unique_count = max(self.unique_count, index + 1)

    def get_or_add(self, value: str) -> int:
        result = self._database.execute(
            "SELECT idx FROM strings WHERE value = ?", (value,)
        ).fetchone()
        if result is None:
            index = self.unique_count
            self._database.execute(
                "INSERT INTO strings(value, idx) VALUES (?, ?)", (value, index)
            )
            self._uncommitted_inserts += 1
            if self._uncommitted_inserts >= 1024:
                self._database.commit()
                self._uncommitted_inserts = 0
            encoded = value.encode("utf-8")
            self._pending.write(struct.pack("<Q", len(encoded)))
            self._pending.write(encoded)
            self._pending.flush()
            self.unique_count += 1
        else:
            index = int(result[0])
        self.total_count += 1
        self.dirty = True
        return index

    def pending_values(self):
        self._pending.flush()
        with self._pending_path.open("rb") as stream:
            while True:
                header = stream.read(8)
                if not header:
                    return
                if len(header) != 8:
                    raise ValueError("Malformed pending shared-string file")
                length = struct.unpack("<Q", header)[0]
                encoded = stream.read(length)
                if len(encoded) != length:
                    raise ValueError("Truncated pending shared-string file")
                yield encoded.decode("utf-8")

    def clear_pending(self) -> None:
        self._database.commit()
        self._uncommitted_inserts = 0
        self._pending.seek(0)
        self._pending.truncate()
        self._pending.flush()
        self.dirty = False

    def flush(self) -> None:
        """Commit pending lookup-index inserts without clearing new values."""

        if self._uncommitted_inserts:
            self._database.commit()
            self._uncommitted_inserts = 0

    def begin_transaction(self) -> None:
        """Start a disk-backed transaction for one streamed replacement."""

        if self._transaction is not None:
            raise RuntimeError("A shared-string transaction is already active")

        self.flush()
        self._pending.flush()
        descriptor, backup_path = tempfile.mkstemp(
            prefix="shared-strings-", suffix=".backup", dir=self._pending_path.parent
        )
        os.close(descriptor)
        backup = Path(backup_path)
        try:
            with self._pending_path.open("rb") as source, backup.open("wb") as target:
                copy_stream(source, target)
        except BaseException:
            try:
                backup.unlink()
            except FileNotFoundError:
                pass
            raise

        self._transaction = {
            "total_count": self.total_count,
            "unique_count": self.unique_count,
            "dirty": self.dirty,
            "uncommitted_inserts": self._uncommitted_inserts,
            "pending_backup": backup,
        }

    def commit_transaction(self) -> None:
        """Commit the current streamed replacement transaction."""

        transaction = self._transaction
        if transaction is None:
            return
        self._database.commit()
        backup = transaction["pending_backup"]
        try:
            backup.unlink()
        except FileNotFoundError:
            pass
        self._transaction = None

    def rollback_transaction(self) -> None:
        """Restore the index and pending values to the transaction snapshot."""

        transaction = self._transaction
        if transaction is None:
            return

        backup = transaction["pending_backup"]
        try:
            # New values always receive indexes at or above the old unique
            # count, so no in-memory list of inserted strings is needed.
            self._database.execute(
                "DELETE FROM strings WHERE idx >= ?",
                (transaction["unique_count"],),
            )
            self._database.commit()
            self._uncommitted_inserts = 0

            self._pending.seek(0)
            self._pending.truncate()
            with backup.open("rb") as source:
                copy_stream(source, self._pending)
            self._pending.flush()

            self.total_count = transaction["total_count"]
            self.unique_count = transaction["unique_count"]
            self.dirty = transaction["dirty"]
        finally:
            try:
                backup.unlink()
            except FileNotFoundError:
                pass
            self._transaction = None

    def close(self) -> None:
        if self._transaction is not None:
            self.rollback_transaction()
        self._pending.close()
        self._database.close()


def rewrite_shared_strings_xml(
    source_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str],
    pending_values: Iterable[str],
    *,
    total_count: int,
    unique_count: int,
) -> None:
    """Rewrite an XLSX SST without materialising its XML payload."""

    with open(source_path, "rb") as source_stream:
        size = os.fstat(source_stream.fileno()).st_size
        if size == 0:
            raise ValueError("Shared strings XML is empty")
        with mmap.mmap(source_stream.fileno(), 0, access=mmap.ACCESS_READ) as source:
            opening_match = re.search(
                rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?sst\b[^>]*>", source
            )
            closing_match = re.search(
                rb"</(?:[A-Za-z_][\w.-]*:)?sst\s*>\s*$", source
            )
            if not opening_match or not closing_match:
                raise ValueError("Malformed XLSX sharedStrings.xml")

            opening = opening_match.group(0).decode("utf-8")
            opening = replace_or_add_xml_attribute(opening, "count", total_count)
            opening = replace_or_add_xml_attribute(opening, "uniqueCount", unique_count)
            prefix = re.match(r"<(?P<prefix>[A-Za-z_][\w.-]*:)?sst\b", opening).group("prefix") or ""

            with open(output_path, "wb") as target:
                copy_file_range(source_path, target, 0, opening_match.start())
                target.write(opening.encode("utf-8"))
                copy_file_range(source_path, target, opening_match.end(), closing_match.start())
                for value in pending_values:
                    target.write(
                        (
                            f"<{prefix}si><{prefix}t{xml_space_attribute(value)}>"
                            f"{escape_xml_text(value)}</{prefix}t></{prefix}si>"
                        ).encode("utf-8")
                    )
                copy_file_range(source_path, target, closing_match.start(), size)


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
    """Lazy ZIP package used by both the regular and streaming updaters.

    The original updater API still exposes byte-oriented ``get``/``set`` and
    ``to_bytes`` methods.  Entries are now read on demand, while the optional
    file-backed replacements are used by the streaming updater so a workbook
    is not materialised in memory.
    """

    def __init__(self, path: str | os.PathLike[str]):
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(self.path)
        if not self.path.is_file():
            raise ValueError(f"Expected a file: {self.path}")

        self._entries: list[zipfile.ZipInfo] = []
        self._indices: dict[str, int] = {}
        with zipfile.ZipFile(self.path, "r") as archive:
            for info in archive.infolist():
                copied_info = copy.copy(info)
                self._entries.append(copied_info)
                self._indices[info.filename] = len(self._entries) - 1
        self._overrides: dict[int, bytes | Path] = {}
        self._temporary_directory = tempfile.TemporaryDirectory(prefix="xlspy-updater-")
        self._package_transaction: tuple[
            list[zipfile.ZipInfo],
            dict[str, int],
            dict[int, bytes | Path],
        ] | None = None
        self._superseded_paths: set[Path] = set()

    def has(self, name: str) -> bool:
        return name in self._indices

    def names(self) -> list[str]:
        """Return current ZIP member names in archive order."""

        return [info.filename for info in self._entries]

    def _index_for(self, name: str) -> int:
        try:
            return self._indices[name]
        except KeyError as exc:
            raise KeyError(f"ZIP member not found: {name}") from exc

    def get(self, name: str) -> bytes:
        index = self._index_for(name)
        override = self._overrides.get(index)
        if isinstance(override, bytes):
            return override
        if isinstance(override, Path):
            return override.read_bytes()
        with zipfile.ZipFile(self.path, "r") as archive:
            return archive.read(self._entries[index])

    @contextmanager
    def open_entry(self, name: str):
        """Open one entry for bounded-memory reading."""

        index = self._index_for(name)
        override = self._overrides.get(index)
        if isinstance(override, bytes):
            stream = BytesIO(override)
            try:
                yield stream
            finally:
                stream.close()
            return
        if isinstance(override, Path):
            with override.open("rb") as stream:
                yield stream
            return

        with zipfile.ZipFile(self.path, "r") as archive:
            with archive.open(self._entries[index], "r") as stream:
                yield stream

    def temporary_path(self, *, suffix: str = "") -> Path:
        """Create a named temporary file owned by this package."""

        descriptor, path = tempfile.mkstemp(
            prefix="part-", suffix=suffix, dir=self._temporary_directory.name
        )
        os.close(descriptor)
        return Path(path)

    @property
    def temporary_directory(self) -> Path:
        """Return the private directory used for staged package parts."""

        return Path(self._temporary_directory.name)

    def begin_transaction(self) -> None:
        """Start a transaction for staged package replacements."""

        if self._package_transaction is not None:
            raise RuntimeError("A ZIP package transaction is already active")
        self._package_transaction = (
            self._entries.copy(),
            self._indices.copy(),
            self._overrides.copy(),
        )
        self._superseded_paths = set()

    def _is_owned_staged_path(self, path: Path) -> bool:
        try:
            path.resolve().relative_to(self.temporary_directory.resolve())
        except (OSError, ValueError):
            return False
        return True

    def _delete_staged_path(self, path: Path) -> None:
        if not self._is_owned_staged_path(path):
            return
        try:
            path.unlink()
        except FileNotFoundError:
            pass

    def _set_override(self, index: int, replacement: bytes | Path) -> None:
        previous = self._overrides.get(index)
        if isinstance(previous, Path) and previous != replacement:
            if self._package_transaction is None:
                self._delete_staged_path(previous)
            else:
                self._superseded_paths.add(previous)
        self._overrides[index] = replacement

    def commit_transaction(self) -> None:
        """Commit staged replacements and remove superseded temporary files."""

        if self._package_transaction is None:
            return
        current_paths = {
            override
            for override in self._overrides.values()
            if isinstance(override, Path)
        }
        for path in self._superseded_paths:
            if path not in current_paths:
                self._delete_staged_path(path)
        self._package_transaction = None
        self._superseded_paths = set()

    def rollback_transaction(self) -> None:
        """Discard staged changes made since the transaction began."""

        transaction = self._package_transaction
        if transaction is None:
            return
        _, _, original_overrides = transaction
        original_paths = {
            override
            for override in original_overrides.values()
            if isinstance(override, Path)
        }
        for override in self._overrides.values():
            if isinstance(override, Path) and override not in original_paths:
                self._delete_staged_path(override)
        self._entries, self._indices, self._overrides = transaction
        self._package_transaction = None
        self._superseded_paths = set()

    def set(self, name: str, data: bytes) -> None:
        if name in self._indices:
            index = self._indices[name]
            self._set_override(index, bytes(data))
            return

        info = zipfile.ZipInfo(name)
        info.compress_type = zipfile.ZIP_DEFLATED
        self._indices[name] = len(self._entries)
        self._entries.append(info)
        self._set_override(len(self._entries) - 1, bytes(data))

    def set_file(self, name: str, path: str | os.PathLike[str]) -> None:
        """Use a file as the replacement payload for one ZIP member."""

        replacement = Path(path)
        if not replacement.is_file():
            raise ValueError(f"Expected a replacement file: {replacement}")
        if name in self._indices:
            self._set_override(self._indices[name], replacement)
            return

        info = zipfile.ZipInfo(name)
        info.compress_type = zipfile.ZIP_DEFLATED
        self._indices[name] = len(self._entries)
        self._entries.append(info)
        self._set_override(len(self._entries) - 1, replacement)

    @contextmanager
    def _replacement_stream(self, index: int):
        override = self._overrides.get(index)
        if isinstance(override, bytes):
            with BytesIO(override) as stream:
                yield stream
            return
        if isinstance(override, Path):
            with override.open("rb") as stream:
                yield stream
            return
        with zipfile.ZipFile(self.path, "r") as archive:
            with archive.open(self._entries[index], "r") as stream:
                yield stream

    def to_bytes(self) -> bytes:
        result = BytesIO()
        with zipfile.ZipFile(result, "w") as archive:
            for index, original_info in enumerate(self._entries):
                with self._replacement_stream(index) as stream:
                    with archive.open(copy.copy(original_info), "w") as target:
                        shutil.copyfileobj(stream, target, length=1024 * 1024)
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
            with zipfile.ZipFile(temporary_path, "w") as archive:
                for index, original_info in enumerate(self._entries):
                    with self._replacement_stream(index) as stream:
                        with archive.open(copy.copy(original_info), "w") as target_stream:
                            shutil.copyfileobj(stream, target_stream, length=1024 * 1024)
                archive.comment = b""
            with open(temporary_path, "rb+") as stream:
                os.fsync(stream.fileno())
            os.replace(temporary_path, target)
            temporary_path = None
            if target == self.path.resolve():
                self._reload_entries(target)
        finally:
            if temporary_path is not None:
                try:
                    os.unlink(temporary_path)
                except FileNotFoundError:
                    pass

    def _reload_entries(self, source_path: str | os.PathLike[str]) -> None:
        """Refresh ZIP metadata after an archive has been replaced in place."""

        entries: list[zipfile.ZipInfo] = []
        indices: dict[str, int] = {}
        with zipfile.ZipFile(source_path, "r") as archive:
            for info in archive.infolist():
                entries.append(copy.copy(info))
                indices[info.filename] = len(entries) - 1
        self._entries = entries
        self._indices = indices
