"""Tests for NAS Parser reader contracts."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
import unittest

from nas_parser.readers import ReaderBase
from nas_parser.source import SourceRow


class DummyReader(ReaderBase):
    """Minimal reader implementation used to verify the contract."""

    @property
    def name(self) -> str:
        """Return the reader name."""
        return "dummy"

    def read(self) -> Iterable[SourceRow]:
        """Return a single in-memory source row."""
        yield SourceRow(
            values=("Topaz", 12, 3.5),
            source_file=Path("input/dummy.xlsx"),
            source_sheet="Sheet1",
            source_row=1,
        )


class TestReaderBase(unittest.TestCase):
    """Coverage for the reader base contract."""

    def test_reader_base_cannot_be_instantiated(self) -> None:
        """Verify that ReaderBase stays abstract."""
        with self.assertRaises(TypeError):
            ReaderBase()  # type: ignore[abstract]

    def test_concrete_reader_returns_source_rows(self) -> None:
        """Verify that a concrete reader yields SourceRow instances."""
        reader = DummyReader()
        rows = list(reader.read())

        self.assertEqual(reader.name, "dummy")
        self.assertEqual(len(rows), 1)
        self.assertIsInstance(rows[0], SourceRow)
        self.assertIsInstance(rows, Iterable)
        self.assertEqual(rows[0].source_sheet, "Sheet1")
