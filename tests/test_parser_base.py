"""Tests for NAS Parser parser contracts."""

from __future__ import annotations

from collections.abc import Iterable
from decimal import Decimal
from pathlib import Path
import unittest

from nas_parser.domain import ProductRecord
from nas_parser.parsers import ParserBase
from nas_parser.source import SourceRow


class DummyParser(ParserBase):
    """Minimal parser implementation used to verify the contract."""

    @property
    def name(self) -> str:
        """Return the parser name."""
        return "dummy"

    def supports(self, source_file: Path) -> bool:
        """Return whether the dummy parser supports the given file."""
        return source_file.suffix == ".xlsx"

    def parse(self, source_rows: Iterable[SourceRow]) -> Iterable[ProductRecord]:
        """Return a single in-memory product record."""
        yield ProductRecord(
            price=Decimal("1"),
            quantity=Decimal("1"),
            color="Topaz",
            size="SS16",
            source_file=Path("input/dummy.xlsx"),
            source_sheet="Sheet1",
            source_row=1,
            parser_name=self.name,
        )


class TestParserBase(unittest.TestCase):
    """Coverage for the parser base contract."""

    def test_parser_base_cannot_be_instantiated(self) -> None:
        """Verify that ParserBase stays abstract."""
        with self.assertRaises(TypeError):
            ParserBase()  # type: ignore[abstract]

    def test_concrete_parser_returns_product_records(self) -> None:
        """Verify that a concrete parser yields ProductRecord instances."""
        parser = DummyParser()
        records = list(parser.parse([]))

        self.assertEqual(parser.name, "dummy")
        self.assertTrue(parser.supports(Path("input/sample.xlsx")))
        self.assertEqual(len(records), 1)
        self.assertIsInstance(records[0], ProductRecord)
        self.assertIsInstance(records, Iterable)
        self.assertEqual(records[0].parser_name, "dummy")
