"""Tests for NAS Parser domain models."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path
import unittest

from nas_parser.domain import ProductRecord


class TestProductRecord(unittest.TestCase):
    """Coverage for the ProductRecord domain model."""

    def test_product_record_can_be_created(self) -> None:
        """Verify that ProductRecord stores required and optional product data."""
        record = ProductRecord(
            price=Decimal("12.50"),
            quantity=Decimal("3.5"),
            color="Topaz",
            size="SS16",
            source_file=Path("input/sample.xlsx"),
            source_sheet="Sheet1",
            source_row=7,
            parser_name="test_parser",
        )

        self.assertEqual(record.price, Decimal("12.50"))
        self.assertEqual(record.quantity, Decimal("3.5"))
        self.assertEqual(record.color, "Topaz")
        self.assertEqual(record.size, "SS16")
        self.assertEqual(record.source_file, Path("input/sample.xlsx"))
        self.assertEqual(record.source_sheet, "Sheet1")
        self.assertEqual(record.source_row, 7)
        self.assertEqual(record.parser_name, "test_parser")
        self.assertIsNone(record.sku)
        self.assertIsNone(record.name)
        self.assertIsNone(record.category)
        self.assertIsNone(record.color_code)
        self.assertIsNone(record.shape)
        self.assertIsNone(record.fixation)
        self.assertIsNone(record.cut)

    def test_source_file_is_hidden_from_repr(self) -> None:
        """Verify that ProductRecord repr does not include the source file path."""
        record = ProductRecord(
            price=Decimal("1"),
            quantity=Decimal("5"),
            color="Topaz",
            size="SS16",
            source_file=Path("input/private/source.xlsx"),
            source_sheet="Sheet1",
            source_row=1,
            parser_name="test_parser",
        )

        self.assertNotIn("source_file", repr(record))
        self.assertNotIn("input/private/source.xlsx", repr(record))
