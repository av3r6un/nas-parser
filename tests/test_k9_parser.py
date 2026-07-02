"""Tests for NAS Parser K9 parser."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path
import unittest

from nas_parser.domain import ProductRecord
from nas_parser.parsers.k9 import K9Parser
from nas_parser.source import SourceRow


def make_source_row(
    source_file: Path,
    source_sheet: str,
    source_row: int,
    values: tuple[object, ...],
) -> SourceRow:
    """Create a SourceRow fixture for K9 parser tests."""
    return SourceRow(
        values=values,
        source_file=source_file,
        source_sheet=source_sheet,
        source_row=source_row,
    )


class TestK9Parser(unittest.TestCase):
    """Coverage for the K9Parser implementation."""

    def test_supports_k9_files(self) -> None:
        """Verify that the parser supports K9 files."""
        parser = K9Parser()

        self.assertTrue(parser.supports(Path("input/K9.xlsx")))
        self.assertTrue(parser.supports(Path("input/my_k9_prices.xlsx")))

    def test_supports_rejects_non_k9_files(self) -> None:
        """Verify that the parser rejects unrelated file types."""
        parser = K9Parser()

        self.assertFalse(parser.supports(Path("input/12cut_hot.xlsx")))

    def test_parse_creates_product_records_from_k9_rows(self) -> None:
        """Verify that K9 rows are converted into ProductRecord objects."""
        parser = K9Parser()
        source_file = Path("input/K9.xlsx")
        rows = [
            make_source_row(
                source_file,
                "sheet1",
                1,
                ("K9 title", None, None, None, None, None, None, None),
            ),
            make_source_row(
                source_file,
                "sheet1",
                2,
                ("color", "shape", "size", "price", "pack", "received", "sold", "stock"),
            ),
            make_source_row(
                source_file,
                "sheet1",
                3,
                ("CRYSTAL", "Drop", "7*12", 11, 72, 1886, None, 1886),
            ),
            make_source_row(
                source_file,
                "sheet1",
                4,
                (None, None, "10*14", 13, 36, 1170, None, 1170),
            ),
            make_source_row(
                source_file,
                "sheet1",
                5,
                (None, "Navette", "7*15", 20, 63, 439, None, 439),
            ),
            make_source_row(
                source_file,
                "sheet1",
                6,
                ("AQUAMARINE", None, "8*8", 7, 84, 651, None, 651),
            ),
            make_source_row(
                source_file,
                "sheet1",
                7,
                (None, None, "9*18", 8, 45, 954, None, 954),
            ),
            make_source_row(
                source_file,
                "sheet1",
                8,
                (None, None, None, 9, 12, 100, None, 100),
            ),
            make_source_row(
                source_file,
                "sheet2",
                1,
                ("Another title", None, None, None, None, None, None, None),
            ),
            make_source_row(
                source_file,
                "sheet2",
                2,
                ("color", "shape", "size", "price", "pack", "received", "sold", "stock"),
            ),
            make_source_row(
                source_file,
                "sheet2",
                3,
                ("TOPAZ", "Rivoli", "12*12", 5, 10, 20, None, 20),
            ),
        ]

        records = list(parser.parse(rows))

        self.assertEqual(len(records), 6)
        self._assert_record(
            records[0],
            price=11,
            quantity=1886,
            color="CRYSTAL",
            shape="Drop",
            size="7*12",
            source_file=source_file,
            source_sheet="sheet1",
            source_row=3,
        )
        self._assert_record(
            records[1],
            price=13,
            quantity=1170,
            color="CRYSTAL",
            shape="Drop",
            size="10*14",
            source_file=source_file,
            source_sheet="sheet1",
            source_row=4,
        )
        self._assert_record(
            records[2],
            price=20,
            quantity=439,
            color="CRYSTAL",
            shape="Navette",
            size="7*15",
            source_file=source_file,
            source_sheet="sheet1",
            source_row=5,
        )
        self._assert_record(
            records[3],
            price=7,
            quantity=651,
            color="AQUAMARINE",
            shape="Navette",
            size="8*8",
            source_file=source_file,
            source_sheet="sheet1",
            source_row=6,
        )
        self._assert_record(
            records[4],
            price=8,
            quantity=954,
            color="AQUAMARINE",
            shape="Navette",
            size="9*18",
            source_file=source_file,
            source_sheet="sheet1",
            source_row=7,
        )
        self._assert_record(
            records[5],
            price=5,
            quantity=20,
            color="TOPAZ",
            shape="Rivoli",
            size="12*12",
            source_file=source_file,
            source_sheet="sheet2",
            source_row=3,
        )

    def test_parse_skips_rows_without_size(self) -> None:
        """Verify that service rows without a size are ignored."""
        parser = K9Parser()
        source_file = Path("input/K9.xlsx")
        rows = [
            make_source_row(
                source_file,
                "sheet1",
                1,
                ("title", None, None, None, None, None, None, None),
            ),
            make_source_row(
                source_file,
                "sheet1",
                2,
                ("color", "shape", "size", "price", "pack", "received", "sold", "stock"),
            ),
            make_source_row(
                source_file,
                "sheet1",
                3,
                ("CRYSTAL", "Drop", None, 11, 72, 1886, None, 1886),
            ),
        ]

        records = list(parser.parse(rows))

        self.assertEqual(records, [])

    def test_parse_evaluates_formulas_and_mm_sizes(self) -> None:
        """Verify that K9 formula values are exported as numbers."""
        parser = K9Parser()
        source_file = Path("input/K9.xlsx")
        rows = [
            make_source_row(
                source_file,
                "sheet1",
                637,
                (
                    "BLUE ZIRCON ",
                    "RIVOLI 4041",
                    "10MM",
                    "=ROUND(12.7926,0)",
                    45,
                    "=5*180",
                    None,
                    "=F637-G637",
                ),
            ),
        ]

        records = list(parser.parse(rows))

        self.assertEqual(len(records), 1)
        self._assert_record(
            records[0],
            price=Decimal("13"),
            quantity=Decimal("900"),
            color="BLUE ZIRCON ",
            shape="RIVOLI 4041",
            size="10MM",
            source_file=source_file,
            source_sheet="sheet1",
            source_row=637,
        )

    def _assert_record(
        self,
        record: ProductRecord,
        *,
        price: object,
        quantity: object,
        color: object,
        shape: object,
        size: object,
        source_file: Path,
        source_sheet: str,
        source_row: int,
    ) -> None:
        """Assert the common fields of a parsed K9 product record."""
        self.assertEqual(record.price, price)
        self.assertEqual(record.quantity, quantity)
        self.assertEqual(record.color, color)
        self.assertEqual(record.shape, shape)
        self.assertEqual(record.size, size)
        self.assertEqual(record.fixation, "sew")
        self.assertIsNone(record.cut)
        self.assertIsNone(record.sku)
        self.assertIsNone(record.name)
        self.assertIsNone(record.category)
        self.assertIsNone(record.color_code)
        self.assertEqual(record.source_file, source_file)
        self.assertEqual(record.source_sheet, source_sheet)
        self.assertEqual(record.source_row, source_row)
        self.assertEqual(record.parser_name, "k9")
