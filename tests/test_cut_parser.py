"""Tests for NAS Parser cut parser."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path
import unittest

from nas_parser.domain import ProductRecord
from nas_parser.parsers.cut import CutParser
from nas_parser.source import SourceRow


def make_source_row(
    source_file: Path,
    source_sheet: str,
    source_row: int,
    values: tuple[object, ...],
) -> SourceRow:
    """Create a SourceRow fixture for parser tests."""
    return SourceRow(
        values=values,
        source_file=source_file,
        source_sheet=source_sheet,
        source_row=source_row,
    )


class TestCutParser(unittest.TestCase):
    """Coverage for the CutParser implementation."""

    def test_supports_12cut_files(self) -> None:
        """Verify that the parser supports 12cut files."""
        parser = CutParser()

        self.assertTrue(parser.supports(Path("input/12cut_hot.xlsx")))

    def test_supports_16cut_files(self) -> None:
        """Verify that the parser supports 16cut files."""
        parser = CutParser()

        self.assertTrue(parser.supports(Path("input/16cut_non.xlsx")))

    def test_supports_rejects_k9_files(self) -> None:
        """Verify that the parser rejects unrelated file types."""
        parser = CutParser()

        self.assertFalse(parser.supports(Path("input/K9.xlsx")))

    def test_parse_creates_product_records_from_12cut_rows(self) -> None:
        """Verify that 12cut source rows are converted into ProductRecord objects."""
        parser = CutParser()
        source_file = Path("input/12cut_hot.xlsx")
        rows = [
            make_source_row(
                source_file,
                "12 граней горячие",
                1,
                ("12 граней горячая фиксация", None, None, None, None, None, None, None),
            ),
            make_source_row(
                source_file,
                "12 граней горячие",
                2,
                ("цвет", "размер", "штук в пачке", "цена продажи в руб", "всего шт пришло", "продано", "остаток на складе", "код цвета"),
            ),
            make_source_row(
                source_file,
                "12 граней горячие",
                3,
                ("Crystal ", "SS3", 1440, 135, 516, 142, 374, "001"),
            ),
            make_source_row(
                source_file,
                "12 граней горячие",
                4,
                (None, "SS4", 1440, 135, 600, 63, 537, None),
            ),
            make_source_row(
                source_file,
                "12 граней горячие",
                5,
                ("Crystal AB классические ", "SS3", 1440, 96, 60, 57, 3, "001+"),
            ),
        ]

        records = list(parser.parse(rows))

        self.assertEqual(len(records), 3)
        self._assert_product_record(
            records[0],
            price=Decimal("135"),
            quantity=Decimal("374"),
            color="Crystal",
            color_code="001",
            size="SS3",
            cut="12cut",
            fixation="hot",
            source_file=source_file,
            source_sheet="12 граней горячие",
            source_row=3,
        )
        self._assert_product_record(
            records[1],
            price=Decimal("135"),
            quantity=Decimal("537"),
            color="Crystal",
            color_code="001",
            size="SS4",
            cut="12cut",
            fixation="hot",
            source_file=source_file,
            source_sheet="12 граней горячие",
            source_row=4,
        )
        self._assert_product_record(
            records[2],
            price=Decimal("96"),
            quantity=Decimal("3"),
            color="Crystal AB классические",
            color_code="001+",
            size="SS3",
            cut="12cut",
            fixation="hot",
            source_file=source_file,
            source_sheet="12 граней горячие",
            source_row=5,
        )

    def test_parse_creates_product_records_from_16cut_rows(self) -> None:
        """Verify that 16cut source rows are converted into ProductRecord objects."""
        parser = CutParser()
        source_file = Path("input/16cut_non.xlsx")
        rows = [
            make_source_row(
                source_file,
                "16 граней холодные",
                1,
                ("16 граней холодная фиксация", None, None, None, None, None, None, None),
            ),
            make_source_row(
                source_file,
                "16 граней холодные",
                2,
                ("цвет", "размер", "штук в пачке", "цена продажи в руб", "всего шт пришло", "продано", "остаток на складе", "код цвета"),
            ),
            make_source_row(
                source_file,
                "16 граней холодные",
                3,
                ("CRYSTAL", "SS6", 1440, 590, 57, 1, 56, "001"),
            ),
            make_source_row(
                source_file,
                "16 граней холодные",
                4,
                (None, "SS8", 1440, 701, 65, 2, 63, None),
            ),
        ]

        records = list(parser.parse(rows))

        self.assertEqual(len(records), 2)
        self._assert_product_record(
            records[0],
            price=Decimal("590"),
            quantity=Decimal("56"),
            color="CRYSTAL",
            color_code="001",
            size="SS6",
            cut="16cut",
            fixation="non",
            source_file=source_file,
            source_sheet="16 граней холодные",
            source_row=3,
        )
        self._assert_product_record(
            records[1],
            price=Decimal("701"),
            quantity=Decimal("63"),
            color="CRYSTAL",
            color_code="001",
            size="SS8",
            cut="16cut",
            fixation="non",
            source_file=source_file,
            source_sheet="16 граней холодные",
            source_row=4,
        )

    def test_parse_reads_article_and_color_code_from_new_columns(self) -> None:
        """Verify regular and mix rows resolve article and color code columns."""
        parser = CutParser()
        source_file = Path("input/12cut_non.xlsx")
        rows = [
            make_source_row(
                source_file,
                "12cut_non",
                1,
                ("title", None, None, None, None, None, None, None, None),
            ),
            make_source_row(
                source_file,
                "12cut_non",
                2,
                (
                    "color",
                    "size",
                    "pack_size",
                    "price",
                    "received",
                    "sold",
                    "stock",
                    "article",
                    "color_code",
                ),
            ),
            make_source_row(
                source_file,
                "12cut_non",
                3,
                (
                    "Crystal",
                    "SS16",
                    1440,
                    135,
                    516,
                    142,
                    374,
                    "12cut/Crystal/SS16/Non/001",
                    "001",
                ),
            ),
            make_source_row(
                source_file,
                "12cut_non",
                4,
                (
                    None,
                    "mix",
                    1440,
                    222,
                    600,
                    63,
                    537,
                    "MIX001",
                    "12cut/Crystal/mix/Non/MIX001",
                ),
            ),
        ]

        records = list(parser.parse(rows))

        self.assertEqual(len(records), 2)
        self._assert_product_record(
            records[0],
            price=Decimal("135"),
            quantity=Decimal("374"),
            color="Crystal",
            color_code="001",
            size="SS16",
            cut="12cut",
            fixation="non",
            source_file=source_file,
            source_sheet="12cut_non",
            source_row=3,
            sku="12cut/Crystal/SS16/Non/001",
        )
        self._assert_product_record(
            records[1],
            price=Decimal("222"),
            quantity=Decimal("537"),
            color="Crystal",
            color_code="MIX001",
            size="mix",
            cut="12cut",
            fixation="non",
            source_file=source_file,
            source_sheet="12cut_non",
            source_row=4,
            sku="12cut/Crystal/mix/Non/MIX001",
        )

    def _assert_product_record(
        self,
        record: ProductRecord,
        *,
        price: Decimal,
        quantity: Decimal,
        color: str,
        color_code: str | None,
        size: str,
        cut: str | None,
        fixation: str,
        source_file: Path,
        source_sheet: str,
        source_row: int,
        sku: str | None = None,
    ) -> None:
        """Assert the common fields of a parsed product record."""
        self.assertEqual(record.price, price)
        self.assertEqual(record.quantity, quantity)
        self.assertEqual(record.color, color)
        self.assertEqual(record.color_code, color_code)
        self.assertEqual(record.size, size)
        self.assertEqual(record.cut, cut)
        self.assertEqual(record.fixation, fixation)
        self.assertEqual(record.source_file, source_file)
        self.assertEqual(record.source_sheet, source_sheet)
        self.assertEqual(record.source_row, source_row)
        self.assertEqual(record.parser_name, "cut")
        self.assertEqual(record.sku, sku)
        self.assertIsNone(record.name)
        self.assertIsNone(record.category)
        self.assertIsNone(record.shape)
