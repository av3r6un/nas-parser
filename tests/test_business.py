"""Tests for NAS Parser product enrichment business logic."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from nas_parser.business import ProductEnricher
from nas_parser.domain import ProductRecord
from nas_parser.references.colors import ColorReferenceLoader
from nas_parser.report import RunReport


def _write_color_reference_workbook(workbook_path: Path) -> None:
    """Create a small color reference workbook for tests."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(("Артикул WB и Ozon", "Грани", "Цвет", "Размер", "Тип", "Артикул цвета"))
    sheet.append(("12cut/Crystal/SS3/Hot/001", "12cut", "Crystal", "SS3", "Hot", "001"))
    sheet.append(("16cut/Crystal AB/SS30/Non/001+", "16cut", "Crystal AB", "SS30", "Non", "001+"))
    sheet.append(("16cut/New AB/SS16/Hot/112", "16cut", "New AB", "SS16", "Hot", "112"))
    workbook.save(workbook_path)
    workbook.close()


def _make_record(
    *,
    price: Decimal | None,
    quantity: Decimal | None,
    color: str,
    size: str,
    cut: str | None,
    fixation: str | None,
    source_file: Path,
    source_sheet: str,
    source_row: int,
    color_code: str | None = None,
    shape: str | None = None,
    parser_name: str = "cut",
) -> ProductRecord:
    """Build a ProductRecord fixture for enrichment tests."""
    return ProductRecord(
        price=price,
        quantity=quantity,
        color=color,
        size=size,
        source_file=source_file,
        source_sheet=source_sheet,
        source_row=source_row,
        parser_name=parser_name,
        cut=cut,
        fixation=fixation,
        color_code=color_code,
        shape=shape,
    )


class TestProductEnricher(unittest.TestCase):
    """Coverage for ProductEnricher behavior."""

    def test_enrich_builds_business_fields(self) -> None:
        """Verify that the enricher builds SKU, name, category and color codes."""
        with tempfile.TemporaryDirectory() as temp_dir:
            reference_file = Path(temp_dir) / "colorcode-articul.xlsx"
            _write_color_reference_workbook(reference_file)

            color_reference = ColorReferenceLoader(reference_file).load()
            enricher = ProductEnricher(color_reference)
            report = RunReport()
            records = [
                _make_record(
                    price=Decimal("135"),
                    quantity=Decimal("374"),
                    color="CRYSTAL",
                    size="SS16",
                    cut="12cut",
                    fixation="hot",
                    source_file=Path("input/12cut_hot.xlsx"),
                    source_sheet="12cut_hot",
                    source_row=3,
                ),
                _make_record(
                    price=Decimal("211"),
                    quantity=Decimal("167"),
                    color="Crystal AB",
                    size="SS30",
                    cut="16cut",
                    fixation="non",
                    source_file=Path("input/16cut_non.xlsx"),
                    source_sheet="16cut_non",
                    source_row=7,
                ),
                _make_record(
                    price=Decimal("175"),
                    quantity=Decimal("122"),
                    color="NEW AB",
                    size="SS16",
                    cut="16cut",
                    fixation="hot",
                    source_file=Path("input/16cut_hot.xlsx"),
                    source_sheet="16cut_hot",
                    source_row=10,
                ),
            ]

            enriched = enricher.enrich(records, report)

        self.assertEqual(len(enriched), 3)
        self.assertEqual(enriched[0].color_code, "001")
        self.assertEqual(enriched[0].sku, "12cut/Crystal/SS16/Hot/001")
        self.assertEqual(
            enriched[0].name, "Стразы 12 граней Crystal SS16 Hotfix"
        )
        self.assertEqual(enriched[0].category, "Стразы горячей фиксации")

        self.assertEqual(enriched[1].color_code, "001+")
        self.assertEqual(enriched[1].sku, "16cut/Crystal Ab/SS30/Non/001+")
        self.assertEqual(
            enriched[1].name, "Стразы 16 граней Crystal AB SS30 Non Hotfix"
        )
        self.assertEqual(enriched[1].category, "Стразы холодной фиксации")

        self.assertEqual(enriched[2].color_code, "112")
        self.assertEqual(enriched[2].sku, "16cut/New Ab/SS16/Hot/112")
        self.assertEqual(enriched[2].name, "Стразы 16 граней New AB SS16 Hotfix")
        self.assertEqual(enriched[2].category, "Стразы горячей фиксации")
        self.assertIn("warnings=0", report.summary())
        self.assertIn("errors=0", report.summary())

    def test_enrich_warns_about_missing_business_data(self) -> None:
        """Verify that enrichment keeps processing even when data is incomplete."""
        with tempfile.TemporaryDirectory() as temp_dir:
            reference_file = Path(temp_dir) / "colorcode-articul.xlsx"
            _write_color_reference_workbook(reference_file)

            color_reference = ColorReferenceLoader(reference_file).load()
            enricher = ProductEnricher(color_reference)
            report = RunReport()
            record = _make_record(
                price=None,
                quantity=None,
                color="",
                size="",
                cut="12cut",
                fixation="hot",
                source_file=Path("input/12cut_hot.xlsx"),
                source_sheet="12cut_hot",
                source_row=99,
            )

            enriched = enricher.enrich([record], report)

        self.assertEqual(len(enriched), 1)
        self.assertIsNone(enriched[0].sku)
        self.assertIsNone(enriched[0].name)
        self.assertEqual(enriched[0].category, "Стразы горячей фиксации")
        self.assertIn("warnings=0", report.summary())
        self.assertIn("errors=0", report.summary())
