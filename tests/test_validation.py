"""Tests for NAS Parser product validation."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path
import unittest

from nas_parser.domain import ProductRecord
from nas_parser.report import RunReport
from nas_parser.validation import ProductValidator


def _make_record(
    *,
    price: Decimal | None,
    quantity: Decimal | None,
    color: str,
    size: str,
    color_code: str | None,
    parser_name: str = "cut",
    cut: str | None = "12cut",
    fixation: str | None = "hot",
    shape: str | None = None,
    sku: str | None = "12cut/Crystal/SS3/Hot/001",
) -> ProductRecord:
    """Create a ProductRecord fixture for validator tests."""
    return ProductRecord(
        price=price,
        quantity=quantity,
        color=color,
        size=size,
        source_file=Path("input/sample.xlsx"),
        source_sheet="Sheet1",
        source_row=1,
        parser_name=parser_name,
        color_code=color_code,
        cut=cut,
        fixation=fixation,
        shape=shape,
        sku=sku,
        name="Стразы 12 граней Crystal SS3 Hotfix",
        category="Стразы горячей фиксации",
    )


class TestProductValidator(unittest.TestCase):
    """Coverage for ProductValidator behavior."""

    def test_missing_color_does_not_emit_warning(self) -> None:
        """Verify that missing color alone is not reported as a warning."""
        record = _make_record(
            price=Decimal("10"),
            quantity=Decimal("1"),
            color="",
            size="SS3",
            color_code="001",
        )
        report = RunReport()

        validated = ProductValidator().validate([record], report)

        self.assertEqual(validated, [record])
        self.assertIn("warnings=0", report.summary())
        self.assertIn("errors=0", report.summary())

    def test_missing_size_emits_warning(self) -> None:
        """Verify that missing size is reported as a warning."""
        record = _make_record(
            price=Decimal("10"),
            quantity=Decimal("1"),
            color="Crystal",
            size="",
            color_code="001",
        )
        report = RunReport()

        ProductValidator().validate([record], report)

        self.assertIn("warnings=1", report.summary())
        self.assertIn("errors=0", report.summary())

    def test_missing_color_code_emits_warning(self) -> None:
        """Verify that missing color_code is reported as a warning."""
        record = _make_record(
            price=Decimal("10"),
            quantity=Decimal("1"),
            color="Crystal",
            size="SS3",
            color_code=None,
        )
        report = RunReport()

        ProductValidator().validate([record], report)

        self.assertIn("warnings=1", report.summary())
        self.assertIn("errors=0", report.summary())

    def test_k9_missing_color_code_does_not_emit_warning(self) -> None:
        """Verify that K9 records do not require color_code."""
        record = _make_record(
            price=Decimal("10"),
            quantity=Decimal("1"),
            color="Crystal",
            size="7*12",
            color_code=None,
            parser_name="k9",
            cut=None,
            fixation="sew",
            shape="Drop",
            sku="Crystal/Drop/7x12",
        )
        report = RunReport()

        ProductValidator().validate([record], report)

        self.assertIn("warnings=0", report.summary())
        self.assertIn("errors=0", report.summary())

    def test_unable_to_build_sku_emits_warning(self) -> None:
        """Verify that a missing expected SKU is reported as a warning."""
        record = _make_record(
            price=Decimal("10"),
            quantity=Decimal("1"),
            color="Crystal",
            size="SS3",
            color_code="001",
            sku=None,
        )
        report = RunReport()

        ProductValidator().validate([record], report)

        self.assertIn("warnings=1", report.summary())
        self.assertIn("errors=0", report.summary())

    def test_missing_price_emits_warning(self) -> None:
        """Verify that missing price is reported as a warning."""
        record = _make_record(
            price=None,
            quantity=Decimal("1"),
            color="Crystal",
            size="SS3",
            color_code="001",
        )
        report = RunReport()

        ProductValidator().validate([record], report)

        self.assertIn("warnings=1", report.summary())
        self.assertIn("errors=0", report.summary())

    def test_missing_quantity_emits_warning(self) -> None:
        """Verify that missing quantity is reported as a warning."""
        record = _make_record(
            price=Decimal("10"),
            quantity=None,
            color="Crystal",
            size="SS3",
            color_code="001",
        )
        report = RunReport()

        ProductValidator().validate([record], report)

        self.assertIn("warnings=1", report.summary())
        self.assertIn("errors=0", report.summary())

    def test_negative_price_emits_error(self) -> None:
        """Verify that a negative price is reported as an error."""
        record = _make_record(
            price=Decimal("-1"),
            quantity=Decimal("1"),
            color="Crystal",
            size="SS3",
            color_code="001",
        )
        report = RunReport()

        ProductValidator().validate([record], report)

        self.assertIn("warnings=0", report.summary())
        self.assertIn("errors=1", report.summary())

    def test_negative_quantity_emits_error(self) -> None:
        """Verify that a negative quantity is reported as an error."""
        record = _make_record(
            price=Decimal("10"),
            quantity=Decimal("-1"),
            color="Crystal",
            size="SS3",
            color_code="001",
        )
        report = RunReport()

        ProductValidator().validate([record], report)

        self.assertIn("warnings=0", report.summary())
        self.assertIn("errors=1", report.summary())

    def test_validation_does_not_mutate_records(self) -> None:
        """Verify that validation leaves ProductRecord unchanged."""
        record = _make_record(
            price=Decimal("10"),
            quantity=Decimal("1"),
            color="Crystal",
            size="SS3",
            color_code="001",
        )
        original_state = (
            record.price,
            record.quantity,
            record.color,
            record.size,
            record.color_code,
            record.sku,
            record.name,
            record.category,
        )
        report = RunReport()

        validated = ProductValidator().validate([record], report)

        self.assertIs(validated[0], record)
        self.assertEqual(
            (
                record.price,
                record.quantity,
                record.color,
                record.size,
                record.color_code,
                record.sku,
                record.name,
                record.category,
            ),
            original_state,
        )

    def test_valid_record_passes_without_messages(self) -> None:
        """Verify that a fully valid record passes validation cleanly."""
        record = _make_record(
            price=Decimal("10"),
            quantity=Decimal("1"),
            color="Crystal",
            size="SS3",
            color_code="001",
        )
        report = RunReport()

        validated = ProductValidator().validate([record], report)

        self.assertEqual(validated, [record])
        self.assertIn("warnings=0", report.summary())
        self.assertIn("errors=0", report.summary())
