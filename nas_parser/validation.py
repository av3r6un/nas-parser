"""Validation layer for NAS Parser products."""

from __future__ import annotations

from collections.abc import Iterable
from decimal import Decimal

from nas_parser.domain import ProductRecord
from nas_parser.report import RunReport


class ProductValidator:
    """Validate ProductRecord objects without mutating them."""

    def validate(
        self, records: Iterable[ProductRecord], report: RunReport
    ) -> list[ProductRecord]:
        """Validate records and return them unchanged as a list."""
        validated_records: list[ProductRecord] = []

        for record in records:
            self._validate_record(record, report)
            validated_records.append(record)

        return validated_records

    def _validate_record(self, record: ProductRecord, report: RunReport) -> None:
        """Validate a single product record and report warnings or errors."""
        if not self._has_text(record.size):
            report.warning(
                f"Missing size for {record.source_file.name}:{record.source_sheet}:{record.source_row}"
            )
        if record.price is None:
            report.warning(
                f"Missing price for {record.source_file.name}:{record.source_sheet}:{record.source_row}"
            )
        elif self._is_negative_number(record.price):
            report.error(
                f"Negative price for {record.source_file.name}:{record.source_sheet}:{record.source_row}"
            )
        if record.quantity is None:
            report.warning(
                f"Missing quantity for {record.source_file.name}:{record.source_sheet}:{record.source_row}"
            )
        elif self._is_negative_number(record.quantity):
            report.error(
                f"Negative quantity for {record.source_file.name}:{record.source_sheet}:{record.source_row}"
            )
        if self._requires_color_code(record) and not self._has_text(record.color_code):
            report.warning(
                f"Missing color_code for {record.source_file.name}:{record.source_sheet}:{record.source_row}"
            )
        if self._requires_sku(record) and not self._has_text(record.sku):
            report.warning(
                f"Unable to build SKU for {record.source_file.name}:{record.source_sheet}:{record.source_row}"
            )

    @staticmethod
    def _has_text(value: object | None) -> bool:
        """Return whether a value contains non-empty text."""
        return isinstance(value, str) and bool(value.strip())

    @staticmethod
    def _is_negative_number(value: object) -> bool:
        """Return whether a numeric value is below zero."""
        return isinstance(value, Decimal) and value < 0

    @staticmethod
    def _requires_color_code(record: ProductRecord) -> bool:
        """Return whether the product type requires a color code."""
        return record.cut in {"12cut", "16cut"}

    @staticmethod
    def _requires_sku(record: ProductRecord) -> bool:
        """Return whether the record has enough core data to expect a SKU."""
        if record.fixation == "sew":
            return (
                ProductValidator._has_text(record.color)
                and ProductValidator._has_text(record.shape)
                and ProductValidator._has_text(record.size)
            )

        return (
            ProductValidator._has_text(record.cut)
            and ProductValidator._has_text(record.color)
            and ProductValidator._has_text(record.size)
            and ProductValidator._has_text(record.fixation)
            and ProductValidator._has_text(record.color_code)
        )
