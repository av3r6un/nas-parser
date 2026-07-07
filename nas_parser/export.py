"""Excel export utilities for NAS Parser."""

from __future__ import annotations

from collections.abc import Iterable
from decimal import Decimal, ROUND_CEILING
from pathlib import Path

from openpyxl import Workbook

from nas_parser.domain import ProductRecord


class ExcelExporter:
    """Export ProductRecord objects to the 1C Excel import format."""

    HEADER = (
        "Артикул",
        "Наименование",
        "Цена",
        "Количество",
        "Категория",
        "Цвет",
        "Размер",
        "Форма",
        "Фиксация",
        "Грани",
    )

    def __init__(self, output_file: Path) -> None:
        """Initialize the exporter with the destination Excel file path."""
        self._output_file = output_file

    def export(self, records: Iterable[ProductRecord]) -> Path:
        """Write product records to an Excel file and return its path."""
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Товары"
        sheet.append(self.HEADER)

        sorted_records = sorted(
            records,
            key=lambda record: (
                record.category or "",
                record.color or "",
                record.shape or "",
                record.size or "",
                record.sku or "",
            ),
        )

        for record in sorted_records:
            sheet.append(self._record_to_row(record))

        workbook.save(self._output_file)
        workbook.close()
        return self._output_file

    def _record_to_row(self, record: ProductRecord) -> tuple[object | None, ...]:
        """Map a ProductRecord to a row in the export template."""
        return (
            record.sku or None,
            record.name or None,
            self._export_price(record.price),
            record.quantity if record.quantity is not None else None,
            record.category or None,
            record.color_code or record.color or None,
            record.size or None,
            record.shape or None,
            self._export_fixation(record.fixation),
            self._export_cut(record.cut),
        )

    @staticmethod
    def _export_fixation(fixation: str | None) -> str | None:
        """Return the export value for the fixation column."""
        if fixation == "hot":
            return "Hot"
        if fixation == "non":
            return "Non"
        if fixation == "sew":
            return "K9"

        return fixation or None

    @staticmethod
    def _export_cut(cut: str | None) -> str | None:
        """Return the export value for the cut column."""
        return cut or None

    @staticmethod
    def _export_price(price: Decimal | None) -> int | None:
        """Return the export value for price as whole rubles rounded upward."""
        if price is None:
            return None

        return int(price.to_integral_value(rounding=ROUND_CEILING))
