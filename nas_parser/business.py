"""Business enrichment for NAS Parser products."""

from __future__ import annotations

from collections.abc import Iterable

from nas_parser.domain import ProductRecord
from nas_parser.references.colors import ColorReference
from nas_parser.report import RunReport


class ProductEnricher:
    """Enrich ProductRecord objects with business fields."""

    def __init__(self, color_reference: ColorReference) -> None:
        """Initialize the enricher with a color reference."""
        self._color_reference = color_reference

    def enrich(
        self, records: Iterable[ProductRecord], report: RunReport
    ) -> list[ProductRecord]:
        """Enrich the given records and return them as a list."""
        enriched_records: list[ProductRecord] = []

        for record in records:
            self._apply_color_code(record)
            self._build_category(record)
            self._build_name(record)
            self._build_sku(record)
            enriched_records.append(record)

        return enriched_records

    def _apply_color_code(self, record: ProductRecord) -> None:
        """Fill color_code from the color reference when possible."""
        reference_code = self._color_reference.get_code(record.color)
        if reference_code is not None:
            record.color_code = reference_code

    def _build_sku(self, record: ProductRecord) -> None:
        """Build SKU from the enriched product fields when possible."""
        if record.fixation == "sew":
            if not record.color or not record.shape or not record.size:
                record.sku = None
                return

            record.sku = "/".join(
                (
                    self._format_color_for_sku(record.color),
                    self._format_shape_for_sku(record.shape),
                    self._format_size_for_sku(record.size),
                )
            )
            return

        if not record.cut or not record.color or not record.size:
            record.sku = None
            return

        if record.color_code is None or not record.fixation:
            record.sku = None
            return

        record.sku = "/".join(
            (
                record.cut,
                self._format_color_for_sku(record.color),
                record.size,
                self._format_fixation_for_sku(record.fixation),
                record.color_code,
            )
        )

    def _build_name(self, record: ProductRecord) -> None:
        """Build a display name from the enriched product fields when possible."""
        if record.fixation == "sew":
            if not record.color or not record.shape:
                record.name = None
                return

            record.name = (
                f"{self._format_color_for_name(record.color)} "
                f"{self._format_shape_for_name(record.shape)}"
            )
            return

        if not record.color:
            record.name = None
            return

        record.name = self._format_color_for_name(record.color)

    def _build_category(self, record: ProductRecord) -> None:
        """Build the product category from fixation."""
        if record.fixation == "hot":
            record.category = "Стразы горячей фиксации"
        elif record.fixation == "non":
            record.category = "Стразы холодной фиксации"
        elif record.fixation == "sew":
            record.category = "Пришивные стразы"
        else:
            record.category = None

    @staticmethod
    def _format_cut_for_name(cut: str) -> str:
        """Convert a cut label into the display format used in names."""
        return cut.replace("cut", "", 1)

    @staticmethod
    def _format_fixation_for_name(fixation: str) -> str:
        """Convert fixation into the display format used in names."""
        return "Hotfix" if fixation == "hot" else "Non Hotfix"

    @staticmethod
    def _format_fixation_for_sku(fixation: str) -> str:
        """Convert fixation into the display format used in SKUs."""
        return "Hot" if fixation == "hot" else "Non"

    @staticmethod
    def _format_color_for_name(color: str) -> str:
        """Convert a color string into the display format used in names."""
        return " ".join(ProductEnricher._format_name_word(word) for word in color.strip().split())

    @staticmethod
    def _format_color_for_sku(color: str) -> str:
        """Convert a color string into the display format used in SKUs."""
        return " ".join(ProductEnricher._format_sku_word(word) for word in color.strip().split())

    @staticmethod
    def _format_shape_for_sku(shape: str) -> str:
        """Convert a shape string into the display format used in SKUs."""
        return " ".join(ProductEnricher._format_sku_word(word) for word in shape.strip().split())

    @staticmethod
    def _format_shape_for_name(shape: str) -> str:
        """Convert a shape string into the display format used in names."""
        return " ".join(ProductEnricher._format_name_word(word) for word in shape.strip().split())

    @staticmethod
    def _format_size_for_sku(size: str) -> str:
        """Convert a size string into the display format used in SKUs."""
        return size.strip().replace("*", "x")

    @staticmethod
    def _format_name_word(word: str) -> str:
        """Format a single color word for a display name."""
        if word.casefold() == "ab":
            return "AB"
        if word.isupper() and len(word) > 2:
            return word.capitalize()
        return word[:1].upper() + word[1:].lower() if word else word

    @staticmethod
    def _format_sku_word(word: str) -> str:
        """Format a single color word for a SKU."""
        if word.casefold() == "ab":
            return "Ab"
        return word[:1].upper() + word[1:].lower() if word else word

    @staticmethod
    def _has_text(value: object | None) -> bool:
        """Return whether a value contains non-empty text."""
        return isinstance(value, str) and bool(value.strip())
