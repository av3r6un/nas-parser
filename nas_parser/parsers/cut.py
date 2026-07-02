"""Cut parser for NAS Crystal 12cut and 16cut files."""

from __future__ import annotations

from collections.abc import Iterable
import re
from decimal import Decimal, InvalidOperation
from pathlib import Path

from nas_parser.domain import ProductRecord
from nas_parser.parsers.base import ParserBase
from nas_parser.source import SourceRow


class CutParser(ParserBase):
    """Parse 12cut and 16cut source rows into ProductRecord objects."""

    @property
    def name(self) -> str:
        """Return the canonical parser name."""
        return "cut"

    def supports(self, source_file: Path) -> bool:
        """Return whether the parser supports a 12cut or 16cut file."""
        file_name = source_file.name.lower()
        return "12cut" in file_name or "16cut" in file_name

    def parse(self, source_rows: Iterable[SourceRow]) -> Iterable[ProductRecord]:
        """Yield ProductRecord objects from the provided source rows."""
        current_context: tuple[Path, str] | None = None
        current_color: str | None = None
        current_color_code: str | None = None
        current_cut: str | None = None
        current_fixation: str | None = None

        for row in source_rows:
            context = (row.source_file, row.source_sheet)
            if context != current_context:
                current_context = context
                current_color = None
                current_color_code = None
                current_cut = self._determine_cut(row.source_file)
                current_fixation = self._determine_fixation(row.source_file)

            if not self._is_product_row(row.values):
                continue

            size_value = self._string_value(self._get_value(row.values, 1))
            price_value = self._decimal_value(self._get_value(row.values, 3), row.values)
            quantity_value = self._decimal_value(self._get_value(row.values, 6), row.values)
            row_color_value = self._string_value(self._get_value(row.values, 0))
            row_color_code = self._string_value(self._get_value(row.values, 7))

            if row_color_value is not None:
                current_color = row_color_value
                if row_color_code is not None:
                    current_color_code = row_color_code

            if current_color is None or size_value is None:
                continue
            if price_value is None or quantity_value is None:
                continue

            yield ProductRecord(
                price=price_value,
                quantity=quantity_value,
                color=current_color,
                color_code=current_color_code,
                size=size_value,
                shape=None,
                fixation=current_fixation,
                cut=current_cut,
                source_file=row.source_file,
                source_sheet=row.source_sheet,
                source_row=row.source_row,
                parser_name=self.name,
            )

    @staticmethod
    def _determine_cut(source_file: Path) -> str | None:
        """Return the cut label derived from the source file name."""
        file_name = source_file.name.lower()
        if "12cut" in file_name:
            return "12cut"
        if "16cut" in file_name:
            return "16cut"
        return None

    @staticmethod
    def _determine_fixation(source_file: Path) -> str:
        """Return the fixation label derived from the source file name."""
        return "hot" if "hot" in source_file.name.lower() else "non"

    @staticmethod
    def _is_product_row(values: tuple[object, ...]) -> bool:
        """Return whether the row looks like a concrete product row."""
        size_value = CutParser._string_value(CutParser._get_value(values, 1))
        if size_value is None:
            return False

        return (
            CutParser._decimal_value(CutParser._get_value(values, 3), values) is not None
            and CutParser._decimal_value(CutParser._get_value(values, 6), values) is not None
        )

    @staticmethod
    def _get_value(values: tuple[object, ...], index: int) -> object | None:
        """Return the value at the requested index, or None when it is absent."""
        if index >= len(values):
            return None

        return values[index]

    @staticmethod
    def _string_value(value: object | None) -> str | None:
        """Convert a cell value to a stripped string when possible."""
        if value is None:
            return None

        text = str(value).strip()
        return text or None

    @staticmethod
    def _decimal_value(
        value: object | None,
        row_values: tuple[object, ...] | None = None,
        seen_columns: set[int] | None = None,
    ) -> Decimal | None:
        """Convert a cell value to Decimal when possible."""
        if value is None:
            return None

        try:
            return Decimal(str(value))
        except (InvalidOperation, ValueError, TypeError):
            if isinstance(value, str) and value.startswith("=") and row_values is not None:
                return CutParser._formula_value(value, row_values, seen_columns or set())

        return None

    @staticmethod
    def _formula_value(
        formula: str,
        row_values: tuple[object, ...],
        seen_columns: set[int],
    ) -> Decimal | None:
        """Evaluate simple same-row addition and subtraction formulas."""
        expression = formula[1:].replace(" ", "")
        if not expression:
            return None

        total = Decimal("0")
        for token in re.findall(r"[+-]?[^+-]+", expression):
            sign = Decimal("-1") if token.startswith("-") else Decimal("1")
            token_value = token[1:] if token[:1] in {"+", "-"} else token
            decimal_value = CutParser._formula_token_value(
                token_value,
                row_values,
                seen_columns,
            )
            if decimal_value is None:
                return None

            total += sign * decimal_value

        return total

    @staticmethod
    def _formula_token_value(
        token: str,
        row_values: tuple[object, ...],
        seen_columns: set[int],
    ) -> Decimal | None:
        """Return a Decimal value for a formula token."""
        cell_match = re.fullmatch(r"([A-Za-z]+)[0-9]+", token)
        if cell_match is None:
            try:
                return Decimal(token)
            except (InvalidOperation, ValueError):
                return None

        column_index = CutParser._column_letters_to_index(cell_match.group(1))
        if column_index in seen_columns:
            return None

        next_seen_columns = set(seen_columns)
        next_seen_columns.add(column_index)
        return CutParser._decimal_value(
            CutParser._get_value(row_values, column_index),
            row_values,
            next_seen_columns,
        )

    @staticmethod
    def _column_letters_to_index(letters: str) -> int:
        """Convert Excel column letters to a zero-based column index."""
        index = 0
        for letter in letters.upper():
            index = index * 26 + ord(letter) - ord("A") + 1

        return index - 1
