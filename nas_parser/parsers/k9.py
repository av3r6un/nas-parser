"""K9 parser for NAS Crystal sew-on stones."""

from __future__ import annotations

from collections.abc import Iterable
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import ast
from pathlib import Path
import re

from nas_parser.domain import ProductRecord
from nas_parser.parsers.base import ParserBase
from nas_parser.source import SourceRow


COLOR_COLUMN = 0
SHAPE_COLUMN = 1
SIZE_COLUMN = 2
PRICE_COLUMN = 3
QUANTITY_COLUMN = 7


class K9Parser(ParserBase):
    """Parse K9 source rows into ProductRecord objects."""

    @property
    def name(self) -> str:
        """Return the canonical parser name."""
        return "k9"

    def supports(self, source_file: Path) -> bool:
        """Return whether the parser supports a K9 source file."""
        return source_file.stem.casefold().startswith("k9")

    def parse(self, source_rows: Iterable[SourceRow]) -> Iterable[ProductRecord]:
        """Yield ProductRecord objects from the provided source rows."""
        current_context: tuple[Path, str] | None = None
        current_color: str | None = None
        current_shape: str | None = None

        for row in source_rows:
            context = (row.source_file, row.source_sheet)
            if context != current_context:
                current_context = context
                current_color = None
                current_shape = None

            size_value = self._value_at(row.values, SIZE_COLUMN)
            if not self._is_size_value(size_value):
                continue

            row_color = self._value_at(row.values, COLOR_COLUMN)
            row_shape = self._value_at(row.values, SHAPE_COLUMN)
            if row_color is not None:
                current_color = row_color
            if row_shape is not None:
                current_shape = row_shape

            yield ProductRecord(
                price=self._decimal_value(self._value_at(row.values, PRICE_COLUMN), row.values),
                quantity=self._decimal_value(
                    self._value_at(row.values, QUANTITY_COLUMN),
                    row.values,
                ),
                color=current_color,
                size=self._normalize_size(size_value),
                shape=current_shape,
                fixation="sew",
                cut=None,
                source_file=row.source_file,
                source_sheet=row.source_sheet,
                source_row=row.source_row,
                parser_name=self.name,
            )

    @staticmethod
    def _value_at(values: tuple[object, ...], index: int) -> object | None:
        """Return a value from the row tuple when the index exists."""
        if index >= len(values):
            return None

        return values[index]

    @staticmethod
    def _is_size_value(value: object | None) -> bool:
        """Return whether the value looks like a K9 size entry."""
        if not isinstance(value, str):
            return False

        text = value.strip()
        return (
            "*" in text
            or re.fullmatch(r"\d+(?:[.,]\d+)?\s*MM", text, re.IGNORECASE)
            is not None
        )

    @staticmethod
    def _normalize_size(value: object | None) -> str | None:
        """Normalize sew size formatting for ProductRecord storage."""
        if not isinstance(value, str):
            return None

        return value.strip().replace("*", "x")

    @staticmethod
    def _decimal_value(
        value: object | None,
        row_values: tuple[object, ...],
        seen_columns: set[int] | None = None,
    ) -> Decimal | None:
        """Convert a source value or same-row formula into Decimal."""
        if value is None:
            return None

        if isinstance(value, str) and value.startswith("="):
            return K9Parser._formula_value(value, row_values, seen_columns or set())

        try:
            return Decimal(str(value))
        except (InvalidOperation, ValueError, TypeError):
            return None

    @staticmethod
    def _formula_value(
        formula: str,
        row_values: tuple[object, ...],
        seen_columns: set[int],
    ) -> Decimal | None:
        """Evaluate simple same-row arithmetic formulas used in K9 sheets."""
        expression = formula[1:]
        if not expression:
            return None

        def replace_cell(match: re.Match[str]) -> str:
            column_index = K9Parser._column_letters_to_index(match.group(1))
            if column_index in seen_columns:
                raise ValueError

            next_seen_columns = set(seen_columns)
            next_seen_columns.add(column_index)
            cell_value = K9Parser._decimal_value(
                K9Parser._value_at(row_values, column_index),
                row_values,
                next_seen_columns,
            )
            return str(cell_value or Decimal("0"))

        try:
            expression = re.sub(r"\b([A-Za-z]+)[0-9]+\b", replace_cell, expression)
            return K9Parser._safe_decimal_expression(expression)
        except (InvalidOperation, ValueError, SyntaxError):
            return None

    @staticmethod
    def _safe_decimal_expression(expression: str) -> Decimal | None:
        """Evaluate a restricted numeric expression as Decimal."""
        node = ast.parse(expression, mode="eval")
        return K9Parser._eval_ast_node(node.body)

    @staticmethod
    def _eval_ast_node(node: ast.AST) -> Decimal:
        """Evaluate supported arithmetic AST nodes."""
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return Decimal(str(node.value))

        if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
            return -K9Parser._eval_ast_node(node.operand)

        if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.UAdd):
            return K9Parser._eval_ast_node(node.operand)

        if isinstance(node, ast.BinOp):
            left = K9Parser._eval_ast_node(node.left)
            right = K9Parser._eval_ast_node(node.right)
            if isinstance(node.op, ast.Add):
                return left + right
            if isinstance(node.op, ast.Sub):
                return left - right
            if isinstance(node.op, ast.Mult):
                return left * right
            if isinstance(node.op, ast.Div):
                return left / right

        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id.upper() == "ROUND"
            and len(node.args) in {1, 2}
        ):
            value = K9Parser._eval_ast_node(node.args[0])
            digits = (
                0
                if len(node.args) == 1
                else int(K9Parser._eval_ast_node(node.args[1]))
            )
            quantum = Decimal("1").scaleb(-digits)
            return value.quantize(quantum, rounding=ROUND_HALF_UP)

        raise ValueError

    @staticmethod
    def _column_letters_to_index(letters: str) -> int:
        """Convert Excel column letters to a zero-based column index."""
        index = 0
        for letter in letters.upper():
            index = index * 26 + ord(letter) - ord("A") + 1

        return index - 1
