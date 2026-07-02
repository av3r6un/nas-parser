"""Tests for NAS Parser source row domain model."""

from __future__ import annotations

from pathlib import Path
import unittest

from nas_parser.source import SourceRow


class TestSourceRow(unittest.TestCase):
    """Coverage for the SourceRow dataclass."""

    def test_source_row_can_be_created(self) -> None:
        """Verify that SourceRow stores origin data and row values."""
        row = SourceRow(
            values=("Topaz", 12, 3.5),
            source_file=Path("input/sample.xlsx"),
            source_sheet="Sheet1",
            source_row=7,
        )

        self.assertEqual(row.values, ("Topaz", 12, 3.5))
        self.assertEqual(row.source_file, Path("input/sample.xlsx"))
        self.assertEqual(row.source_sheet, "Sheet1")
        self.assertEqual(row.source_row, 7)

    def test_source_row_preserves_tuple_values(self) -> None:
        """Verify that SourceRow keeps the tuple payload unchanged."""
        values = ("A", "B", "C")
        row = SourceRow(
            values=values,
            source_file=Path("input/another.xlsx"),
            source_sheet="Data",
            source_row=1,
        )

        self.assertIs(row.values, values)
