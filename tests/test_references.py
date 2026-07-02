"""Tests for NAS Parser reference infrastructure."""

from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from nas_parser.references import ColorReference, ReferenceBase, ReferenceLoader
from nas_parser.references import ColorReferenceLoader


class TestColorReference(unittest.TestCase):
    """Coverage for the color reference contract."""

    def test_color_reference_exposes_the_expected_contract(self) -> None:
        """Verify that ColorReference exposes the base reference fields."""
        reference = ColorReference(source_file=Path("reference/colorcode-articul.xlsx"))

        self.assertEqual(reference.name, "colors")
        self.assertEqual(reference.source_file, Path("reference/colorcode-articul.xlsx"))
        self.assertFalse(reference.loaded)
        self.assertIsInstance(reference, ReferenceBase)


class TestReferenceLoader(unittest.TestCase):
    """Coverage for the reference loader infrastructure."""

    def test_reference_loader_registers_and_returns_references(self) -> None:
        """Verify that the loader stores and retrieves references by name."""
        reference = ColorReference()
        loader = ReferenceLoader()

        loader.register(reference)

        self.assertIs(loader.get("colors"), reference)
        self.assertEqual(loader.all(), (reference,))

    def test_reference_loader_returns_none_for_unknown_reference(self) -> None:
        """Verify that the loader returns None when a name is not registered."""
        loader = ReferenceLoader()

        self.assertIsNone(loader.get("missing"))


class TestColorReferenceLoader(unittest.TestCase):
    """Coverage for loading color codes from an Excel workbook."""

    def test_color_reference_loader_loads_codes_into_lookup_structure(self) -> None:
        """Verify that color names are loaded into a fast lookup structure."""
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "colorcode-articul.xlsx"

            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Лист1"
            sheet.append(["Артикул WB и Ozon", "Грани", "Цвет", "Размер", "Тип", "Артикул цвета"])
            sheet.append(["12cut/Crystal/SS3/Hot/001", "12cut", "Crystal", "SS3", "Hot", "001"])
            sheet.append(["12cut/Topaz/SS3/Hot/008", "12cut", "Topaz", "SS3", "Hot", "008"])
            sheet.append([None, None, None, None, None, None])
            workbook.save(workbook_path)
            workbook.close()

            loader = ColorReferenceLoader(workbook_path)
            reference = loader.load()

        self.assertTrue(reference.loaded)
        self.assertEqual(reference.source_file, workbook_path)
        self.assertEqual(reference.get_code("Crystal"), "001")
        self.assertEqual(reference.get_code("Topaz"), "008")
        self.assertEqual(reference.get_code("Unknown"), None)

    def test_color_reference_loader_strips_outer_whitespace_only(self) -> None:
        """Verify that lookup handles surrounding whitespace without further normalization."""
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "colorcode-articul.xlsx"

            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Лист1"
            sheet.append(["Артикул WB и Ozon", "Грани", "Цвет", "Размер", "Тип", "Артикул цвета"])
            sheet.append(["12cut/Crystal/SS3/Hot/001", "12cut", "Crystal", "SS3", "Hot", "001"])
            workbook.save(workbook_path)
            workbook.close()

            reference = ColorReferenceLoader(workbook_path).load()

        self.assertEqual(reference.get_code("  Crystal  "), "001")
