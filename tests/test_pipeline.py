"""Tests for the integrated NAS Parser pipeline."""

from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

from nas_parser.config import AppConfig
from nas_parser.pipeline import Pipeline


def _write_color_reference_workbook(workbook_path: Path) -> None:
    """Create a small color reference workbook for pipeline tests."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(
        (
            "\u0410\u0440\u0442\u0438\u043a\u0443\u043b WB \u0438 Ozon",
            "\u0413\u0440\u0430\u043d\u0438",
            "\u0426\u0432\u0435\u0442",
            "\u0420\u0430\u0437\u043c\u0435\u0440",
            "\u0422\u0438\u043f",
            "\u0410\u0440\u0442\u0438\u043a\u0443\u043b \u0446\u0432\u0435\u0442\u0430",
        )
    )
    sheet.append(("12cut/Crystal/SS3/Hot/001", "12cut", "Crystal", "SS3", "Hot", "001"))
    sheet.append(
        ("16cut/Crystal AB/SS30/Non/001+", "16cut", "Crystal AB", "SS30", "Non", "001+")
    )
    sheet.append(("16cut/New AB/SS16/Hot/112", "16cut", "New AB", "SS16", "Hot", "112"))
    workbook.save(workbook_path)
    workbook.close()


def _write_cut_workbook(
    workbook_path: Path,
    *,
    sheet_title: str,
    title_row: tuple[object, ...],
    color_row: tuple[object, ...],
    second_row: tuple[object, ...],
) -> None:
    """Create a small cut workbook for pipeline tests."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_title
    sheet.append(title_row)
    sheet.append(("color", "size", "pack_size", "price", "received", "sold", "stock", "color_code"))
    sheet.append(color_row)
    sheet.append(second_row)
    workbook.save(workbook_path)
    workbook.close()


class TestPipeline(unittest.TestCase):
    """Coverage for the first working pipeline."""

    def test_pipeline_processes_supported_files_and_skips_unsupported_ones(self) -> None:
        """Verify the pipeline runs end to end across supported and unsupported files."""
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = Path(temp_dir) / "input"
            reference_dir = Path(temp_dir) / "reference"
            output_dir = Path(temp_dir) / "output"
            output_file = output_dir / "catalog.xlsx"
            logs_dir = Path(temp_dir) / "logs"
            input_dir.mkdir()
            reference_dir.mkdir()
            output_dir.mkdir()

            _write_color_reference_workbook(reference_dir / "colorcode-articul.xlsx")

            _write_cut_workbook(
                input_dir / "12cut_hot.xlsx",
                sheet_title="12cut_hot",
                title_row=("hot fixation", None, None, None, None, None, None, None),
                color_row=("Crystal", "SS3", 1440, 135, 516, 142, 374, "001"),
                second_row=(None, "SS4", 1440, 135, 600, 63, 537, None),
            )
            _write_cut_workbook(
                input_dir / "16cut_non.xlsx",
                sheet_title="16cut_non",
                title_row=("non fixation", None, None, None, None, None, None, None),
                color_row=("CRYSTAL", "SS6", 1440, 590, 57, 1, 56, "001"),
                second_row=(None, "SS8", 1440, 701, 65, 2, 63, None),
            )

            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "K9"
            sheet.append(("K9 title", None, None, None, None, None, None, None))
            sheet.append(("color", "shape", "size", "price", "pack", "received", "sold", "stock"))
            sheet.append(("CRYSTAL", "Drop", "7*12", 11, 72, 1886, None, 1886))
            workbook.save(input_dir / "K9.xlsx")
            workbook.close()

            pipeline = Pipeline(
                AppConfig(
                    input_dir=input_dir,
                    reference_dir=reference_dir,
                    output_dir=output_dir,
                    output_file=output_file,
                    logs_dir=logs_dir,
                )
            )
            records, report = pipeline.run()

            self.assertTrue(output_file.is_file())
            workbook = load_workbook(output_file, read_only=True, data_only=True)
            try:
                self.assertIn("\u0422\u043e\u0432\u0430\u0440\u044b", workbook.sheetnames)
                sheet = workbook["\u0422\u043e\u0432\u0430\u0440\u044b"]
                output_rows = list(sheet.iter_rows(values_only=True))
            finally:
                workbook.close()

        self.assertEqual(len(records), 5)
        self.assertGreater(len(records), 0)
        self.assertGreater(len(output_rows), 1)
        self.assertEqual(records[0].color_code, "001")
        self.assertEqual(records[0].sku, "12cut/Crystal/SS3/Hot/001")
        self.assertEqual(records[0].category, "Стразы горячей фиксации")
        self.assertIn("Crystal SS3 Hotfix", records[0].name)
        self.assertEqual(records[2].color_code, "001")
        self.assertEqual(records[2].sku, "16cut/Crystal/SS6/Non/001")
        self.assertEqual(records[2].category, "Стразы холодной фиксации")
        self.assertIn("Crystal SS6 Non Hotfix", records[2].name)
        self.assertEqual(records[4].parser_name, "k9")
        self.assertEqual(records[4].color, "CRYSTAL")
        self.assertEqual(records[4].shape, "Drop")
        self.assertEqual(records[4].size, "7*12")
        self.assertEqual(records[4].fixation, "sew")
        self.assertIsNone(records[4].cut)
        self.assertEqual(records[4].sku, "Crystal/Drop/7x12")
        self.assertEqual(records[4].name, "Пришивные стразы Crystal Drop 7*12")
        self.assertEqual(records[4].category, "Пришивные стразы")
        self.assertEqual(records[4].color_code, "001")
        self.assertIn("files_found=3", report.summary())
        self.assertIn("files_processed=3", report.summary())
        self.assertIn("files_skipped=0", report.summary())
        self.assertIn("records_created=5", report.summary())
        self.assertIn("warnings=0", report.summary())
        self.assertIn("errors=0", report.summary())
        self.assertIn(f"output_file={output_file}", report.summary())
