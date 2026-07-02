"""Smoke tests for the NAS Parser project scaffold."""

from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from nas_parser.config import AppConfig
from nas_parser.pipeline import Pipeline
from nas_parser.report import RunReport


class TestSmoke(unittest.TestCase):
    """Smoke coverage for the NAS Parser scaffold."""

    def test_pipeline_runs(self) -> None:
        """Verify that the pipeline can process a minimal supported workbook."""
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = Path(temp_dir) / "input"
            reference_dir = Path(temp_dir) / "reference"
            output_dir = Path(temp_dir) / "output"
            output_file = output_dir / "catalog.xlsx"
            logs_dir = Path(temp_dir) / "logs"
            input_dir.mkdir()
            reference_dir.mkdir()
            output_dir.mkdir()

            workbook_path = input_dir / "12cut_hot.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "12cut_hot"
            sheet.append(("hot fixation", None, None, None, None, None, None, None))
            sheet.append(("color", "size", "pack_size", "price", "received", "sold", "stock", "color_code"))
            sheet.append(("Crystal", "SS3", 1440, 135, 516, 142, 374, "001"))
            workbook.save(workbook_path)
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
            info_log_exists = (logs_dir / "info.log").is_file()
            warnings_log_exists = (logs_dir / "warnings.log").is_file()
            errors_log_exists = (logs_dir / "errors.log").is_file()

        self.assertEqual(len(records), 1)
        self.assertIsInstance(report, RunReport)
        self.assertIn("files_found=1", report.summary())
        self.assertIn("files_processed=1", report.summary())
        self.assertIn("files_skipped=0", report.summary())
        self.assertIn("records_created=1", report.summary())
        self.assertIn("warnings=0", report.summary())
        self.assertIn("errors=0", report.summary())
        self.assertTrue(info_log_exists)
        self.assertTrue(warnings_log_exists)
        self.assertTrue(errors_log_exists)
