"""Tests for NAS Parser run reports."""

from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from nas_parser.report import RunReport


class TestRunReport(unittest.TestCase):
    """Coverage for RunReport formatting."""

    def test_logs_include_message_sections_line_by_line(self) -> None:
        """Verify that detailed logs expose collected report messages."""
        report = RunReport()
        report.info("Pipeline started.")
        report.warning("Missing price for K9.xlsx:пришивные:637")
        report.error("Failed to process broken.xlsx: bad file")

        lines = report.logs().splitlines()

        self.assertIn("info=1 warnings=1 errors=1", lines[0])
        self.assertIn("INFO (1):", lines)
        self.assertIn("  - Pipeline started.", lines)
        self.assertIn("WARNINGS (1):", lines)
        self.assertIn("  - Missing price for K9.xlsx:пришивные:637", lines)
        self.assertIn("ERRORS (1):", lines)
        self.assertIn("  - Failed to process broken.xlsx: bad file", lines)

    def test_logs_include_empty_sections(self) -> None:
        """Verify that empty sections are still visible in logs."""
        report = RunReport()

        lines = report.logs().splitlines()

        self.assertIn("INFO (0):", lines)
        self.assertIn("WARNINGS (0):", lines)
        self.assertIn("ERRORS (0):", lines)
        self.assertEqual(lines.count("  - none"), 3)

    def test_logs_include_reference_updates_without_changes(self) -> None:
        """Verify reference update logs include the no-changes state."""
        report = RunReport()

        report.set_reference_update(generated_colors=0)
        lines = report.logs().splitlines()

        self.assertIn("Reference updates:", lines)
        self.assertIn("  No changes", lines)

    def test_logs_include_generated_reference_update(self) -> None:
        """Verify reference update logs include generated file details."""
        report = RunReport()
        generated_reference = Path("reference/generated/colorcode-articul_gen1.xlsx")

        report.set_reference_update(
            generated_colors=2,
            generated_reference=generated_reference,
        )
        lines = report.logs().splitlines()

        self.assertIn("Reference updates:", lines)
        self.assertIn("  Generated colors: 2", lines)
        self.assertIn(f"  Generated reference: {generated_reference}", lines)

    def test_write_logs_creates_named_log_files(self) -> None:
        """Verify that report messages are written into separate files."""
        report = RunReport()
        report.info("Pipeline started.")
        report.warning("Missing quantity for K9.xlsx:пришивные:783")
        report.error("Failed to process broken.xlsx: bad file")

        with tempfile.TemporaryDirectory() as temp_dir:
            logs_dir = Path(temp_dir) / "logs"

            report.write_logs(logs_dir)

            self.assertEqual(
                (logs_dir / "info.log").read_text(encoding="utf-8"),
                "Pipeline started.\n",
            )
            self.assertEqual(
                (logs_dir / "warnings.log").read_text(encoding="utf-8"),
                "Missing quantity for K9.xlsx:пришивные:783\n",
            )
            self.assertEqual(
                (logs_dir / "errors.log").read_text(encoding="utf-8"),
                "Failed to process broken.xlsx: bad file\n",
            )
