"""Execution reporting utilities for NAS Parser."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

from nas_parser.domain import ProductRecord


class RunReport:
    """Collect informational messages, warnings, errors, and run summary data."""

    def __init__(self) -> None:
        """Initialize an empty run report."""
        self._info_messages: list[str] = []
        self._warning_messages: list[str] = []
        self._error_messages: list[str] = []
        self._files_found = 0
        self._files_processed = 0
        self._files_skipped = 0
        self._records_created = 0
        self._records_12cut = 0
        self._records_16cut = 0
        self._records_k9 = 0
        self._output_file: Path | None = None

    def info(self, message: str) -> None:
        """Record an informational message."""
        self._info_messages.append(message)

    def warning(self, message: str) -> None:
        """Record a warning message."""
        self._warning_messages.append(message)

    def error(self, message: str) -> None:
        """Record an error message."""
        self._error_messages.append(message)

    def set_output_file(self, path: Path) -> None:
        """Store the generated output file path for summary output."""
        self._output_file = path

    def set_record_statistics(self, records: Iterable[ProductRecord]) -> None:
        """Store product type statistics based on existing records."""
        self._records_12cut = 0
        self._records_16cut = 0
        self._records_k9 = 0

        for record in records:
            if record.cut == "12cut":
                self._records_12cut += 1
            elif record.cut == "16cut":
                self._records_16cut += 1
            elif record.parser_name == "k9":
                self._records_k9 += 1

    def set_statistics(
        self,
        *,
        files_found: int,
        files_processed: int,
        files_skipped: int,
        records_created: int,
    ) -> None:
        """Store the final run statistics for summary output."""
        self._files_found = files_found
        self._files_processed = files_processed
        self._files_skipped = files_skipped
        self._records_created = records_created

    def summary(self) -> str:
        """Return a compact summary of collected report messages."""
        summary = (
            f"files_found={self._files_found} "
            f"files_processed={self._files_processed} "
            f"files_skipped={self._files_skipped} "
            f"records_created={self._records_created} "
            f"records_12cut={self._records_12cut} "
            f"records_16cut={self._records_16cut} "
            f"records_k9={self._records_k9} "
            f"info={len(self._info_messages)} "
            f"warnings={len(self._warning_messages)} "
            f"errors={len(self._error_messages)}"
        )
        if self._output_file is not None:
            summary = f"{summary} output_file={self._output_file}"

        return summary

    def logs(self) -> str:
        """Return a line-by-line run log with summary and collected messages."""
        lines = [self.summary()]
        lines.extend(self._message_section("INFO", self._info_messages))
        lines.extend(self._message_section("WARNINGS", self._warning_messages))
        lines.extend(self._message_section("ERRORS", self._error_messages))
        return "\n".join(lines)

    def write_logs(self, logs_dir: Path) -> None:
        """Write report messages into separate log files."""
        logs_dir.mkdir(parents=True, exist_ok=True)
        self._write_message_file(logs_dir / "info.log", self._info_messages)
        self._write_message_file(logs_dir / "warnings.log", self._warning_messages)
        self._write_message_file(logs_dir / "errors.log", self._error_messages)

    @staticmethod
    def _message_section(title: str, messages: list[str]) -> list[str]:
        """Format one report message section."""
        lines = [f"{title} ({len(messages)}):"]
        if not messages:
            lines.append("  - none")
            return lines

        lines.extend(f"  - {message}" for message in messages)
        return lines

    @staticmethod
    def _write_message_file(path: Path, messages: list[str]) -> None:
        """Write one message per line to a log file."""
        content = "\n".join(messages)
        if content:
            content = f"{content}\n"
        path.write_text(content, encoding="utf-8")
