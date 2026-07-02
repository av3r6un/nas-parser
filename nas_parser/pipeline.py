"""Pipeline orchestration for NAS Parser."""

from __future__ import annotations

from pathlib import Path

from nas_parser.config import AppConfig
from nas_parser.business import ProductEnricher
from nas_parser.domain import ProductRecord
from nas_parser.export import ExcelExporter
from nas_parser.parsers import CutParser, K9Parser, ParserRegistry
from nas_parser.readers import ExcelReader
from nas_parser.references.colors import ColorReference, ColorReferenceLoader
from nas_parser.report import RunReport
from nas_parser.validation import ProductValidator


class Pipeline:
    """Coordinate the high-level NAS Parser execution flow."""

    def __init__(self, config: AppConfig | None = None) -> None:
        """Initialize the pipeline with application configuration."""
        self._config = config if config is not None else AppConfig()
        self._parser_registry = ParserRegistry([CutParser(), K9Parser()])

    def run(self) -> tuple[list[ProductRecord], RunReport]:
        """Run the current pipeline and return parsed records with the report."""
        report = RunReport()
        report.info("Pipeline started.")

        color_reference = self._load_color_reference()
        enricher = ProductEnricher(color_reference)
        validator = ProductValidator()
        records: list[ProductRecord] = []
        input_files = self._collect_input_files()
        processed_files = 0
        skipped_files = 0

        for source_file in input_files:
            parser = self._parser_registry.find(source_file)
            if parser is None:
                report.warning(f"No parser found for {source_file.name}.")
                skipped_files += 1
                continue

            try:
                reader = ExcelReader(source_file)
                source_rows = reader.read()
                parsed_records = list(parser.parse(source_rows))
                parsed_records = enricher.enrich(parsed_records, report)
                parsed_records = validator.validate(parsed_records, report)
            except Exception as exc:  # pragma: no cover - defensive pipeline guard
                report.error(f"Failed to process {source_file.name}: {exc}")
                skipped_files += 1
                continue

            records.extend(parsed_records)
            processed_files += 1
            report.info(
                f"Processed {source_file.name}: {len(parsed_records)} records."
            )

        report.set_statistics(
            files_found=len(input_files),
            files_processed=processed_files,
            files_skipped=skipped_files,
            records_created=len(records),
        )
        report.set_record_statistics(records)
        self._config.output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file = ExcelExporter(self._config.output_file).export(records)
        report.set_output_file(output_file)
        report.info("Pipeline finished.")
        report.write_logs(self._config.logs_dir)
        return records, report

    def _load_color_reference(self) -> ColorReference:
        """Load the color reference workbook when it exists."""
        reference_file = self._config.reference_dir / "colorcode-articul.xlsx"
        if reference_file.is_file():
            return ColorReferenceLoader(reference_file).load()

        return ColorReference(source_file=reference_file)

    def _collect_input_files(self) -> list[Path]:
        """Return all Excel files from the configured input directory."""
        if not self._config.input_dir.is_dir():
            return []

        return sorted(
            path
            for path in self._config.input_dir.iterdir()
            if path.is_file() and path.suffix.lower() == ".xlsx"
        )
