"""Application entry point for NAS Parser."""

from __future__ import annotations

from nas_parser.config import AppConfig
from nas_parser.pipeline import Pipeline


def main() -> None:
    """Run NAS Parser from the command line."""
    pipeline = Pipeline(AppConfig())
    _records, report = pipeline.run()
    print(report.summary())


if __name__ == "__main__":
    main()
