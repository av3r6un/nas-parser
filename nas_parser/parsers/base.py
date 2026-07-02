"""Base parser contract for NAS Parser."""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Iterable
from pathlib import Path

from nas_parser.domain import ProductRecord
from nas_parser.source import SourceRow


class ParserBase(ABC):
    """Define the common contract for all future NAS Parser parsers."""

    @property
    @abstractmethod
    def name(self) -> str:
        """Return the canonical parser name."""

    @abstractmethod
    def supports(self, source_file: Path) -> bool:
        """Return whether this parser can handle the given source file."""

    @abstractmethod
    def parse(self, source_rows: Iterable[SourceRow]) -> Iterable[ProductRecord]:
        """Return parsed product records from the given source rows."""
