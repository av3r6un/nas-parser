"""Parser registry for NAS Parser."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

from nas_parser.parsers.base import ParserBase


class ParserRegistry:
    """Store parser instances and resolve a matching parser by source file."""

    def __init__(self, parsers: Iterable[ParserBase] | None = None) -> None:
        """Initialize the registry with an optional iterable of parsers."""
        self._parsers: list[ParserBase] = list(parsers) if parsers is not None else []

    def register(self, parser: ParserBase) -> None:
        """Register a parser instance in the registry."""
        if any(existing is parser for existing in self._parsers):
            raise ValueError("The same parser instance is already registered.")

        self._parsers.append(parser)

    def all(self) -> tuple[ParserBase, ...]:
        """Return all registered parsers as an immutable tuple."""
        return tuple(self._parsers)

    def find(self, source_file: Path) -> ParserBase | None:
        """Return the first registered parser that supports the given file."""
        for parser in self._parsers:
            if parser.supports(source_file):
                return parser

        return None
