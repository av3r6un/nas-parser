"""Tests for NAS Parser parser registry."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
import unittest

from nas_parser.domain import ProductRecord
from nas_parser.parsers import ParserBase, ParserRegistry
from nas_parser.source import SourceRow


class DummyParser(ParserBase):
    """Minimal parser implementation used to verify registry behavior."""

    def __init__(self, name: str, supports_result: bool) -> None:
        """Initialize the dummy parser with a fixed support result."""
        self._name = name
        self._supports_result = supports_result
        self.supports_calls: list[Path] = []

    @property
    def name(self) -> str:
        """Return the parser name."""
        return self._name

    def supports(self, source_file: Path) -> bool:
        """Record the call and return the configured support result."""
        self.supports_calls.append(source_file)
        return self._supports_result

    def parse(self, source_rows: Iterable[SourceRow]) -> Iterable[ProductRecord]:
        """Parsing is not part of this stage."""
        raise NotImplementedError


class TestParserRegistry(unittest.TestCase):
    """Coverage for the parser registry."""

    def test_registers_one_parser(self) -> None:
        """Verify that a single parser can be registered."""
        registry = ParserRegistry()
        parser = DummyParser("one", True)

        registry.register(parser)

        self.assertEqual(registry.all(), (parser,))

    def test_registers_multiple_parsers(self) -> None:
        """Verify that multiple parsers are stored in order."""
        registry = ParserRegistry()
        first = DummyParser("first", False)
        second = DummyParser("second", True)

        registry.register(first)
        registry.register(second)

        self.assertEqual(registry.all(), (first, second))

    def test_all_returns_tuple(self) -> None:
        """Verify that all() returns an immutable tuple."""
        parser = DummyParser("one", True)
        registry = ParserRegistry([parser])

        self.assertIsInstance(registry.all(), tuple)

    def test_registering_same_instance_twice_raises_value_error(self) -> None:
        """Verify that the same parser instance cannot be registered twice."""
        registry = ParserRegistry()
        parser = DummyParser("one", True)

        registry.register(parser)

        with self.assertRaises(ValueError):
            registry.register(parser)

    def test_find_returns_first_supported_parser(self) -> None:
        """Verify that the first parser returning True is selected."""
        first = DummyParser("first", False)
        second = DummyParser("second", True)
        registry = ParserRegistry([first, second])
        source_file = Path("input/sample.xlsx")

        found = registry.find(source_file)

        self.assertIs(found, second)
        self.assertEqual(first.supports_calls, [source_file])
        self.assertEqual(second.supports_calls, [source_file])

    def test_find_returns_none_when_no_parser_supports_file(self) -> None:
        """Verify that find() returns None when no parser matches."""
        registry = ParserRegistry([DummyParser("first", False)])

        self.assertIsNone(registry.find(Path("input/sample.xlsx")))

    def test_find_uses_supports_instead_of_filename_rules(self) -> None:
        """Verify that the registry delegates selection to supports()."""
        parser = DummyParser("custom", True)
        registry = ParserRegistry([parser])
        source_file = Path("input/does-not-match-name.xlsx")

        found = registry.find(source_file)

        self.assertIs(found, parser)
        self.assertEqual(parser.supports_calls, [source_file])
