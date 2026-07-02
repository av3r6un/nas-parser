"""Base reader contract for NAS Parser."""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Iterable

from nas_parser.source import SourceRow


class ReaderBase(ABC):
    """Define the common contract for all future NAS Parser readers."""

    @property
    @abstractmethod
    def name(self) -> str:
        """Return the canonical reader name."""

    @abstractmethod
    def read(self) -> Iterable[SourceRow]:
        """Return source rows as an iterable."""

