"""Base reference contracts and loader utilities for NAS Parser."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
from typing import Protocol, runtime_checkable


@runtime_checkable
class ReferenceBase(Protocol):
    """Define the shared interface for reference objects."""

    @property
    def name(self) -> str:
        """Return the canonical reference name."""

    source_file: Path | None
    loaded: bool


class ReferenceLoader:
    """Store and retrieve reference objects by their canonical name."""

    def __init__(
        self, references: Iterable[ReferenceBase] | None = None
    ) -> None:
        """Initialize the loader with an optional iterable of references."""
        self._references: list[ReferenceBase] = (
            list(references) if references is not None else []
        )

    def register(self, reference: ReferenceBase) -> None:
        """Register a reference, replacing an existing one with the same name."""
        for index, existing in enumerate(self._references):
            if existing.name == reference.name:
                self._references[index] = reference
                return

        self._references.append(reference)

    def get(self, name: str) -> ReferenceBase | None:
        """Return a registered reference by name, or `None` if it is absent."""
        for reference in self._references:
            if reference.name == name:
                return reference

        return None

    def all(self) -> tuple[ReferenceBase, ...]:
        """Return all registered references as an immutable tuple."""
        return tuple(self._references)

