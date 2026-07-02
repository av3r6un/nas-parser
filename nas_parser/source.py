"""Domain model for a source row in NAS Parser."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class SourceRow:
    """Represent a raw source row with origin metadata."""

    values: tuple[object, ...]
    source_file: Path
    source_sheet: str
    source_row: int
