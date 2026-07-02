"""Domain models for NAS Parser."""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path


@dataclass(slots=True)
class ProductRecord:
    """Represent a product as it moves through the NAS Parser pipeline."""

    price: Decimal
    quantity: Decimal
    color: str
    size: str
    source_file: Path = field(repr=False)
    source_sheet: str
    source_row: int
    parser_name: str
    sku: str | None = field(default=None)
    name: str | None = field(default=None)
    category: str | None = field(default=None)
    color_code: str | None = field(default=None)
    shape: str | None = field(default=None)
    fixation: str | None = field(default=None)
    cut: str | None = field(default=None)
