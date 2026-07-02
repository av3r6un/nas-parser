"""Application configuration for NAS Parser."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(slots=True)
class AppConfig:
    """Configuration values used by the NAS Parser pipeline."""

    input_dir: Path = field(default_factory=lambda: Path("input"))
    reference_dir: Path = field(default_factory=lambda: Path("reference"))
    output_dir: Path = field(default_factory=lambda: Path("output"))
    output_file: Path = field(default_factory=lambda: Path("output/catalog.xlsx"))
    logs_dir: Path = field(default_factory=lambda: Path("logs"))
