"""Shared extractor result types."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class ExtractResult:
    """Successful image entries plus reasons for skipped candidates."""

    entries: list[tuple[str, Path]] = field(default_factory=list)
    skip_reasons: list[str] = field(default_factory=list)
