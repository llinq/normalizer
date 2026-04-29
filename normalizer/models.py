"""Data models for spreadsheet comparison divergences."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any


class DivergenceType(str, Enum):
    # Sheet-level
    MISSING_SHEET = "MISSING_SHEET"
    EXTRA_SHEET = "EXTRA_SHEET"

    # Column-level
    MISSING_COLUMN = "MISSING_COLUMN"
    EXTRA_COLUMN = "EXTRA_COLUMN"
    COLUMN_ORDER = "COLUMN_ORDER"

    # Cell-level
    FORMULA_MISSING = "FORMULA_MISSING"
    FORMULA_MISMATCH = "FORMULA_MISMATCH"
    UNEXPECTED_FORMULA = "UNEXPECTED_FORMULA"
    DATA_TYPE_MISMATCH = "DATA_TYPE_MISMATCH"


@dataclass
class Divergence:
    """Represents a single divergence between the template and user spreadsheet."""

    type: DivergenceType
    sheet: str | None
    description: str
    location: str | None = None
    template_value: Any = None
    user_value: Any = None

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["type"] = self.type.value
        return d


@dataclass
class ComparisonResult:
    """Aggregated result of comparing two spreadsheets."""

    divergences: list[Divergence] = field(default_factory=list)

    @property
    def has_divergences(self) -> bool:
        return bool(self.divergences)

    def by_type(self, dtype: DivergenceType) -> list[Divergence]:
        return [d for d in self.divergences if d.type == dtype]

    def to_dict(self) -> dict[str, Any]:
        return {
            "has_divergences": self.has_divergences,
            "total": len(self.divergences),
            "divergences": [d.to_dict() for d in self.divergences],
        }
