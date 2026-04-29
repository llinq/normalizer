"""Normalizer – fiscal spreadsheet comparison library."""

from .comparator import compare
from .models import ComparisonResult, Divergence, DivergenceType

__all__ = ["compare", "ComparisonResult", "Divergence", "DivergenceType"]
