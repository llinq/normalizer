"""Utilities for reading Excel workbooks."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl import Workbook
from openpyxl.cell.cell import Cell


def load_workbook_with_formulas(path: str | Path) -> Workbook:
    """Load a workbook keeping formula strings intact (data_only=False)."""
    return openpyxl.load_workbook(str(path), data_only=False)


def get_sheet_names(wb: Workbook) -> list[str]:
    """Return the sheet names of a workbook."""
    return list(wb.sheetnames)


def get_headers(sheet: Any, header_row: int = 1) -> dict[int, str | None]:
    """Return a mapping of {column_index: header_value} for a given row."""
    headers: dict[int, str | None] = {}
    for cell in sheet[header_row]:
        value = cell.value
        if value is not None:
            headers[cell.column] = str(value).strip()
        else:
            headers[cell.column] = None
    return headers


def cell_address(row: int, col: int) -> str:
    """Convert (row, col) to an Excel-style address like A1."""
    return f"{openpyxl.utils.get_column_letter(col)}{row}"


def is_formula(cell: Cell) -> bool:
    """Return True if the cell contains a formula."""
    return isinstance(cell.value, str) and cell.value.startswith("=")


def normalize_formula(formula: str) -> str:
    """Normalise a formula by collapsing absolute/relative row references.

    This makes structural formula comparison possible even when the user's
    spreadsheet has more or fewer data rows (e.g. ``=SUM(B2:B100)`` vs
    ``=SUM(B2:B50)``).  Column references and function names are preserved
    so that structural changes (wrong column, missing function, etc.) are
    still detected.

    Single quotes around sheet names (added by Excel when a sheet name starts
    with a digit or contains special characters) are also stripped so that
    ``Sheet1!A1`` and ``'Sheet1'!A1`` compare as equal.
    """
    if not formula:
        return formula
    # Upper-case for case-insensitive comparison
    formula = formula.upper()
    # Strip single quotes around sheet names (e.g. '1_Sheet'!A1 → 1_SHEET!A1)
    formula = re.sub(r"'([^']+)'!", r"\1!", formula)
    # Replace row numbers inside cell references (e.g. B12 → B#, $B$12 → $B$#)
    formula = re.sub(r"(\$?[A-Z]+)\$?[0-9]+", r"\1#", formula)
    return formula
