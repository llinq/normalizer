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


def normalize_formula(formula: str, current_col: int = 0) -> str:
    """Normalise a formula for structural comparison.

    The following transformations are applied:

    * Case-folded to upper-case so comparisons are case-insensitive.
    * Single quotes around sheet names are stripped (Excel adds them when the
      name starts with a digit or contains special characters), so
      ``'1_Sheet'!A1`` and ``1_SHEET!A1`` compare as equal.
    * Row numbers inside cell references are replaced with ``#`` so that the
      same structural formula in different rows still compares equal
      (e.g. ``=SUM(B2:B100)`` vs ``=SUM(B2:B50)``).
    * When *current_col* is given (> 0), non-absolute same-sheet column
      references are converted to signed offsets relative to *current_col*
      (e.g. ``=D2`` in column C becomes ``=[+1]#``).  This lets equivalent
      formulas compare equal even when the column containing the formula
      differs between the template and the user workbook.  Cross-sheet
      references and absolute column references (``$D2``) are left as-is so
      intentional fixed anchors are still detected as mismatches.
    """
    if not formula:
        return formula
    # Upper-case for case-insensitive comparison
    formula = formula.upper()
    # Strip single quotes around sheet names (e.g. '1_Sheet'!A1 → 1_SHEET!A1)
    formula = re.sub(r"'([^']+)'!", r"\1!", formula)

    if current_col > 0:
        # ------------------------------------------------------------------ #
        # Step 1 – Normalise cross-sheet cell/range references first so their
        # column letters are NOT relativised in step 2.
        # Pattern matches:  SHEETNAME!<cell_or_range>
        # e.g.  1_SHEET!$B$38  or  SHEET!A1:B2
        # ------------------------------------------------------------------ #
        def _normalize_xref(m: re.Match) -> str:
            # Replace row numbers inside the cross-sheet ref, keep columns
            return re.sub(r"(\$?[A-Z]+)\$?[0-9]+", r"\1#", m.group(0))

        formula = re.sub(
            r"[A-Z0-9_]+!\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?",
            _normalize_xref,
            formula,
        )

        # ------------------------------------------------------------------ #
        # Step 2 – Relativise non-absolute same-sheet column references.
        # A relative column ref is NOT preceded by $ (which marks an absolute
        # column) and NOT preceded by ! (which marks a cross-sheet ref, already
        # handled above).
        # ------------------------------------------------------------------ #
        def _relativize(m: re.Match) -> str:
            col_num = openpyxl.utils.column_index_from_string(m.group(1))
            offset = col_num - current_col
            return f"[{offset:+d}]#"

        formula = re.sub(r"(?<![!$])([A-Z]+)\$?[0-9]+", _relativize, formula)

        # Step 3 – Replace row numbers in any remaining absolute column refs
        # (e.g. $D$2 → $D# that survived steps 1-2).
        formula = re.sub(r"(\$[A-Z]+)\$?[0-9]+", r"\1#", formula)
    else:
        # Without column context: just replace row numbers
        formula = re.sub(r"(\$?[A-Z]+)\$?[0-9]+", r"\1#", formula)

    return formula
