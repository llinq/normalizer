"""Helpers to build in-memory .xlsx fixtures for tests."""

from __future__ import annotations

import io
from typing import Any

import openpyxl
from openpyxl import Workbook


def make_workbook(sheets: dict[str, list[list[Any]]]) -> io.BytesIO:
    """Return an in-memory .xlsx bytes buffer.

    ``sheets`` maps sheet name → list of rows, where each row is a list of
    cell values.  A value starting with ``=`` is treated as a formula.
    """
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    for sheet_name, rows in sheets.items():
        ws = wb.create_sheet(title=sheet_name)
        for row_values in rows:
            ws.append(row_values)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
