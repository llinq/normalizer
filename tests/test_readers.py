"""Tests for normalizer.readers utilities."""

from __future__ import annotations

import io

import openpyxl
import pytest

from normalizer.readers import (
    get_headers,
    is_formula,
    normalize_formula,
)


class TestGetHeaders:
    def _make_sheet(self, row: list):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(row)
        return ws

    def test_returns_column_header_map(self):
        ws = self._make_sheet(["Name", "Value", "Tax"])
        headers = get_headers(ws, header_row=1)
        assert headers[1] == "Name"
        assert headers[2] == "Value"
        assert headers[3] == "Tax"

    def test_strips_whitespace(self):
        ws = self._make_sheet(["  Name  ", " Value"])
        headers = get_headers(ws, header_row=1)
        assert headers[1] == "Name"
        assert headers[2] == "Value"

    def test_none_values_included(self):
        ws = self._make_sheet(["Name", None, "Tax"])
        headers = get_headers(ws, header_row=1)
        assert headers[2] is None


class TestIsFormula:
    def _make_cell(self, value):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = value
        return ws["A1"]

    def test_formula_detected(self):
        cell = self._make_cell("=SUM(A1:A10)")
        assert is_formula(cell) is True

    def test_plain_string_not_formula(self):
        cell = self._make_cell("hello")
        assert is_formula(cell) is False

    def test_number_not_formula(self):
        cell = self._make_cell(42)
        assert is_formula(cell) is False

    def test_none_not_formula(self):
        cell = self._make_cell(None)
        assert is_formula(cell) is False


class TestNormalizeFormula:
    def test_row_numbers_replaced(self):
        assert normalize_formula("=SUM(B2:B100)") == "=SUM(B#:B#)"

    def test_absolute_refs_normalised(self):
        result = normalize_formula("=$B$2+$C$3")
        assert "#" in result
        assert "2" not in result
        assert "3" not in result

    def test_function_name_preserved(self):
        result = normalize_formula("=VLOOKUP(A2,B2:C10,2,0)")
        assert "VLOOKUP" in result

    def test_empty_formula_passthrough(self):
        assert normalize_formula("") == ""

    def test_case_insensitive(self):
        assert normalize_formula("=sum(b2:b10)") == normalize_formula("=SUM(B2:B10)")
