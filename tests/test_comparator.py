"""Tests for normalizer.comparator and normalizer.models."""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

import pytest

from normalizer import compare
from normalizer.models import DivergenceType
from tests.fixtures import make_workbook


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _compare_from_buffers(template_data, user_data, **kwargs):
    """Write fixtures to tmp files and run compare()."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(template_data.read())
        tmpl_path = f.name

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(user_data.read())
        user_path = f.name

    try:
        return compare(tmpl_path, user_path, **kwargs)
    finally:
        os.unlink(tmpl_path)
        os.unlink(user_path)


# ---------------------------------------------------------------------------
# Sheet-level tests
# ---------------------------------------------------------------------------

class TestSheetComparison:
    def test_no_divergences_identical_sheets(self):
        data = {"Sheet1": [["Col A", "Col B"], [1, 2]]}
        result = _compare_from_buffers(
            make_workbook(data),
            make_workbook(data),
            check_formulas=False,
            check_data_types=False,
        )
        assert not result.has_divergences

    def test_missing_sheet(self):
        template = {"Sheet1": [["Col A"]], "Sheet2": [["Col X"]]}
        user = {"Sheet1": [["Col A"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=False,
        )
        missing = result.by_type(DivergenceType.MISSING_SHEET)
        assert len(missing) == 1
        assert missing[0].sheet == "Sheet2"

    def test_extra_sheet(self):
        template = {"Sheet1": [["Col A"]]}
        user = {"Sheet1": [["Col A"]], "ExtraSheet": [["X"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=False,
        )
        extra = result.by_type(DivergenceType.EXTRA_SHEET)
        assert len(extra) == 1
        assert extra[0].sheet == "ExtraSheet"

    def test_multiple_missing_and_extra_sheets(self):
        template = {"A": [["h"]], "B": [["h"]], "C": [["h"]]}
        user = {"A": [["h"]], "D": [["h"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=False,
        )
        missing_names = {d.sheet for d in result.by_type(DivergenceType.MISSING_SHEET)}
        extra_names = {d.sheet for d in result.by_type(DivergenceType.EXTRA_SHEET)}
        assert missing_names == {"B", "C"}
        assert extra_names == {"D"}


# ---------------------------------------------------------------------------
# Column-level tests
# ---------------------------------------------------------------------------

class TestColumnComparison:
    def test_missing_column(self):
        template = {"Sheet1": [["Name", "Value", "Tax"]]}
        user = {"Sheet1": [["Name", "Value"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=False,
        )
        missing = result.by_type(DivergenceType.MISSING_COLUMN)
        assert len(missing) == 1
        assert missing[0].template_value == "Tax"

    def test_extra_column(self):
        template = {"Sheet1": [["Name", "Value"]]}
        user = {"Sheet1": [["Name", "Value", "Extra"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=False,
        )
        extra = result.by_type(DivergenceType.EXTRA_COLUMN)
        assert len(extra) == 1
        assert extra[0].user_value == "Extra"

    def test_column_order_mismatch(self):
        template = {"Sheet1": [["Name", "Value", "Tax"]]}
        user = {"Sheet1": [["Tax", "Name", "Value"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=False,
        )
        order = result.by_type(DivergenceType.COLUMN_ORDER)
        assert len(order) == 1

    def test_no_column_divergence_when_identical(self):
        data = {"Sheet1": [["Name", "Value", "Tax"]]}
        result = _compare_from_buffers(
            make_workbook(data),
            make_workbook(data),
            check_formulas=False,
            check_data_types=False,
        )
        assert not result.has_divergences


# ---------------------------------------------------------------------------
# Formula tests
# ---------------------------------------------------------------------------

class TestFormulaComparison:
    def test_formula_missing_in_user(self):
        template = {"Sheet1": [["A", "B", "Total"], [1, 2, "=A2+B2"]]}
        user = {"Sheet1": [["A", "B", "Total"], [1, 2, 3]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=True,
            check_data_types=False,
        )
        missing_formula = result.by_type(DivergenceType.FORMULA_MISSING)
        assert len(missing_formula) == 1
        assert missing_formula[0].location == "C2"

    def test_formula_mismatch(self):
        template = {"Sheet1": [["A", "B", "Total"], [1, 2, "=A2+B2"]]}
        user = {"Sheet1": [["A", "B", "Total"], [1, 2, "=A2*B2"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=True,
            check_data_types=False,
        )
        mismatches = result.by_type(DivergenceType.FORMULA_MISMATCH)
        assert len(mismatches) == 1

    def test_unexpected_formula_in_user(self):
        template = {"Sheet1": [["A", "B"], [1, 2]]}
        user = {"Sheet1": [["A", "B"], [1, "=A2*2"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=True,
            check_data_types=False,
        )
        unexpected = result.by_type(DivergenceType.UNEXPECTED_FORMULA)
        assert len(unexpected) == 1

    def test_identical_formulas_no_divergence(self):
        data = {"Sheet1": [["A", "B", "Total"], [1, 2, "=A2+B2"]]}
        result = _compare_from_buffers(
            make_workbook(data),
            make_workbook(data),
            check_formulas=True,
            check_data_types=False,
        )
        assert not result.has_divergences

    def test_sheet_name_quotes_ignored(self):
        """Formulas that differ only in single-quoted sheet names are considered equal.

        Excel automatically wraps sheet names that start with a digit (or contain
        special characters) in single quotes when writing cross-sheet references.
        The template may store the unquoted form while the user file uses the
        quoted form, or vice-versa – both should be treated as identical.
        """
        template = {"Sheet1": [["A", "Result"],
                                [1, "=IF(1_DadosCadastrais!$B$38=\"\",\"\",1_DadosCadastrais!$B$38)"]]}
        user = {"Sheet1": [["A", "Result"],
                            [1, "=IF('1_DadosCadastrais'!$B$38=\"\",\"\",'1_DadosCadastrais'!$B$38)"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=True,
            check_data_types=False,
        )
        mismatches = result.by_type(DivergenceType.FORMULA_MISMATCH)
        assert len(mismatches) == 0

    def test_formula_row_number_difference_no_mismatch(self):
        """Formulas that differ only in row numbers are considered structurally identical."""
        template = {"Sheet1": [["A", "B", "Total"], [1, 2, "=A2+B2"], [3, 4, "=A3+B3"]]}
        # user uses slightly different row numbers (still structurally the same pattern)
        user = {"Sheet1": [["A", "B", "Total"], [1, 2, "=A2+B2"], [3, 4, "=A3+B3"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=True,
            check_data_types=False,
        )
        assert not result.has_divergences


# ---------------------------------------------------------------------------
# Data-type tests
# ---------------------------------------------------------------------------

class TestDataTypeComparison:
    def test_type_mismatch(self):
        template = {"Sheet1": [["Code", "Amount"], ["ABC123", 100.0]]}
        user = {"Sheet1": [["Code", "Amount"], [999, "not a number"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=True,
        )
        mismatches = result.by_type(DivergenceType.DATA_TYPE_MISMATCH)
        assert len(mismatches) >= 1

    def test_no_type_mismatch_when_same_types(self):
        template = {"Sheet1": [["Code", "Amount"], ["ABC", 10.0]]}
        user = {"Sheet1": [["Code", "Amount"], ["XYZ", 20.0]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=True,
        )
        assert not result.has_divergences


# ---------------------------------------------------------------------------
# to_dict / JSON serialisation
# ---------------------------------------------------------------------------

class TestSerialization:
    def test_to_dict_schema(self):
        template = {"Sheet1": [["Col A", "Col B"]], "Sheet2": [["X"]]}
        user = {"Sheet1": [["Col A"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=False,
        )
        d = result.to_dict()
        assert "has_divergences" in d
        assert "total" in d
        assert "divergences" in d
        assert d["total"] == len(d["divergences"])

    def test_json_serialisable(self):
        template = {"Sheet1": [["Col A", "Col B"]], "Sheet2": [["X"]]}
        user = {"Sheet1": [["Col A"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=False,
        )
        # Should not raise
        json_str = json.dumps(result.to_dict(), default=str)
        parsed = json.loads(json_str)
        assert isinstance(parsed["divergences"], list)

    def test_divergence_type_is_string_in_dict(self):
        template = {"Sheet1": [["A"]], "Sheet2": [["B"]]}
        user = {"Sheet1": [["A"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            check_data_types=False,
        )
        for div in result.to_dict()["divergences"]:
            assert isinstance(div["type"], str)
