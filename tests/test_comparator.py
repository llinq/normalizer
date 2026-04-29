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
        """Same-structure formula at different rows compares equal (same relative refs)."""
        template = {"Sheet1": [["A", "B", "Total"], [1, 2, "=A2+B2"], [3, 4, "=A3+B3"]]}
        user = {"Sheet1": [["A", "B", "Total"], [1, 2, "=A2+B2"], [3, 4, "=A3+B3"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=True,
            check_data_types=False,
        )
        assert not result.has_divergences

    def test_formula_row_ref_change_detected(self):
        """Changing the row a formula references (e.g. A11→A10 in row 11) must be detected.

        Real-world scenario: user accidentally edits a row-guard formula from
        =SE(A11="";...;PROCV(E11;...)) to =SE(A10="";...;PROCV(E11;...)).
        The A11→A10 change shifts the relative row offset from [0] to [-1] and
        must trigger a FORMULA_MISMATCH.
        """
        # Row 2 in column C: template references A2 (same row), user changed to A1
        template = {"Sheet1": [["A", "B", "C"], ["x", "y", '=IF(A2="","",B2)']]}
        user = {"Sheet1": [["A", "B", "C"], ["x", "y", '=IF(A1="","",B2)']]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=True,
            check_data_types=False,
        )
        mismatches = result.by_type(DivergenceType.FORMULA_MISMATCH)
        assert len(mismatches) == 1
        assert mismatches[0].location == "C2"

    def test_formula_at_different_column_positions_no_mismatch(self):
        """Formulas with identical structure at different column positions compare equal.

        Template column C (3) has =D2&E2 (refs at C+1=D, C+2=E → offsets +1, +2).
        User column Q (17) has =R2&S2 (refs at Q+1=R, Q+2=S → offsets +1, +2).
        Both express the same pattern: concatenate the next two columns.
        """
        # Template: headers at cols 1-5 (A-E)
        template_headers = ["ID", "Name", "Combined", "First", "Last"]
        template_data = [template_headers, [1, "x", "=D2&E2", "a", "b"]]

        # User: 14 filler Nones so that 'Combined' lands at col 17 (Q),
        # 'First' at R (18), 'Last' at S (19).
        # Formula =R2&S2 in col Q: R=18=Q+1, S=19=Q+2 → offsets +1, +2 ✓
        filler_count = 14
        user_headers = ["ID", "Name"] + [None] * filler_count + ["Combined", "First", "Last"]
        user_row = [1, "x"] + [None] * filler_count + ["=R2&S2", "a", "b"]
        user_data = [user_headers, user_row]

        result = _compare_from_buffers(
            make_workbook({"Sheet1": template_data}),
            make_workbook({"Sheet1": user_data}),
            check_formulas=True,
            check_data_types=False,
        )
        mismatches = result.by_type(DivergenceType.FORMULA_MISMATCH)
        assert len(mismatches) == 0


# ---------------------------------------------------------------------------
# Empty-row tests
# ---------------------------------------------------------------------------

class TestEmptyRowSkipping:
    def test_empty_rows_not_reported_as_formula_missing(self):
        """Completely empty template rows are silently skipped."""
        template = {"Sheet1": [["A", "Formula"],
                                [1, "=A2*2"],
                                [None, None],   # empty row
                                [3, "=A4*2"]]}
        user = {"Sheet1": [["A", "Formula"],
                            [1, "=A2*2"],
                            [None, None],
                            [3, "=A4*2"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=True,
            check_data_types=False,
        )
        assert not result.has_divergences

    def test_empty_template_row_not_compared_to_user_formula(self):
        """A completely empty template row is skipped even if user has data there."""
        template = {"Sheet1": [["A", "B"],
                                [1, "=A2"],
                                [None, None]]}
        user = {"Sheet1": [["A", "B"],
                            [1, "=A2"],
                            [5, "=A3*3"]]}  # user has formula on empty-template row
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=True,
            check_data_types=False,
        )
        # The empty template row should be ignored – no UNEXPECTED_FORMULA reported
        unexpected = result.by_type(DivergenceType.UNEXPECTED_FORMULA)
        assert len(unexpected) == 0

    def test_none_template_cell_with_user_formula_not_flagged(self):
        """User formula in a cell that is None in the template must not raise UNEXPECTED_FORMULA.

        Scenario: template has only 2 data rows with a formula in column B.
        User dragged the formula down, so B11 (which the template leaves blank) has a formula.
        The partially-populated row means the whole-row empty check doesn't skip row 11,
        but the per-cell None guard should suppress the UNEXPECTED_FORMULA flag.
        """
        # Row 11 in the template: col A has a value but col B is None
        template_rows = [
            ["A", "B"],
            [1, "=A2*2"],
            [2, "=A3*2"],
        ] + [[i, None] for i in range(4, 12)]  # rows 4-11: A has a value, B is None
        # User dragged the formula down through row 11
        user_rows = [
            ["A", "B"],
            [1, "=A2*2"],
            [2, "=A3*2"],
        ] + [[i, f"=A{i+1}*2"] for i in range(4, 12)]
        result = _compare_from_buffers(
            make_workbook({"Sheet1": template_rows}),
            make_workbook({"Sheet1": user_rows}),
            check_formulas=True,
            check_data_types=False,
        )
        unexpected = result.by_type(DivergenceType.UNEXPECTED_FORMULA)
        assert len(unexpected) == 0


    def test_formula_missing_not_flagged_when_user_row_empty(self):
        """FORMULA_MISSING must not fire when the user cell is None.

        Real-world scenario: the template has a pre-filled formula in every row
        up to a large maximum (e.g. row 4358). The user only filled a few rows,
        leaving the rest completely empty. Reporting FORMULA_MISSING for every
        unpopulated template row is a false positive.
        """
        # Template: 3 rows with formulas in column B
        template = {
            "Sheet1": [
                ["A", "B"],
                [1, "=IF(A2=\"\",\"\",A2*2)"],
                [None, "=IF(A3=\"\",\"\",A3*2)"],  # row 3 – A is empty, B has guard formula
            ]
        }
        # User only filled row 2; row 3 is entirely empty
        user = {
            "Sheet1": [
                ["A", "B"],
                [1, "=IF(A2=\"\",\"\",A2*2)"],
                [None, None],
            ]
        }
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=True,
            check_data_types=False,
        )
        missing = result.by_type(DivergenceType.FORMULA_MISSING)
        assert len(missing) == 0

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

    def test_check_data_types_disabled_by_default(self):
        """compare() must not report type divergences when check_data_types is not given."""
        template = {"Sheet1": [["Code", "Amount"], ["ABC123", 100.0]]}
        user = {"Sheet1": [["Code", "Amount"], [999, "not a number"]]}
        result = _compare_from_buffers(
            make_workbook(template),
            make_workbook(user),
            check_formulas=False,
            # check_data_types intentionally omitted – should default to False
        )
        mismatches = result.by_type(DivergenceType.DATA_TYPE_MISMATCH)
        assert len(mismatches) == 0


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
