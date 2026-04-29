"""Tests for the CLI entry-point (main.py)."""

from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path

import pytest

from tests.fixtures import make_workbook


def _write_workbook(sheets, suffix=".xlsx") -> str:
    buf = make_workbook(sheets)
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as f:
        f.write(buf.read())
        return f.name


class TestCLI:
    def _run(self, args: list[str]):
        from main import main
        return main(args)

    def test_exit_0_no_divergences(self):
        data = {"Sheet1": [["A", "B"], [1, 2]]}
        tmpl = _write_workbook(data)
        user = _write_workbook(data)
        try:
            rc = self._run([tmpl, user, "--no-formulas"])
            assert rc == 0
        finally:
            os.unlink(tmpl)
            os.unlink(user)

    def test_exit_1_with_divergences(self):
        tmpl = _write_workbook({"Sheet1": [["A"]], "Sheet2": [["B"]]})
        user = _write_workbook({"Sheet1": [["A"]]})
        try:
            rc = self._run([str(tmpl), str(user), "--no-formulas"])
            assert rc == 1
        finally:
            os.unlink(tmpl)
            os.unlink(user)

    def test_output_is_valid_json(self, capsys):
        data = {"Sheet1": [["A", "B"], [1, 2]]}
        tmpl = _write_workbook(data)
        user = _write_workbook(data)
        try:
            self._run([tmpl, user, "--no-formulas"])
            captured = capsys.readouterr()
            parsed = json.loads(captured.out)
            assert "divergences" in parsed
        finally:
            os.unlink(tmpl)
            os.unlink(user)

    def test_output_file(self, tmp_path):
        data = {"Sheet1": [["A", "B"], [1, 2]]}
        tmpl = _write_workbook(data)
        user = _write_workbook(data)
        out_file = str(tmp_path / "result.json")
        try:
            self._run([tmpl, user, "--no-formulas", "--output", out_file])
            content = Path(out_file).read_text(encoding="utf-8")
            parsed = json.loads(content)
            assert "divergences" in parsed
        finally:
            os.unlink(tmpl)
            os.unlink(user)

    def test_missing_file_returns_2(self):
        rc = self._run(["nonexistent_template.xlsx", "nonexistent_user.xlsx"])
        assert rc == 2

    def test_check_types_flag_enables_type_comparison(self):
        """--check-types enables data-type divergences that are off by default."""
        template = _write_workbook({"Sheet1": [["Code", "Amount"], ["ABC123", 100.0]]})
        user = _write_workbook({"Sheet1": [["Code", "Amount"], [999, "not a number"]]})
        try:
            # Without the flag: no divergences from type checking
            rc_no_flag = self._run([template, user, "--no-formulas"])
            # With the flag: type divergences should be detected
            rc_with_flag = self._run([template, user, "--no-formulas", "--check-types"])
            assert rc_no_flag == 0
            assert rc_with_flag == 1
        finally:
            os.unlink(template)
            os.unlink(user)
