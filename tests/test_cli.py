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
            rc = self._run([tmpl, user, "--no-formulas", "--no-types"])
            assert rc == 0
        finally:
            os.unlink(tmpl)
            os.unlink(user)

    def test_exit_1_with_divergences(self):
        tmpl = _write_workbook({"Sheet1": [["A"]], "Sheet2": [["B"]]})
        user = _write_workbook({"Sheet1": [["A"]]})
        try:
            rc = self._run([str(tmpl), str(user), "--no-formulas", "--no-types"])
            assert rc == 1
        finally:
            os.unlink(tmpl)
            os.unlink(user)

    def test_output_is_valid_json(self, capsys):
        data = {"Sheet1": [["A", "B"], [1, 2]]}
        tmpl = _write_workbook(data)
        user = _write_workbook(data)
        try:
            self._run([tmpl, user, "--no-formulas", "--no-types"])
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
            self._run([tmpl, user, "--no-formulas", "--no-types", "--output", out_file])
            content = Path(out_file).read_text(encoding="utf-8")
            parsed = json.loads(content)
            assert "divergences" in parsed
        finally:
            os.unlink(tmpl)
            os.unlink(user)

    def test_missing_file_returns_2(self):
        rc = self._run(["nonexistent_template.xlsx", "nonexistent_user.xlsx"])
        assert rc == 2
