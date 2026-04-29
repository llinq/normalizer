#!/usr/bin/env python3
"""CLI entry-point for the normalizer tool.

Usage
-----
    python main.py template.xlsx user.xlsx
    python main.py template.xlsx user.xlsx --header-row 2 --output result.json
    python main.py template.xlsx user.xlsx --no-formulas --no-types
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from normalizer import compare


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="normalizer",
        description="Compare a fiscal spreadsheet template against a user-filled file and report divergences.",
    )
    parser.add_argument("template", type=Path, help="Path to the template workbook (.xlsx).")
    parser.add_argument("user_file", type=Path, help="Path to the user-filled workbook (.xlsx).")
    parser.add_argument(
        "--header-row",
        type=int,
        default=1,
        metavar="N",
        help="Row number (1-based) that contains column headers (default: 1).",
    )
    parser.add_argument(
        "--no-formulas",
        action="store_true",
        default=False,
        help="Skip formula comparison.",
    )
    parser.add_argument(
        "--no-types",
        action="store_true",
        default=False,
        help="Skip data-type comparison.",
    )
    parser.add_argument(
        "--max-formula-rows",
        type=int,
        default=None,
        metavar="N",
        help="Limit formula checking to the first N data rows per sheet.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        metavar="FILE",
        help="Write JSON output to FILE instead of stdout.",
    )
    parser.add_argument(
        "--indent",
        type=int,
        default=2,
        metavar="N",
        help="JSON indentation level (default: 2).",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    for path, label in [(args.template, "template"), (args.user_file, "user_file")]:
        if not path.exists():
            print(f"Error: {label} file not found: {path}", file=sys.stderr)
            return 2
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            print(
                f"Warning: {label} file does not have an .xlsx/.xlsm extension: {path}",
                file=sys.stderr,
            )

    result = compare(
        template_path=args.template,
        user_path=args.user_file,
        header_row=args.header_row,
        check_formulas=not args.no_formulas,
        check_data_types=not args.no_types,
        max_formula_rows=args.max_formula_rows,
    )

    output_data = result.to_dict()
    json_str = json.dumps(output_data, indent=args.indent, ensure_ascii=False, default=str)

    if args.output:
        args.output.write_text(json_str, encoding="utf-8")
        print(f"Results written to {args.output}", file=sys.stderr)
    else:
        print(json_str)

    return 1 if result.has_divergences else 0


if __name__ == "__main__":
    sys.exit(main())
