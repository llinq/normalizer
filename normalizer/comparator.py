"""Core comparison logic between a template workbook and a user-filled workbook."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .models import ComparisonResult, Divergence, DivergenceType
from .readers import (
    cell_address,
    get_headers,
    get_sheet_names,
    is_formula,
    load_workbook_with_formulas,
    normalize_formula,
)


def compare(
    template_path: str | Path,
    user_path: str | Path,
    header_row: int = 1,
    check_formulas: bool = True,
    check_data_types: bool = True,
    max_formula_rows: int | None = None,
) -> ComparisonResult:
    """Compare *template* spreadsheet against *user* spreadsheet.

    Parameters
    ----------
    template_path:
        Path to the template / model workbook.
    user_path:
        Path to the user-filled workbook.
    header_row:
        Row index (1-based) that contains column headers.  Defaults to 1.
    check_formulas:
        When True, cells that contain formulas in the template are compared
        against the corresponding cells in the user workbook.
    check_data_types:
        When True, the data type of non-formula cells in the template is
        compared against the user workbook.
    max_formula_rows:
        If given, formula checking stops after this many data rows per sheet
        (useful for very large sheets).  ``None`` means check all rows.

    Returns
    -------
    ComparisonResult
        Object containing all detected divergences.
    """
    result = ComparisonResult()

    template_wb = load_workbook_with_formulas(template_path)
    user_wb = load_workbook_with_formulas(user_path)

    template_sheets = get_sheet_names(template_wb)
    user_sheets = get_sheet_names(user_wb)

    _compare_sheets(result, template_sheets, user_sheets)

    common_sheets = [s for s in template_sheets if s in user_sheets]

    for sheet_name in common_sheets:
        tmpl_sheet = template_wb[sheet_name]
        user_sheet = user_wb[sheet_name]

        tmpl_headers = get_headers(tmpl_sheet, header_row)
        user_headers = get_headers(user_sheet, header_row)

        _compare_columns(result, sheet_name, tmpl_headers, user_headers)

        if check_formulas or check_data_types:
            _compare_cells(
                result,
                sheet_name,
                tmpl_sheet,
                user_sheet,
                tmpl_headers,
                user_headers,
                header_row,
                check_formulas,
                check_data_types,
                max_formula_rows,
            )

    return result


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _compare_sheets(
    result: ComparisonResult,
    template_sheets: list[str],
    user_sheets: list[str],
) -> None:
    template_set = set(template_sheets)
    user_set = set(user_sheets)

    for sheet in sorted(template_set - user_set):
        result.divergences.append(
            Divergence(
                type=DivergenceType.MISSING_SHEET,
                sheet=sheet,
                description=f"Sheet '{sheet}' is present in the template but missing in the user file.",
            )
        )

    for sheet in sorted(user_set - template_set):
        result.divergences.append(
            Divergence(
                type=DivergenceType.EXTRA_SHEET,
                sheet=sheet,
                description=f"Sheet '{sheet}' exists in the user file but is not part of the template.",
            )
        )


def _compare_columns(
    result: ComparisonResult,
    sheet_name: str,
    tmpl_headers: dict[int, str | None],
    user_headers: dict[int, str | None],
) -> None:
    # Only consider named (non-None) headers
    tmpl_named = {v: k for k, v in tmpl_headers.items() if v is not None}
    user_named = {v: k for k, v in user_headers.items() if v is not None}

    tmpl_set = set(tmpl_named)
    user_set = set(user_named)

    for col in sorted(tmpl_set - user_set):
        result.divergences.append(
            Divergence(
                type=DivergenceType.MISSING_COLUMN,
                sheet=sheet_name,
                description=(
                    f"Column '{col}' is present in the template sheet '{sheet_name}' "
                    "but missing in the user file."
                ),
                template_value=col,
            )
        )

    for col in sorted(user_set - tmpl_set):
        result.divergences.append(
            Divergence(
                type=DivergenceType.EXTRA_COLUMN,
                sheet=sheet_name,
                description=(
                    f"Column '{col}' exists in the user sheet '{sheet_name}' "
                    "but is not part of the template."
                ),
                user_value=col,
            )
        )

    # Check column order for columns that exist in both
    common_cols = [c for c in tmpl_named if c in user_named]
    tmpl_order = sorted(common_cols, key=lambda c: tmpl_named[c])
    user_order = sorted(common_cols, key=lambda c: user_named[c])

    if tmpl_order != user_order:
        result.divergences.append(
            Divergence(
                type=DivergenceType.COLUMN_ORDER,
                sheet=sheet_name,
                description=(
                    f"Column order in sheet '{sheet_name}' differs from the template."
                ),
                template_value=tmpl_order,
                user_value=user_order,
            )
        )


def _compare_cells(
    result: ComparisonResult,
    sheet_name: str,
    tmpl_sheet: Any,
    user_sheet: Any,
    tmpl_headers: dict[int, str | None],
    user_headers: dict[int, str | None],
    header_row: int,
    check_formulas: bool,
    check_data_types: bool,
    max_formula_rows: int | None,
) -> None:
    # Build a mapping from template column index → user column index
    # using header names so we compare semantically equivalent columns
    tmpl_col_to_name = {k: v for k, v in tmpl_headers.items() if v is not None}
    user_name_to_col = {v: k for k, v in user_headers.items() if v is not None}

    col_mapping: dict[int, int] = {}
    for tmpl_col, name in tmpl_col_to_name.items():
        if name in user_name_to_col:
            col_mapping[tmpl_col] = user_name_to_col[name]

    tmpl_max_row = tmpl_sheet.max_row
    user_max_row = user_sheet.max_row

    data_start = header_row + 1
    rows_to_check = range(data_start, tmpl_max_row + 1)

    if max_formula_rows is not None:
        rows_to_check = range(data_start, min(tmpl_max_row, data_start + max_formula_rows - 1) + 1)

    for row in rows_to_check:
        for tmpl_col, user_col in col_mapping.items():
            tmpl_cell = tmpl_sheet.cell(row=row, column=tmpl_col)
            tmpl_addr = cell_address(row, tmpl_col)

            if row > user_max_row:
                # User sheet has fewer rows than the template – skip cell checks
                # (row count differences are not flagged as structural errors)
                continue

            user_cell = user_sheet.cell(row=row, column=user_col)

            tmpl_is_formula = is_formula(tmpl_cell)
            user_is_formula = is_formula(user_cell)

            if check_formulas:
                if tmpl_is_formula and not user_is_formula:
                    result.divergences.append(
                        Divergence(
                            type=DivergenceType.FORMULA_MISSING,
                            sheet=sheet_name,
                            location=tmpl_addr,
                            description=(
                                f"Cell {tmpl_addr} in sheet '{sheet_name}' has a formula "
                                "in the template but is missing one in the user file."
                            ),
                            template_value=tmpl_cell.value,
                            user_value=user_cell.value,
                        )
                    )
                    continue

                if tmpl_is_formula and user_is_formula:
                    tmpl_norm = normalize_formula(tmpl_cell.value)
                    user_norm = normalize_formula(user_cell.value)
                    if tmpl_norm != user_norm:
                        result.divergences.append(
                            Divergence(
                                type=DivergenceType.FORMULA_MISMATCH,
                                sheet=sheet_name,
                                location=tmpl_addr,
                                description=(
                                    f"Formula mismatch at {tmpl_addr} in sheet '{sheet_name}'."
                                ),
                                template_value=tmpl_cell.value,
                                user_value=user_cell.value,
                            )
                        )
                    continue

                if not tmpl_is_formula and user_is_formula:
                    result.divergences.append(
                        Divergence(
                            type=DivergenceType.UNEXPECTED_FORMULA,
                            sheet=sheet_name,
                            location=tmpl_addr,
                            description=(
                                f"Cell {tmpl_addr} in sheet '{sheet_name}' should contain "
                                "a plain value but the user file has a formula."
                            ),
                            template_value=tmpl_cell.value,
                            user_value=user_cell.value,
                        )
                    )
                    continue

            if check_data_types and not tmpl_is_formula and not user_is_formula:
                tmpl_type = _cell_type_label(tmpl_cell.value)
                user_type = _cell_type_label(user_cell.value)
                # Only report when template has a value (sample data) and types differ
                if tmpl_cell.value is not None and tmpl_type != user_type and user_cell.value is not None:
                    result.divergences.append(
                        Divergence(
                            type=DivergenceType.DATA_TYPE_MISMATCH,
                            sheet=sheet_name,
                            location=tmpl_addr,
                            description=(
                                f"Data type mismatch at {tmpl_addr} in sheet '{sheet_name}': "
                                f"template has {tmpl_type}, user has {user_type}."
                            ),
                            template_value=tmpl_type,
                            user_value=user_type,
                        )
                    )


def _cell_type_label(value: Any) -> str:
    if value is None:
        return "empty"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, int):
        return "integer"
    if isinstance(value, float):
        return "float"
    if isinstance(value, str):
        return "string"
    # datetime, date, time from openpyxl
    return type(value).__name__
