# normalizer

Python application to validate and compare two fiscal Excel spreadsheets — a **template** (model) and a **user-filled** file — and return all structural divergences as JSON.

## Problem

Users download a template spreadsheet, fill it in, and upload it to the system. Sometimes they accidentally alter the template's structure (delete columns, remove formulas, rename sheets, etc.), causing the system to reject the file as "corrupted". This tool detects those divergences automatically so the user gets clear, actionable feedback.

## Features

| Check | Description |
|---|---|
| Missing/extra sheets | Sheets present in one file but not the other |
| Missing/extra columns | Column headers added or removed |
| Column order | Headers reordered relative to the template |
| Formula missing | A cell that should have a formula has plain data |
| Formula mismatch | A formula was altered (different function/column reference) |
| Unexpected formula | A cell that should have a plain value has a formula |
| Data type mismatch | Sample values in the template have a different type than the user's values |

## Requirements

- Python 3.10+
- openpyxl >= 3.1.0

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Command line

```bash
# Basic comparison (prints JSON to stdout)
python main.py template.xlsx user_file.xlsx

# Write result to a file
python main.py template.xlsx user_file.xlsx --output result.json

# Use a different header row (e.g. row 2)
python main.py template.xlsx user_file.xlsx --header-row 2

# Skip formula or data-type checks
python main.py template.xlsx user_file.xlsx --no-formulas --no-types

# Limit formula check to first 500 data rows (faster for large sheets)
python main.py template.xlsx user_file.xlsx --max-formula-rows 500
```

Exit code is **0** when no divergences are found, **1** when divergences exist, and **2** when a file is not found.

### Python API

```python
from normalizer import compare

result = compare("template.xlsx", "user_file.xlsx")

if result.has_divergences:
    import json
    print(json.dumps(result.to_dict(), indent=2, ensure_ascii=False))
```

### JSON output format

```json
{
  "has_divergences": true,
  "total": 3,
  "divergences": [
    {
      "type": "MISSING_SHEET",
      "sheet": "Impostos",
      "description": "Sheet 'Impostos' is present in the template but missing in the user file.",
      "location": null,
      "template_value": null,
      "user_value": null
    },
    {
      "type": "FORMULA_MISSING",
      "sheet": "Produtos",
      "description": "Cell C2 in sheet 'Produtos' has a formula in the template but is missing one in the user file.",
      "location": "C2",
      "template_value": "=A2*B2",
      "user_value": 42
    }
  ]
}
```

## Divergence types

| Type | Meaning |
|---|---|
| `MISSING_SHEET` | Sheet in template not found in user file |
| `EXTRA_SHEET` | Sheet in user file not in template |
| `MISSING_COLUMN` | Column header in template not found in user sheet |
| `EXTRA_COLUMN` | Column header in user sheet not in template |
| `COLUMN_ORDER` | Columns exist but in a different order |
| `FORMULA_MISSING` | Template cell has a formula; user cell does not |
| `FORMULA_MISMATCH` | Both cells have formulas but with different structure |
| `UNEXPECTED_FORMULA` | Template cell has plain data; user cell has a formula |
| `DATA_TYPE_MISMATCH` | Cell value types differ (e.g. string vs number) |

## Running tests

```bash
pip install pytest
pytest
```
