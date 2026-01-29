---
description: Validate formula specs against real Excel behavior
---

# Validate Formula Specs Against Excel

This workflow validates formula specifications by running them through a local Excel instance.

## Prerequisites

- Microsoft Excel installed locally
- Python with `openpyxl` and `pywin32` (Windows) or `appscript` (Mac)
- Formula specs in ODF format

## Steps

### 1. Review ODF Specification

Read the OpenDocument Formula (ODF) specification to understand:
- Formula syntax requirements
- Function definitions
- Expected behavior

### 2. Create Test Formulas

Add test formulas to `specs/odf-formulas/` directory in ODF format:

```xml
<formula name="SUM_basic">
  <expression>=SUM(1,2,3)</expression>
  <expected>6</expected>
  <description>Basic SUM with literal numbers</description>
</formula>
```

### 3. Generate Excel Test File

Run the spec-to-excel converter:

```bash
python scripts/generate_excel_tests.py specs/odf-formulas/ -o test-cases/generated/
```

This creates an Excel file with all formulas from specs.

### 4. Calculate in Excel

Run the Excel validation script:

```bash
python scripts/validate_with_excel.py test-cases/generated/formulas.xlsx
```

This script:
- Opens Excel file
- Lets Excel calculate all formulas
- Extracts calculated values
- Compares with expected results from specs
- Reports any mismatches

### 5. Update Specs

If Excel behavior differs from specs:
- Update the spec to match Excel's actual behavior
- Document any quirks or edge cases
- Add notes about Excel version tested

### 6. Commit Changes

```bash
git add specs/odf-formulas/
git commit -m "Updated specs based on Excel validation"
```

## Output

The validation script produces:
- `validation-report.json` - Detailed results
- `mismatches.txt` - List of spec/Excel differences
- `coverage-report.html` - Which functions are tested

## Continuous Validation

Run validation on every spec change:

```bash
# Watch for changes and auto-validate
python scripts/watch_and_validate.py
```
