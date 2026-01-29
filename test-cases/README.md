# Example Test Cases for Excel Formula Engine

This directory contains Excel files used for testing the formula engine.

## Test File Organization

### basic_formulas.xlsx
Simple formulas to test core functionality:
- Arithmetic operations
- Basic functions (SUM, AVERAGE)
- Cell references

### complex_formulas.xlsx
More complex scenarios:
- Nested functions
- Cross-sheet references
- Named ranges
- Multiple dependencies

### edge_cases.xlsx
Edge cases and error handling:
- Division by zero
- Circular references
- Error propagation
- Type coercion

## Creating Test Files

### Manual Creation

1. Open Excel
2. Create formulas in cells
3. Let Excel calculate results
4. Save file
5. Export formulas and values (see export script below)

### Export Script (VBA)

```vba
Sub ExportFormulasAndValues()
    Dim ws As Worksheet
    Dim cell As Range
    Dim output As String
    Dim jsonOutput As String
    
    jsonOutput = "{" & vbCrLf
    jsonOutput = jsonOutput & "  ""test_name"": """ & ActiveWorkbook.Name & """," & vbCrLf
    jsonOutput = jsonOutput & "  ""cells"": {" & vbCrLf
    
    Dim firstCell As Boolean
    firstCell = True
    
    For Each ws In ActiveWorkbook.Worksheets
        For Each cell In ws.UsedRange
            If Not IsEmpty(cell.Value) Then
                If Not firstCell Then
                    jsonOutput = jsonOutput & "," & vbCrLf
                End If
                firstCell = False
                
                Dim address As String
                address = ws.Name & "!" & cell.address(False, False)
                
                jsonOutput = jsonOutput & "    """ & address & """: {" & vbCrLf
                
                If cell.HasFormula Then
                    jsonOutput = jsonOutput & "      ""formula"": """ & Replace(cell.Formula, """", "\""") & """," & vbCrLf
                End If
                
                Dim cellValue As String
                If IsError(cell.Value) Then
                    cellValue = CStr(cell.Value)
                ElseIf IsNumeric(cell.Value) Then
                    cellValue = CStr(cell.Value)
                ElseIf VarType(cell.Value) = vbBoolean Then
                    cellValue = IIf(cell.Value, "true", "false")
                Else
                    cellValue = """" & Replace(CStr(cell.Value), """", "\""") & """"
                End If
                
                jsonOutput = jsonOutput & "      ""value"": " & cellValue & vbCrLf
                jsonOutput = jsonOutput & "    }"
            End If
        Next cell
    Next ws
    
    jsonOutput = jsonOutput & vbCrLf & "  }" & vbCrLf & "}"
    
    ' Write to file
    Dim fileName As String
    fileName = Replace(ActiveWorkbook.Name, ".xlsx", "_expected.json")
    
    Open ActiveWorkbook.Path & "\" & fileName For Output As #1
    Print #1, jsonOutput
    Close #1
    
    MsgBox "Exported to " & fileName
End Sub
```

## Test Case Format

### JSON Expected Results

```json
{
  "test_name": "basic_formulas.xlsx",
  "cells": {
    "Sheet1!A1": {
      "value": 10
    },
    "Sheet1!A2": {
      "value": 20
    },
    "Sheet1!A3": {
      "formula": "=A1+A2",
      "value": 30
    },
    "Sheet1!B1": {
      "formula": "=SUM(A1:A3)",
      "value": 60
    }
  }
}
```

## Sample Test Cases

### Test Case 1: Basic Arithmetic

**File**: basic_arithmetic.xlsx

| Cell | Formula | Expected |
|------|---------|----------|
| A1   | 10      | 10       |
| A2   | 20      | 20       |
| A3   | =A1+A2  | 30       |
| A4   | =A1*A2  | 200      |
| A5   | =A2/A1  | 2        |
| A6   | =A2^2   | 400      |

### Test Case 2: SUM Function

**File**: sum_function.xlsx

| Cell | Formula | Expected |
|------|---------|----------|
| A1   | 1       | 1        |
| A2   | 2       | 2        |
| A3   | 3       | 3        |
| A4   | "text"  | "text"   |
| A5   |         |          |
| B1   | =SUM(A1:A5) | 6    |
| B2   | =SUM(A1,A2,A3) | 6 |

### Test Case 3: IF Function

**File**: if_function.xlsx

| Cell | Formula | Expected |
|------|---------|----------|
| A1   | 10      | 10       |
| A2   | 5       | 5        |
| B1   | =IF(A1>A2,"Greater","Less") | "Greater" |
| B2   | =IF(A2>A1,"Greater","Less") | "Less" |
| B3   | =IF(A1=10,"Equal","Not Equal") | "Equal" |

### Test Case 4: VLOOKUP

**File**: vlookup.xlsx

**Sheet1 - Data**:
| A    | B     |
|------|-------|
| ID   | Value |
| 1    | Apple |
| 2    | Banana|
| 3    | Cherry|

**Sheet2 - Lookup**:
| Cell | Formula | Expected |
|------|---------|----------|
| A1   | 2       | 2        |
| B1   | =VLOOKUP(A1,Data!A:B,2,FALSE) | "Banana" |

### Test Case 5: Nested Functions

**File**: nested_functions.xlsx

| Cell | Formula | Expected |
|------|---------|----------|
| A1   | 5       | 5        |
| A2   | 10      | 10       |
| A3   | 15      | 15       |
| B1   | =IF(A1>0,SUM(A1:A3),0) | 30 |
| B2   | =ROUND(AVERAGE(A1:A3),0) | 10 |

### Test Case 6: Error Handling

**File**: errors.xlsx

| Cell | Formula | Expected |
|------|---------|----------|
| A1   | =1/0    | #DIV/0!  |
| A2   | =A1+5   | #DIV/0!  |
| B1   | =VLOOKUP("X",A:B,2,FALSE) | #N/A |
| B2   | =IFERROR(B1,"Not Found") | "Not Found" |

### Test Case 7: Circular Reference

**File**: circular.xlsx

| Cell | Formula | Expected Behavior |
|------|---------|-------------------|
| A1   | =B1+1   | Circular reference detected |
| B1   | =A1+1   | Circular reference detected |

### Test Case 8: Cross-Sheet References

**File**: cross_sheet.xlsx

**Sheet1**:
| Cell | Value |
|------|-------|
| A1   | 100   |

**Sheet2**:
| Cell | Formula | Expected |
|------|---------|----------|
| A1   | =Sheet1!A1*2 | 200 |

### Test Case 9: Text Functions

**File**: text_functions.xlsx

| Cell | Formula | Expected |
|------|---------|----------|
| A1   | "Hello" | "Hello"  |
| A2   | "World" | "World"  |
| B1   | =CONCATENATE(A1," ",A2) | "Hello World" |
| B2   | =LEFT(A1,2) | "He" |
| B3   | =LEN(A1) | 5 |
| B4   | =UPPER(A1) | "HELLO" |

### Test Case 10: Date Functions

**File**: date_functions.xlsx

| Cell | Formula | Expected (approx) |
|------|---------|-------------------|
| A1   | =TODAY() | Current date serial |
| A2   | =DATE(2024,1,15) | 45306 |
| A3   | =YEAR(A2) | 2024 |
| A4   | =MONTH(A2) | 1 |
| A5   | =DAY(A2) | 15 |

## Running Tests

### Python Example

```python
import openpyxl
import json
from excel_engine import FormulaEngine

def run_test_case(xlsx_file, expected_json):
    """Run a test case and verify results"""
    # Load expected results
    with open(expected_json) as f:
        expected = json.load(f)
    
    # Load and calculate workbook
    engine = FormulaEngine()
    workbook = engine.load(xlsx_file)
    engine.calculate(workbook)
    
    # Verify each cell
    failures = []
    for cell_address, cell_data in expected['cells'].items():
        if 'formula' in cell_data:
            expected_value = cell_data['value']
            actual_value = workbook.get_value(cell_address)
            
            if not values_match(actual_value, expected_value):
                failures.append({
                    'cell': cell_address,
                    'formula': cell_data['formula'],
                    'expected': expected_value,
                    'actual': actual_value
                })
    
    return failures

# Run all tests
test_cases = [
    ('basic_arithmetic.xlsx', 'basic_arithmetic_expected.json'),
    ('sum_function.xlsx', 'sum_function_expected.json'),
    # ... more test cases
]

for xlsx, json_file in test_cases:
    failures = run_test_case(xlsx, json_file)
    if failures:
        print(f"FAILED: {xlsx}")
        for failure in failures:
            print(f"  {failure}")
    else:
        print(f"PASSED: {xlsx}")
```

## Adding New Test Cases

1. Create Excel file with formulas
2. Verify results in Excel
3. Export using VBA script
4. Add to test suite
5. Document in this file

## Notes

- All test files should be in `.xlsx` format (Excel 2007+)
- Include both simple and complex cases
- Test edge cases and error conditions
- Compare results with actual Excel output
- Keep test files small and focused
