# Testing and Validation Strategy

## Overview

Testing an Excel formula engine requires validating against Excel's actual behavior. This document outlines the comprehensive testing approach.

## Test Pyramid

```
         /\
        /  \  End-to-End Tests (Real Excel Files)
       /____\
      /      \  Integration Tests (Multi-component)
     /________\
    /          \  Unit Tests (Individual Functions)
   /____________\
```

## 1. Unit Tests

### Function Tests

Test each Excel function independently:

```python
# test_functions.py

def test_sum_numbers():
    """SUM with simple numbers"""
    assert SUM(1, 2, 3) == 6

def test_sum_range():
    """SUM with range containing mixed types"""
    range_data = create_range([1, "text", 3, True, None])
    assert SUM(range_data) == 4  # 1 + 3, ignores text and empty

def test_sum_error_propagation():
    """SUM propagates errors"""
    result = SUM(1, ExcelError("#DIV/0!"), 3)
    assert is_error(result, "#DIV/0!")

def test_vlookup_exact_match():
    """VLOOKUP with exact match"""
    table = [["A", 1], ["B", 2], ["C", 3]]
    assert VLOOKUP("B", table, 2, False) == 2

def test_vlookup_not_found():
    """VLOOKUP returns #N/A when not found"""
    table = [["A", 1], ["B", 2]]
    result = VLOOKUP("C", table, 2, False)
    assert is_error(result, "#N/A")
```

### Parser Tests

```python
# test_parser.py

def test_parse_number():
    ast = parse("=42")
    assert isinstance(ast, NumberNode)
    assert ast.value == 42

def test_parse_arithmetic():
    ast = parse("=1+2*3")
    assert isinstance(ast, BinaryOpNode)
    assert ast.operator == '+'
    # Verify precedence: 1 + (2 * 3)
    assert isinstance(ast.right, BinaryOpNode)
    assert ast.right.operator == '*'

def test_parse_function_call():
    ast = parse("=SUM(A1:A10)")
    assert isinstance(ast, FunctionCallNode)
    assert ast.function_name == "SUM"
    assert len(ast.arguments) == 1
    assert isinstance(ast.arguments[0], RangeReferenceNode)

def test_parse_nested_functions():
    ast = parse("=IF(A1>0, SUM(B:B), 0)")
    assert isinstance(ast, FunctionCallNode)
    assert ast.function_name == "IF"
    assert len(ast.arguments) == 3
```

### Dependency Graph Tests

```python
# test_dependency_graph.py

def test_simple_chain():
    """A1 -> B1 -> C1"""
    graph = build_graph({
        'A1': '=10',
        'B1': '=A1*2',
        'C1': '=B1+5'
    })
    order = graph.calculation_order
    assert order.index('A1') < order.index('B1')
    assert order.index('B1') < order.index('C1')

def test_circular_reference_detection():
    """Detect A1 <-> B1 cycle"""
    graph = build_graph({
        'A1': '=B1+1',
        'B1': '=A1+1'
    })
    assert len(graph.circular_refs) == 1
    cycle = graph.circular_refs[0]
    assert set(cycle) == {'A1', 'B1'}

def test_diamond_dependency():
    """
         D1
        /  \
       B1  C1
        \  /
         A1
    """
    graph = build_graph({
        'A1': '=10',
        'B1': '=A1*2',
        'C1': '=A1+5',
        'D1': '=B1+C1'
    })
    order = graph.calculation_order
    # A1 must be before B1 and C1
    assert order.index('A1') < order.index('B1')
    assert order.index('A1') < order.index('C1')
    # B1 and C1 must be before D1
    assert order.index('B1') < order.index('D1')
    assert order.index('C1') < order.index('D1')
```

## 2. Integration Tests

### Multi-Component Tests

```python
# test_integration.py

def test_simple_workbook():
    """Test complete calculation flow"""
    workbook = create_workbook({
        'Sheet1': {
            'A1': 10,
            'A2': 20,
            'A3': '=A1+A2',
            'B1': '=SUM(A1:A3)'
        }
    })
    
    engine = FormulaEngine()
    engine.calculate(workbook)
    
    assert workbook.get_value('Sheet1!A3') == 30
    assert workbook.get_value('Sheet1!B1') == 60  # 10+20+30

def test_cross_sheet_references():
    """Test references across sheets"""
    workbook = create_workbook({
        'Data': {
            'A1': 100
        },
        'Summary': {
            'A1': '=Data!A1*2'
        }
    })
    
    engine = FormulaEngine()
    engine.calculate(workbook)
    
    assert workbook.get_value('Summary!A1') == 200

def test_named_ranges():
    """Test named range support"""
    workbook = create_workbook({
        'Sheet1': {
            'A1': 10,
            'A2': 20,
            'A3': 30,
            'B1': '=SUM(MyRange)'
        }
    })
    workbook.define_name('MyRange', 'Sheet1!A1:A3')
    
    engine = FormulaEngine()
    engine.calculate(workbook)
    
    assert workbook.get_value('Sheet1!B1') == 60
```

## 3. Excel Compatibility Tests

### Test Against Real Excel Files

Create Excel files with known formulas and results, then validate:

```python
# test_excel_compatibility.py

def test_excel_file_basic():
    """Load Excel file and verify calculations match"""
    # Create test file in Excel with formulas
    workbook = load_excel('test-cases/basic_formulas.xlsx')
    
    # Store expected values from Excel
    expected = {
        'Sheet1!A3': 30,
        'Sheet1!B1': 60,
        'Sheet1!C1': 2.5
    }
    
    # Calculate with our engine
    engine = FormulaEngine()
    engine.calculate(workbook)
    
    # Verify results match
    for cell, expected_value in expected.items():
        actual = workbook.get_value(cell)
        assert_close(actual, expected_value, tolerance=1e-10)

def test_excel_file_complex():
    """Test complex Excel file with 100+ formulas"""
    workbook = load_excel('test-cases/complex_workbook.xlsx')
    
    # Load expected results (exported from Excel)
    expected = load_json('test-cases/complex_workbook_expected.json')
    
    engine = FormulaEngine()
    engine.calculate(workbook)
    
    mismatches = []
    for cell, expected_value in expected.items():
        actual = workbook.get_value(cell)
        if not values_match(actual, expected_value):
            mismatches.append({
                'cell': cell,
                'expected': expected_value,
                'actual': actual
            })
    
    assert len(mismatches) == 0, f"Mismatches: {mismatches}"
```

### Creating Test Excel Files

**Process**:
1. Create Excel file with formulas
2. Let Excel calculate results
3. Export both formulas and results
4. Use as test case

**Example VBA Script to Export**:
```vba
Sub ExportFormulasAndValues()
    Dim ws As Worksheet
    Dim cell As Range
    Dim output As String
    
    For Each ws In ActiveWorkbook.Worksheets
        For Each cell In ws.UsedRange
            If cell.HasFormula Then
                output = output & ws.Name & "!" & cell.Address & vbTab
                output = output & cell.Formula & vbTab
                output = output & cell.Value & vbCrLf
            End If
        Next cell
    Next ws
    
    ' Write to file
    Open "formulas_export.txt" For Output As #1
    Print #1, output
    Close #1
End Sub
```

## 4. Edge Case Tests

### Precision Tests

```python
def test_floating_point_precision():
    """Test Excel's 15-digit precision"""
    # Excel stores 15 significant digits
    result = evaluate("=1.23456789012345")
    assert result == 1.23456789012345
    
    # Beyond 15 digits gets rounded
    result = evaluate("=1.234567890123456789")
    assert result == 1.23456789012346  # Rounded

def test_comparison_tolerance():
    """Excel has tolerance in comparisons"""
    # These should be equal in Excel despite floating point
    result = evaluate("=(0.1+0.2)=0.3")
    assert result == True
```

### Type Coercion Tests

```python
def test_text_to_number_coercion():
    """Test automatic type conversion"""
    assert evaluate('="5"+3') == 8  # Text "5" -> number 5
    assert evaluate('=5&3') == "53"  # Numbers -> text

def test_boolean_arithmetic():
    """Booleans convert to 1/0 in arithmetic"""
    assert evaluate('=TRUE+1') == 2
    assert evaluate('=FALSE*10') == 0
```

### Error Handling Tests

```python
def test_error_types():
    """Test all Excel error types"""
    assert is_error(evaluate('=1/0'), '#DIV/0!')
    assert is_error(evaluate('=VLOOKUP("X",A1:B10,2,0)'), '#N/A')
    assert is_error(evaluate('=UNKNOWNFUNC()'), '#NAME?')
    assert is_error(evaluate('=A1:B2 A3:B4'), '#NULL!')
    assert is_error(evaluate('=SQRT(-1)'), '#NUM!')
    assert is_error(evaluate('=Sheet99!A1'), '#REF!')
    assert is_error(evaluate('="text"+1'), '#VALUE!')

def test_error_propagation():
    """Errors propagate through formulas"""
    workbook = create_workbook({
        'Sheet1': {
            'A1': '=1/0',
            'A2': '=A1+5',
            'A3': '=SUM(A1:A2)'
        }
    })
    
    engine = FormulaEngine()
    engine.calculate(workbook)
    
    assert is_error(workbook.get_value('Sheet1!A1'), '#DIV/0!')
    assert is_error(workbook.get_value('Sheet1!A2'), '#DIV/0!')
    assert is_error(workbook.get_value('Sheet1!A3'), '#DIV/0!')
```

## 5. Performance Tests

```python
# test_performance.py

def test_large_workbook_performance():
    """Test performance on large workbook"""
    import time
    
    # Create workbook with 10,000 cells
    workbook = create_large_workbook(rows=100, cols=100)
    
    engine = FormulaEngine()
    
    start = time.time()
    engine.calculate(workbook)
    duration = time.time() - start
    
    # Should complete in under 1 second
    assert duration < 1.0, f"Too slow: {duration}s"

def test_incremental_calculation():
    """Test incremental recalculation"""
    workbook = create_workbook({
        'Sheet1': {
            'A1': 10,
            'B1': '=A1*2',
            'C1': '=B1+5',
            'D1': 100  # Independent cell
        }
    })
    
    engine = FormulaEngine()
    engine.calculate(workbook)
    
    # Change A1
    workbook.set_value('Sheet1!A1', 20)
    
    # Track which cells were recalculated
    recalculated = engine.recalculate_changed(workbook, ['Sheet1!A1'])
    
    # Should only recalculate A1, B1, C1 (not D1)
    assert set(recalculated) == {'Sheet1!A1', 'Sheet1!B1', 'Sheet1!C1'}
```

## 6. Test Data Organization

### Directory Structure

```
test-cases/
├── unit/
│   ├── functions/
│   │   ├── math_functions.xlsx
│   │   ├── text_functions.xlsx
│   │   └── lookup_functions.xlsx
│   └── parser/
│       └── formula_samples.txt
├── integration/
│   ├── simple_workbook.xlsx
│   ├── cross_sheet.xlsx
│   └── named_ranges.xlsx
├── compatibility/
│   ├── excel_2016_formulas.xlsx
│   ├── excel_365_formulas.xlsx
│   └── expected_results.json
└── edge_cases/
    ├── precision.xlsx
    ├── errors.xlsx
    └── circular_refs.xlsx
```

### Test Case Format

**JSON format for expected results**:
```json
{
  "test_name": "Basic Arithmetic",
  "file": "basic_formulas.xlsx",
  "expected_values": {
    "Sheet1!A3": 30,
    "Sheet1!B1": 60,
    "Sheet1!C1": 2.5
  },
  "expected_formulas": {
    "Sheet1!A3": "=A1+A2",
    "Sheet1!B1": "=SUM(A1:A3)"
  }
}
```

## 7. Continuous Testing

### Automated Test Suite

```bash
# Run all tests
pytest tests/ -v

# Run specific category
pytest tests/unit/ -v
pytest tests/integration/ -v
pytest tests/compatibility/ -v

# Run with coverage
pytest tests/ --cov=excel_engine --cov-report=html

# Run performance tests
pytest tests/performance/ -v --benchmark
```

### CI/CD Integration

```yaml
# .github/workflows/test.yml
name: Test Suite

on: [push, pull_request]

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - name: Set up Python
        uses: actions/setup-python@v2
        with:
          python-version: '3.11'
      - name: Install dependencies
        run: |
          pip install -r requirements.txt
          pip install pytest pytest-cov
      - name: Run tests
        run: pytest tests/ -v --cov=excel_engine
      - name: Upload coverage
        uses: codecov/codecov-action@v2
```

## 8. Validation Checklist

Before considering implementation complete:

- [ ] All unit tests pass (100+ tests)
- [ ] Integration tests pass (20+ scenarios)
- [ ] Excel compatibility tests pass (95%+ match rate)
- [ ] Edge cases handled correctly
- [ ] Performance targets met
- [ ] Code coverage > 80%
- [ ] No memory leaks in long-running tests
- [ ] Documentation complete
- [ ] Examples work correctly

## 9. Test-Driven Development Approach

### Red-Green-Refactor Cycle

1. **Red**: Write failing test
```python
def test_sum_function():
    assert SUM(1, 2, 3) == 6  # FAILS - not implemented yet
```

2. **Green**: Implement minimal code to pass
```python
def SUM(*args):
    return sum(args)  # Simple implementation
```

3. **Refactor**: Improve implementation
```python
def SUM(*args):
    total = 0
    for arg in args:
        if is_error(arg):
            return arg
        if is_range(arg):
            for cell in flatten_range(arg):
                if is_number(cell):
                    total += cell
        elif is_number(arg):
            total += arg
    return total
```

## 10. Debugging Failed Tests

### Comparison Tool

```python
def compare_with_excel(formula: str, context: dict):
    """
    Helper to compare our result with Excel
    
    Usage:
        compare_with_excel("=SUM(A1:A3)", {"A1": 1, "A2": 2, "A3": 3})
    """
    # Calculate with our engine
    our_result = evaluate(formula, context)
    
    # Create Excel file and calculate
    excel_result = calculate_in_excel(formula, context)
    
    print(f"Formula: {formula}")
    print(f"Our result: {our_result}")
    print(f"Excel result: {excel_result}")
    print(f"Match: {values_match(our_result, excel_result)}")
```

## Success Metrics

- **Correctness**: 95%+ match with Excel on test suite
- **Coverage**: 100+ Excel functions implemented
- **Performance**: 10K cells in < 1 second
- **Reliability**: All edge cases handled gracefully
