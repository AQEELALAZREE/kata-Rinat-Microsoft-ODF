# Excel Formula Engine - Quick Start Examples

## Example 1: Basic Usage (Python)

```python
from excel_engine import FormulaEngine

# Load Excel file
engine = FormulaEngine()
workbook = engine.load('my_spreadsheet.xlsx')

# Calculate all formulas
engine.calculate(workbook)

# Get calculated values
value = workbook.get_value('Sheet1!A3')
print(f"Result: {value}")

# Save updated workbook
workbook.save('my_spreadsheet_calculated.xlsx')
```

## Example 2: Evaluating Single Formula

```python
from excel_engine.parser import FormulaParser
from excel_engine.engine import CalculationContext

# Parse a formula
parser = FormulaParser()
ast = parser.parse("=SUM(A1:A10)")

# Create context with cell values
context = CalculationContext()
context.set_range('A1:A10', [1, 2, 3, 4, 5, 6, 7, 8, 9, 10])

# Evaluate
result = ast.evaluate(context)
print(f"SUM result: {result}")  # Output: 55
```

## Example 3: Adding Custom Function

```python
from excel_engine.functions import register_function

def CUSTOM_DOUBLE(value):
    """Custom function that doubles a value"""
    return value * 2

# Register the function
register_function('CUSTOM_DOUBLE', CUSTOM_DOUBLE)

# Now you can use it in formulas
result = engine.evaluate("=CUSTOM_DOUBLE(5)")
print(result)  # Output: 10
```

## Example 4: Incremental Recalculation

```python
# Load and calculate workbook
workbook = engine.load('spreadsheet.xlsx')
engine.calculate(workbook)

# Change a cell value
workbook.set_value('Sheet1!A1', 100)

# Recalculate only affected cells
affected = engine.recalculate_changed(workbook, ['Sheet1!A1'])
print(f"Recalculated {len(affected)} cells")
```

## Example 5: Handling Errors

```python
from excel_engine.types import ExcelError, is_error

# Evaluate formula that might error
result = engine.evaluate("=1/0")

if is_error(result):
    print(f"Error: {result.type}")  # Output: Error: #DIV/0!
else:
    print(f"Result: {result}")
```

## Example 6: Working with Ranges

```python
from excel_engine.types import Range

# Create a range
data = Range([
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
])

# Use in formula
context = CalculationContext()
context.set_range('A1:C3', data)

result = engine.evaluate("=SUM(A1:C3)", context)
print(result)  # Output: 45
```

## Example 7: Testing Against Excel

```python
import openpyxl

# Create test workbook
wb = openpyxl.Workbook()
ws = wb.active
ws['A1'] = 10
ws['A2'] = 20
ws['A3'] = '=A1+A2'

# Save and let Excel calculate
wb.save('test.xlsx')

# Load with our engine
engine_wb = engine.load('test.xlsx')
engine.calculate(engine_wb)

# Compare results
excel_value = 30  # Value calculated by Excel
our_value = engine_wb.get_value('Sheet1!A3')

assert our_value == excel_value, f"Mismatch: {our_value} != {excel_value}"
print("✓ Results match!")
```

## Example 8: Performance Benchmarking

```python
import time

# Create large workbook
workbook = create_large_workbook(rows=100, cols=100)

# Benchmark calculation
start = time.time()
engine.calculate(workbook)
duration = time.time() - start

print(f"Calculated 10,000 cells in {duration:.2f}s")
print(f"Rate: {10000/duration:.0f} cells/second")
```

## Example 9: Dependency Graph Visualization

```python
from excel_engine.engine import DependencyGraph

# Build dependency graph
graph = engine.build_dependency_graph(workbook)

# Print calculation order
print("Calculation order:")
for i, cell in enumerate(graph.calculation_order):
    print(f"{i+1}. {cell}")

# Check for circular references
if graph.circular_refs:
    print("\nCircular references detected:")
    for cycle in graph.circular_refs:
        print(f"  {' → '.join(cycle)}")
```

## Example 10: Formula Parsing and AST

```python
from excel_engine.parser import FormulaParser

parser = FormulaParser()

# Parse complex formula
formula = "=IF(A1>10, SUM(B:B), AVERAGE(C:C))"
ast = parser.parse(formula)

# Print AST structure
def print_ast(node, indent=0):
    print("  " * indent + type(node).__name__)
    if hasattr(node, 'left'):
        print_ast(node.left, indent+1)
    if hasattr(node, 'right'):
        print_ast(node.right, indent+1)
    if hasattr(node, 'arguments'):
        for arg in node.arguments:
            print_ast(arg, indent+1)

print_ast(ast)
```

## Example 11: Type Conversion

```python
from excel_engine.types import to_number, to_text, to_boolean

# Excel-style type conversion
print(to_number("123"))      # 123.0
print(to_number(True))       # 1.0
print(to_number(""))         # 0.0

print(to_text(123))          # "123"
print(to_text(True))         # "TRUE"

print(to_boolean(1))         # True
print(to_boolean("TRUE"))    # True
print(to_boolean(0))         # False
```

## Example 12: Building Workbook Programmatically

```python
from excel_engine.workbook import Workbook

# Create workbook
wb = Workbook()
sheet = wb.add_sheet('Data')

# Add values
sheet.set_value('A1', 10)
sheet.set_value('A2', 20)
sheet.set_value('A3', '=A1+A2')

# Calculate
engine.calculate(wb)

# Get result
print(wb.get_value('Data!A3'))  # Output: 30
```

## Running the Examples

### Setup
```bash
# Install the engine
pip install -e .

# Run example
python examples/example_01_basic.py
```

### All Examples
```bash
# Run all examples
python examples/run_all.py
```

## Next Steps

1. Read the [Implementation Guide](../docs/IMPLEMENTATION_GUIDE.md)
2. Review the [Architecture](../docs/ARCHITECTURE.md)
3. Check the [Function Library Spec](../specs/FUNCTION_LIBRARY_SPEC.md)
4. Run the test suite: `pytest tests/`
