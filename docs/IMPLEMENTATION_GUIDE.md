# Implementation Guide

## Step-by-Step Development Plan

This guide walks through implementing the Excel Formula Engine from scratch.

## Prerequisites

- Choose your language (recommended: Python)
- Set up development environment
- Install required libraries
- Read ARCHITECTURE.md and TECHNOLOGY_STACK.md

## Phase 1: Project Setup (Day 1)

### 1.1 Create Project Structure

```bash
mkdir excel-formula-engine
cd excel-formula-engine

# For Python
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Create directory structure
mkdir -p excel_engine/{parser,engine,functions,reader,tests}
touch excel_engine/__init__.py
touch excel_engine/parser/__init__.py
touch excel_engine/engine/__init__.py
touch excel_engine/functions/__init__.py
touch excel_engine/reader/__init__.py
```

### 1.2 Install Dependencies

```bash
# requirements.txt
openpyxl>=3.1.0
lark-parser>=1.1.0
pytest>=7.4.0
pytest-cov>=4.1.0
```

```bash
pip install -r requirements.txt
```

### 1.3 Create First Test

```python
# tests/test_basic.py
def test_placeholder():
    """Placeholder test to verify setup"""
    assert True
```

Run: `pytest tests/ -v`

## Phase 2: Excel File Reader (Days 2-3)

### 2.1 Implement Basic Reader

```python
# excel_engine/reader/excel_reader.py
from openpyxl import load_workbook

class ExcelReader:
    """Read Excel files and extract cell data"""
    
    def __init__(self, filepath):
        self.workbook = load_workbook(filepath, data_only=False)
    
    def get_sheets(self):
        """Return list of sheet names"""
        return self.workbook.sheetnames
    
    def get_cell_value(self, sheet_name, row, col):
        """Get cell value"""
        sheet = self.workbook[sheet_name]
        cell = sheet.cell(row, col)
        return cell.value
    
    def get_cell_formula(self, sheet_name, row, col):
        """Get cell formula (if any)"""
        sheet = self.workbook[sheet_name]
        cell = sheet.cell(row, col)
        if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
            return cell.value
        return None
    
    def get_all_formulas(self):
        """Extract all formulas from workbook"""
        formulas = {}
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                        address = f"{sheet_name}!{cell.coordinate}"
                        formulas[address] = cell.value
        return formulas
```

### 2.2 Test Reader

```python
# tests/test_reader.py
from excel_engine.reader import ExcelReader

def test_read_simple_file():
    """Test reading a simple Excel file"""
    # Create test file first (manually or in test setup)
    reader = ExcelReader('test-cases/simple.xlsx')
    
    sheets = reader.get_sheets()
    assert 'Sheet1' in sheets
    
    formulas = reader.get_all_formulas()
    assert 'Sheet1!A3' in formulas
    assert formulas['Sheet1!A3'] == '=A1+A2'
```

## Phase 3: Formula Parser (Days 4-7)

### 3.1 Define Grammar

```python
# excel_engine/parser/grammar.lark
?start: formula

formula: "=" expression

?expression: comparison

?comparison: concat (COMP_OP concat)*

?concat: add_expr ("&" add_expr)*

?add_expr: mult_expr (("+"|"-") mult_expr)*

?mult_expr: power_expr (("*"|"/") power_expr)*

?power_expr: unary_expr ("^" unary_expr)*

?unary_expr: ("+"|"-")? atom

?atom: NUMBER           -> number
     | STRING           -> string
     | "TRUE"           -> true
     | "FALSE"          -> false
     | cell_ref         -> cell
     | range_ref        -> range
     | function_call    -> function
     | "(" expression ")"

function_call: FUNCTION "(" [arguments] ")"

arguments: expression ("," expression)*

cell_ref: COLUMN ROW

range_ref: cell_ref ":" cell_ref

COLUMN: /[A-Z]+/
ROW: /[0-9]+/
FUNCTION: /[A-Z][A-Z0-9_]*/
COMP_OP: "=" | "<>" | "<" | ">" | "<=" | ">="

STRING: /"([^"]|"")*"/

%import common.NUMBER
%import common.WS
%ignore WS
```

### 3.2 Create AST Nodes

```python
# excel_engine/parser/ast_nodes.py

class ASTNode:
    """Base class for AST nodes"""
    def evaluate(self, context):
        raise NotImplementedError

class NumberNode(ASTNode):
    def __init__(self, value):
        self.value = float(value)
    
    def evaluate(self, context):
        return self.value

class BinaryOpNode(ASTNode):
    def __init__(self, operator, left, right):
        self.operator = operator
        self.left = left
        self.right = right
    
    def evaluate(self, context):
        left_val = self.left.evaluate(context)
        right_val = self.right.evaluate(context)
        
        if self.operator == '+':
            return left_val + right_val
        elif self.operator == '-':
            return left_val - right_val
        elif self.operator == '*':
            return left_val * right_val
        elif self.operator == '/':
            if right_val == 0:
                return ExcelError("#DIV/0!")
            return left_val / right_val
        # ... more operators

class CellReferenceNode(ASTNode):
    def __init__(self, column, row):
        self.column = column
        self.row = int(row)
    
    def evaluate(self, context):
        return context.get_cell_value(self.column, self.row)

class FunctionCallNode(ASTNode):
    def __init__(self, function_name, arguments):
        self.function_name = function_name.upper()
        self.arguments = arguments
    
    def evaluate(self, context):
        func = context.get_function(self.function_name)
        args = [arg.evaluate(context) for arg in self.arguments]
        return func(*args)
```

### 3.3 Create Transformer

```python
# excel_engine/parser/parser.py
from lark import Lark, Transformer
from .ast_nodes import *

class FormulaTransformer(Transformer):
    """Transform parse tree to AST"""
    
    def number(self, items):
        return NumberNode(items[0])
    
    def add_expr(self, items):
        if len(items) == 1:
            return items[0]
        # Build left-associative tree
        result = items[0]
        for i in range(1, len(items), 2):
            op = items[i]
            right = items[i+1]
            result = BinaryOpNode(op, result, right)
        return result
    
    def cell(self, items):
        col = items[0].value
        row = items[1].value
        return CellReferenceNode(col, row)
    
    def function(self, items):
        func_name = items[0].value
        args = items[1:] if len(items) > 1 else []
        return FunctionCallNode(func_name, args)
    
    # ... more transformations

class FormulaParser:
    """Main parser class"""
    
    def __init__(self):
        with open('excel_engine/parser/grammar.lark') as f:
            grammar = f.read()
        self.parser = Lark(grammar, start='formula', parser='lalr')
        self.transformer = FormulaTransformer()
    
    def parse(self, formula):
        """Parse formula string to AST"""
        tree = self.parser.parse(formula)
        ast = self.transformer.transform(tree)
        return ast
```

### 3.4 Test Parser

```python
# tests/test_parser.py
from excel_engine.parser import FormulaParser
from excel_engine.parser.ast_nodes import *

def test_parse_number():
    parser = FormulaParser()
    ast = parser.parse("=42")
    assert isinstance(ast, NumberNode)
    assert ast.value == 42

def test_parse_addition():
    parser = FormulaParser()
    ast = parser.parse("=1+2")
    assert isinstance(ast, BinaryOpNode)
    assert ast.operator == '+'

def test_parse_precedence():
    parser = FormulaParser()
    ast = parser.parse("=1+2*3")
    # Should be: 1 + (2 * 3)
    assert ast.operator == '+'
    assert ast.right.operator == '*'
```

## Phase 4: Calculation Engine (Days 8-10)

### 4.1 Create Calculation Context

```python
# excel_engine/engine/context.py

class CalculationContext:
    """Context for formula evaluation"""
    
    def __init__(self, workbook):
        self.workbook = workbook
        self.value_cache = {}
        self.functions = {}
        self._register_functions()
    
    def get_cell_value(self, column, row, sheet=None):
        """Get value of a cell"""
        address = f"{sheet or 'Sheet1'}!{column}{row}"
        
        if address in self.value_cache:
            return self.value_cache[address]
        
        # Get from workbook
        value = self.workbook.get_value(address)
        return value
    
    def set_cell_value(self, address, value):
        """Set calculated value"""
        self.value_cache[address] = value
    
    def get_function(self, name):
        """Get function implementation"""
        if name not in self.functions:
            raise ValueError(f"Unknown function: {name}")
        return self.functions[name]
    
    def _register_functions(self):
        """Register all Excel functions"""
        from excel_engine.functions import *
        self.functions['SUM'] = SUM
        self.functions['AVERAGE'] = AVERAGE
        # ... register more functions
```

### 4.2 Create Calculator

```python
# excel_engine/engine/calculator.py

class FormulaCalculator:
    """Main calculation engine"""
    
    def __init__(self, parser):
        self.parser = parser
    
    def calculate_workbook(self, workbook):
        """Calculate all formulas in workbook"""
        # 1. Extract all formulas
        formulas = workbook.get_all_formulas()
        
        # 2. Parse formulas
        asts = {}
        for address, formula in formulas.items():
            asts[address] = self.parser.parse(formula)
        
        # 3. Build dependency graph
        graph = self._build_dependency_graph(asts)
        
        # 4. Calculate in order
        context = CalculationContext(workbook)
        
        for address in graph.calculation_order:
            ast = asts[address]
            value = ast.evaluate(context)
            context.set_cell_value(address, value)
            workbook.set_value(address, value)
        
        return workbook
    
    def _build_dependency_graph(self, asts):
        """Build dependency graph from ASTs"""
        # Simplified version - see DEPENDENCY_GRAPH_SPEC.md
        from excel_engine.engine.dependency_graph import DependencyGraph
        graph = DependencyGraph()
        
        for address, ast in asts.items():
            deps = self._extract_dependencies(ast)
            graph.add_node(address, deps)
        
        graph.compute_order()
        return graph
```

## Phase 5: Implement Functions (Days 11-20)

### 5.1 Start with Basic Functions

```python
# excel_engine/functions/math_functions.py

def SUM(*args):
    """SUM function implementation"""
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

def AVERAGE(*args):
    """AVERAGE function implementation"""
    total = 0
    count = 0
    for arg in args:
        if is_error(arg):
            return arg
        if is_range(arg):
            for cell in flatten_range(arg):
                if is_number(cell):
                    total += cell
                    count += 1
        elif is_number(arg):
            total += arg
            count += 1
    
    if count == 0:
        return ExcelError("#DIV/0!")
    
    return total / count
```

### 5.2 Test Each Function

```python
# tests/test_functions.py

def test_sum():
    assert SUM(1, 2, 3) == 6

def test_average():
    assert AVERAGE(1, 2, 3) == 2
```

### 5.3 Implement Priority Functions

See FUNCTION_LIBRARY_SPEC.md for implementation order:
1. Week 1: SUM, AVERAGE, COUNT, MIN, MAX, IF, AND, OR
2. Week 2: VLOOKUP, INDEX, MATCH, text functions
3. Week 3: Date functions, SUMIF, COUNTIF
4. Week 4: Advanced functions

## Phase 6: Integration & Testing (Days 21-25)

### 6.1 Create Test Excel Files

Create Excel files with formulas:
- basic_formulas.xlsx
- complex_formulas.xlsx
- edge_cases.xlsx

### 6.2 Run Compatibility Tests

```python
# tests/test_compatibility.py

def test_basic_formulas_file():
    from excel_engine import FormulaEngine
    
    engine = FormulaEngine()
    workbook = engine.load('test-cases/basic_formulas.xlsx')
    engine.calculate(workbook)
    
    # Verify against expected results
    assert workbook.get_value('Sheet1!A3') == 30
```

### 6.3 Fix Issues

Iterate on failing tests, comparing with Excel's behavior.

## Phase 7: Optimization (Days 26-30)

### 7.1 Add Caching

```python
class FormulaCalculator:
    def __init__(self, parser):
        self.parser = parser
        self.ast_cache = {}  # Cache parsed formulas
    
    def parse_formula(self, formula):
        if formula not in self.ast_cache:
            self.ast_cache[formula] = self.parser.parse(formula)
        return self.ast_cache[formula]
```

### 7.2 Implement Incremental Calculation

Only recalculate changed cells and dependents.

### 7.3 Profile Performance

```python
import cProfile

profiler = cProfile.Profile()
profiler.enable()

engine.calculate(large_workbook)

profiler.disable()
profiler.print_stats(sort='cumulative')
```

## Common Pitfalls

### 1. Operator Precedence
Make sure parser respects Excel's precedence rules.

### 2. Type Coercion
Excel has complex type conversion rules - implement carefully.

### 3. Error Handling
Errors must propagate correctly through formulas.

### 4. Circular References
Detect and handle appropriately (iterative calculation or error).

### 5. Precision
Use proper floating-point comparison with tolerance.

## Debugging Tips

### 1. Compare with Excel
For any failing test, create the same formula in Excel and compare.

### 2. Print AST
Visualize the parsed AST to verify structure.

### 3. Step Through Evaluation
Add logging to see evaluation order.

### 4. Use Small Test Cases
Isolate issues with minimal examples.

## Success Criteria

- [ ] Can read Excel files
- [ ] Can parse formulas correctly
- [ ] Can build dependency graph
- [ ] Can calculate formulas in order
- [ ] Implements 50+ functions
- [ ] Passes compatibility tests (95%+ match)
- [ ] Handles errors correctly
- [ ] Performance acceptable (<1s for 10K cells)

## Next Steps After Implementation

1. Add more functions (target 100+)
2. Support array formulas
3. Support structured references (Excel tables)
4. Add formula auditing tools
5. Create web interface
6. Optimize for very large workbooks

## Resources

- See specs/ directory for detailed specifications
- See test-cases/ for example Excel files
- See review-checklists/ for code review guidelines
