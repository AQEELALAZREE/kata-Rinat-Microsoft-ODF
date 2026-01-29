# Formula Parser Specification

## Overview

The Formula Parser converts Excel formula strings into Abstract Syntax Trees (AST) that can be evaluated by the calculation engine.

## Input/Output

**Input**: Formula string (e.g., `"=SUM(A1:A10) + B1 * 2"`)  
**Output**: AST representing the formula structure

## Grammar Specification

### Complete EBNF Grammar

```ebnf
formula          ::= '=' expression
expression       ::= comparison_expr
comparison_expr  ::= concat_expr (comparison_op concat_expr)*
concat_expr      ::= add_expr ('&' add_expr)*
add_expr         ::= mult_expr (('+' | '-') mult_expr)*
mult_expr        ::= power_expr (('*' | '/') power_expr)*
power_expr       ::= percent_expr ('^' percent_expr)*
percent_expr     ::= unary_expr '%'?
unary_expr       ::= ('+' | '-')? postfix_expr
postfix_expr     ::= primary_expr (':' primary_expr)?

primary_expr     ::= number
                   | string
                   | boolean
                   | error
                   | cell_reference
                   | range_reference
                   | function_call
                   | named_range
                   | '(' expression ')'
                   | array_constant

function_call    ::= function_name '(' argument_list? ')'
argument_list    ::= expression (',' expression)* | expression (';' expression)*
array_constant   ::= '{' array_rows '}'
array_rows       ::= array_row (';' array_row)*
array_row        ::= expression (',' expression)*

cell_reference   ::= sheet_ref? column row
range_reference  ::= cell_reference ':' cell_reference
sheet_ref        ::= (sheet_name | quoted_sheet) '!'
column           ::= [A-Z]+
row              ::= [0-9]+
function_name    ::= [A-Z][A-Z0-9_.]*

comparison_op    ::= '=' | '<>' | '<' | '>' | '<=' | '>='
number           ::= [0-9]+ ('.' [0-9]+)? ([eE][+-]?[0-9]+)?
string           ::= '"' ([^"] | '""')* '"'
boolean          ::= 'TRUE' | 'FALSE'
error            ::= '#DIV/0!' | '#N/A' | '#NAME?' | '#NULL!' | '#NUM!' | '#REF!' | '#VALUE!'
```

## AST Node Types

### Base Node
```python
class ASTNode:
    """Base class for all AST nodes"""
    def evaluate(self, context: CalculationContext) -> Value:
        raise NotImplementedError
```

### Literal Nodes

#### NumberNode
```python
class NumberNode(ASTNode):
    def __init__(self, value: float):
        self.value = value
    
    def evaluate(self, context):
        return self.value
```

#### StringNode
```python
class StringNode(ASTNode):
    def __init__(self, value: str):
        self.value = value
    
    def evaluate(self, context):
        return self.value
```

#### BooleanNode
```python
class BooleanNode(ASTNode):
    def __init__(self, value: bool):
        self.value = value
    
    def evaluate(self, context):
        return self.value
```

#### ErrorNode
```python
class ErrorNode(ASTNode):
    def __init__(self, error_type: str):
        self.error_type = error_type  # e.g., "#DIV/0!"
    
    def evaluate(self, context):
        return ExcelError(self.error_type)
```

### Reference Nodes

#### CellReferenceNode
```python
class CellReferenceNode(ASTNode):
    def __init__(self, sheet: str | None, column: str, row: int, 
                 absolute_col: bool = False, absolute_row: bool = False):
        self.sheet = sheet
        self.column = column
        self.row = row
        self.absolute_col = absolute_col  # $A vs A
        self.absolute_row = absolute_row  # $1 vs 1
    
    def evaluate(self, context):
        address = self.resolve_address(context)
        return context.get_cell_value(address)
```

#### RangeReferenceNode
```python
class RangeReferenceNode(ASTNode):
    def __init__(self, start: CellReferenceNode, end: CellReferenceNode):
        self.start = start
        self.end = end
    
    def evaluate(self, context):
        # Returns a 2D array of values
        return context.get_range_values(self.start, self.end)
```

### Operator Nodes

#### BinaryOpNode
```python
class BinaryOpNode(ASTNode):
    def __init__(self, operator: str, left: ASTNode, right: ASTNode):
        self.operator = operator  # '+', '-', '*', '/', '^', '&', '=', '<>', etc.
        self.left = left
        self.right = right
    
    def evaluate(self, context):
        left_val = self.left.evaluate(context)
        right_val = self.right.evaluate(context)
        return self.apply_operator(self.operator, left_val, right_val)
```

#### UnaryOpNode
```python
class UnaryOpNode(ASTNode):
    def __init__(self, operator: str, operand: ASTNode):
        self.operator = operator  # '+', '-', '%'
        self.operand = operand
    
    def evaluate(self, context):
        value = self.operand.evaluate(context)
        return self.apply_operator(self.operator, value)
```

### Function Node

#### FunctionCallNode
```python
class FunctionCallNode(ASTNode):
    def __init__(self, function_name: str, arguments: list[ASTNode]):
        self.function_name = function_name.upper()
        self.arguments = arguments
    
    def evaluate(self, context):
        # Get function implementation from registry
        func = context.get_function(self.function_name)
        
        # Evaluate arguments (unless function requires lazy evaluation)
        if func.is_lazy:
            args = self.arguments  # Pass AST nodes
        else:
            args = [arg.evaluate(context) for arg in self.arguments]
        
        return func.execute(args, context)
```

## Parsing Examples

### Example 1: Simple Arithmetic
**Input**: `=1+2*3`

**AST**:
```
BinaryOpNode('+')
├── NumberNode(1)
└── BinaryOpNode('*')
    ├── NumberNode(2)
    └── NumberNode(3)
```

### Example 2: Cell References
**Input**: `=A1+B2`

**AST**:
```
BinaryOpNode('+')
├── CellReferenceNode(sheet=None, col='A', row=1)
└── CellReferenceNode(sheet=None, col='B', row=2)
```

### Example 3: Function Call
**Input**: `=SUM(A1:A10)`

**AST**:
```
FunctionCallNode('SUM')
└── RangeReferenceNode
    ├── CellReferenceNode(col='A', row=1)
    └── CellReferenceNode(col='A', row=10)
```

### Example 4: Nested Functions
**Input**: `=IF(A1>10, SUM(B1:B5), AVERAGE(C1:C5))`

**AST**:
```
FunctionCallNode('IF')
├── BinaryOpNode('>')
│   ├── CellReferenceNode(col='A', row=1)
│   └── NumberNode(10)
├── FunctionCallNode('SUM')
│   └── RangeReferenceNode(B1:B5)
└── FunctionCallNode('AVERAGE')
    └── RangeReferenceNode(C1:C5)
```

### Example 5: Complex Formula
**Input**: `=VLOOKUP(A2, Sheet2!$A$1:$B$100, 2, FALSE) & " - " & TEXT(TODAY(), "yyyy-mm-dd")`

**AST**:
```
BinaryOpNode('&')
├── BinaryOpNode('&')
│   ├── FunctionCallNode('VLOOKUP')
│   │   ├── CellReferenceNode(col='A', row=2)
│   │   ├── RangeReferenceNode(sheet='Sheet2', A1:B100, absolute)
│   │   ├── NumberNode(2)
│   │   └── BooleanNode(False)
│   └── StringNode(" - ")
└── FunctionCallNode('TEXT')
    ├── FunctionCallNode('TODAY')
    └── StringNode("yyyy-mm-dd")
```

## Special Parsing Cases

### 1. Absolute References
- `$A$1` - Both column and row absolute
- `$A1` - Column absolute, row relative
- `A$1` - Column relative, row absolute
- `A1` - Both relative

### 2. Sheet References
- `Sheet1!A1` - Simple sheet name
- `'Sheet Name'!A1` - Sheet name with spaces
- `'Sheet''s Data'!A1` - Sheet name with apostrophe (escaped)

### 3. Named Ranges
- `=SUM(SalesData)` - Named range instead of cell reference
- Must be resolved during evaluation

### 4. Array Constants
- `{1,2,3}` - Horizontal array
- `{1;2;3}` - Vertical array
- `{1,2;3,4}` - 2D array

### 5. Structured References (Excel Tables)
- `=[@Column1]` - Current row
- `=[Column1]` - Entire column
- `=Table1[[#Headers],[Column1]]` - Header cell

## Error Handling

### Syntax Errors
Return descriptive error messages:
- "Unexpected token ')' at position 15"
- "Missing closing parenthesis"
- "Invalid cell reference: 'ABC123456'"
- "Unknown function: 'SUMM'"

### Recovery Strategy
- Try to identify error location
- Provide helpful suggestions
- Don't crash on invalid input

## Implementation Guidelines

### 1. Tokenization (Lexer)
```python
class Token:
    def __init__(self, type: str, value: any, position: int):
        self.type = type      # 'NUMBER', 'OPERATOR', 'FUNCTION', etc.
        self.value = value    # Actual value
        self.position = position

def tokenize(formula: str) -> list[Token]:
    """Convert formula string to tokens"""
    # Implementation details...
```

### 2. Parsing (Parser)
```python
class FormulaParser:
    def __init__(self, tokens: list[Token]):
        self.tokens = tokens
        self.position = 0
    
    def parse(self) -> ASTNode:
        """Parse tokens into AST"""
        if self.current_token().value != '=':
            raise ParseError("Formula must start with '='")
        self.advance()
        return self.parse_expression()
    
    def parse_expression(self) -> ASTNode:
        """Parse expression with operator precedence"""
        # Implementation using recursive descent or Pratt parsing
```

### 3. Operator Precedence (Highest to Lowest)
1. `()` - Parentheses
2. `:` - Range operator
3. `%` - Percent
4. `^` - Exponentiation
5. `*`, `/` - Multiplication, Division
6. `+`, `-` - Addition, Subtraction
7. `&` - String concatenation
8. `=`, `<>`, `<`, `>`, `<=`, `>=` - Comparison

## Testing Strategy

### Unit Tests for Parser

```python
def test_parse_number():
    ast = parse("=42")
    assert isinstance(ast, NumberNode)
    assert ast.value == 42

def test_parse_addition():
    ast = parse("=1+2")
    assert isinstance(ast, BinaryOpNode)
    assert ast.operator == '+'
    assert ast.left.value == 1
    assert ast.right.value == 2

def test_parse_function():
    ast = parse("=SUM(A1:A10)")
    assert isinstance(ast, FunctionCallNode)
    assert ast.function_name == 'SUM'
    assert len(ast.arguments) == 1

def test_parse_nested_functions():
    ast = parse("=IF(A1>0, SUM(B:B), 0)")
    assert isinstance(ast, FunctionCallNode)
    assert ast.function_name == 'IF'
    assert len(ast.arguments) == 3
```

### Edge Cases to Test
- Empty formula: `=`
- Just a value: `=42`
- Nested parentheses: `=((1+2)*3)`
- Multiple operators: `=1+2*3-4/5^6`
- String with quotes: `="He said ""Hello"""`
- Error values: `=#DIV/0!`
- Complex ranges: `=SUM(Sheet1!$A$1:$Z$100)`

## Performance Considerations

### Optimization Strategies
1. **Cache parsed formulas**: Don't re-parse unchanged formulas
2. **Lazy parsing**: Only parse formulas when needed
3. **Parallel parsing**: Parse independent formulas in parallel
4. **Incremental parsing**: Re-parse only changed parts

### Expected Performance
- Simple formula (5 tokens): <1ms
- Medium formula (20 tokens): <5ms
- Complex formula (100 tokens): <20ms

## References

- [Excel Formula Grammar](https://github.com/spreadsheetlab/XLParser)
- [OpenFormula Specification](https://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os-part2.html)
- [Lark Parser Documentation](https://lark-parser.readthedocs.io/)
