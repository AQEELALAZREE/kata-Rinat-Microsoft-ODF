# Technology Stack Recommendations

## Language Selection Criteria

When choosing a language for implementing the Excel Formula Engine, consider:

1. **Excel File Reading Libraries**: Mature libraries for reading `.xlsx` files
2. **Performance**: Ability to handle large spreadsheets efficiently
3. **Type System**: Strong typing helps prevent errors in formula evaluation
4. **Ecosystem**: Available parsing libraries and testing frameworks
5. **Team Expertise**: Familiarity with the language

## Recommended Options

### Option 1: Python ⭐ (Recommended for Rapid Development)

**Pros**:
- Excellent Excel libraries (`openpyxl`, `xlrd`, `pandas`)
- Rich ecosystem for parsing (PLY, Lark, pyparsing)
- Easy to prototype and test
- Great for data processing
- Extensive testing frameworks (pytest)

**Cons**:
- Slower than compiled languages
- GIL limits parallel processing
- Dynamic typing can hide bugs

**Key Libraries**:
```python
openpyxl      # Read/write Excel 2010+ files
xlrd          # Read older Excel files
lark          # Parser generator
pytest        # Testing framework
numpy         # Numerical operations
```

**Recommended Structure**:
```
excel_engine/
├── parser/
│   ├── lexer.py
│   ├── parser.py
│   └── ast_nodes.py
├── engine/
│   ├── calculator.py
│   ├── dependency_graph.py
│   └── context.py
├── functions/
│   ├── math_functions.py
│   ├── text_functions.py
│   ├── logical_functions.py
│   └── lookup_functions.py
├── reader/
│   └── excel_reader.py
└── tests/
    ├── test_parser.py
    ├── test_functions.py
    └── test_integration.py
```

---

### Option 2: TypeScript/JavaScript ⭐ (Recommended for Web Integration)

**Pros**:
- Excellent for web-based applications
- Great Excel library (SheetJS/xlsx)
- Existing formula parsers available
- Can run in browser or Node.js
- TypeScript provides type safety

**Cons**:
- Floating-point precision quirks
- Async complexity for large files
- Less mature parsing libraries than Python

**Key Libraries**:
```json
{
  "xlsx": "^0.18.5",           // Excel file reading
  "chevrotain": "^11.0.0",     // Parser building
  "mathjs": "^12.0.0",         // Mathematical operations
  "jest": "^29.0.0",           // Testing
  "typescript": "^5.0.0"       // Type safety
}
```

**Recommended Structure**:
```
src/
├── parser/
│   ├── lexer.ts
│   ├── parser.ts
│   └── ast.ts
├── engine/
│   ├── calculator.ts
│   ├── dependency-graph.ts
│   └── context.ts
├── functions/
│   ├── math.ts
│   ├── text.ts
│   ├── logical.ts
│   └── lookup.ts
├── reader/
│   └── excel-reader.ts
└── __tests__/
    ├── parser.test.ts
    ├── functions.test.ts
    └── integration.test.ts
```

---

### Option 3: Java (Enterprise-Grade)

**Pros**:
- Apache POI is industry-standard for Excel
- Excellent performance
- Strong type system
- Great for large-scale systems
- Mature testing frameworks

**Cons**:
- More verbose code
- Slower development cycle
- Heavier runtime

**Key Libraries**:
```xml
<dependencies>
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
    </dependency>
    <dependency>
        <groupId>org.antlr</groupId>
        <artifactId>antlr4</artifactId>
    </dependency>
    <dependency>
        <groupId>junit</groupId>
        <artifactId>junit</artifactId>
    </dependency>
</dependencies>
```

---

### Option 4: C# (Windows-Focused)

**Pros**:
- EPPlus and ClosedXML are excellent
- LINQ for data processing
- Strong type system
- Great IDE support (Visual Studio)
- Good performance

**Cons**:
- Primarily Windows-focused
- Smaller open-source ecosystem than Java

**Key Libraries**:
```
EPPlus           # Excel file handling
Sprache          # Parser combinator library
xUnit            # Testing framework
```

---

### Option 5: Rust (Maximum Performance)

**Pros**:
- Excellent performance
- Memory safety
- Growing Excel ecosystem
- Can compile to WebAssembly

**Cons**:
- Steeper learning curve
- Smaller ecosystem
- Longer development time

**Key Libraries**:
```toml
[dependencies]
calamine = "0.24"        # Excel reading
pest = "2.7"             # Parser generator
```

---

## Recommended Choice: Python

For this kata/learning project, **Python is recommended** because:

1. **Fast prototyping**: Get results quickly
2. **Excellent libraries**: `openpyxl` for Excel, `lark` for parsing
3. **Easy testing**: pytest makes validation straightforward
4. **Readable code**: Clear syntax for complex logic
5. **Great for learning**: Easy to understand and modify

## Parser Library Recommendation

### For Python: Lark

**Why Lark?**
- Declarative grammar definition
- Automatic AST construction
- Good error messages
- Fast enough for formulas

**Example Grammar**:
```lark
?start: expression

?expression: term
           | expression "+" term   -> add
           | expression "-" term   -> subtract

?term: factor
     | term "*" factor   -> multiply
     | term "/" factor   -> divide

?factor: NUMBER          -> number
       | cell_ref        -> cell
       | function_call
       | "(" expression ")"

function_call: FUNCTION "(" [arguments] ")"
arguments: expression ("," expression)*

cell_ref: COLUMN ROW
COLUMN: /[A-Z]+/
ROW: /[0-9]+/
FUNCTION: /[A-Z][A-Z0-9]*/

%import common.NUMBER
%import common.WS
%ignore WS
```

### For TypeScript: Chevrotain

**Why Chevrotain?**
- Pure TypeScript
- Excellent error recovery
- Fast performance
- Good documentation

---

## Development Tools

### Version Control
```bash
git init
# Add .gitignore for your language
```

### Testing Framework
- **Python**: pytest with pytest-cov for coverage
- **TypeScript**: Jest with ts-jest
- **Java**: JUnit 5
- **C#**: xUnit

### Code Quality
- **Python**: pylint, black, mypy
- **TypeScript**: ESLint, Prettier, TypeScript strict mode
- **Java**: Checkstyle, SpotBugs
- **C#**: StyleCop, Roslyn analyzers

### Documentation
- **Python**: Sphinx
- **TypeScript**: TypeDoc
- **Java**: JavaDoc
- **C#**: DocFX

---

## Getting Started Template (Python)

### 1. Project Setup
```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install openpyxl lark-parser pytest pytest-cov
```

### 2. Create Basic Structure
```bash
mkdir -p excel_engine/{parser,engine,functions,reader,tests}
touch excel_engine/__init__.py
touch excel_engine/parser/__init__.py
touch excel_engine/engine/__init__.py
touch excel_engine/functions/__init__.py
touch excel_engine/reader/__init__.py
```

### 3. First Test
```python
# tests/test_basic.py
def test_simple_addition():
    # This will fail initially - implement to make it pass
    from excel_engine.engine import evaluate
    result = evaluate("=1+1")
    assert result == 2
```

### 4. Run Tests
```bash
pytest tests/ -v
```

---

## Performance Targets

Based on the technology choice:

| Language   | 1K cells | 10K cells | 100K cells |
|------------|----------|-----------|------------|
| Python     | <50ms    | <500ms    | <5s        |
| TypeScript | <30ms    | <300ms    | <3s        |
| Java       | <20ms    | <200ms    | <2s        |
| C#         | <20ms    | <200ms    | <2s        |
| Rust       | <10ms    | <100ms    | <1s        |

These are rough estimates for typical formula complexity.

---

## Next Steps

1. Choose your language based on project requirements
2. Set up development environment
3. Install recommended libraries
4. Create project structure
5. Start with parser implementation (see `specs/PARSER_SPEC.md`)
