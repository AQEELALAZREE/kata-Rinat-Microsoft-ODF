# Excel Spec Files

This directory contains JSON-based specification files for testing the Excel formula engine. These specs are defined according to [RFC 004: Intermediate Text Format for Excel Specs](../../specs/RFC_004_TEXT_FORMAT.md).

## Directory Structure

```
specs/
├── basic/          # Simple arithmetic and basic formulas
├── functions/      # Function-specific tests
├── errors/         # Error handling and propagation
└── edge_cases/     # Precision, type coercion, edge cases
```

## What is a Spec File?

A spec file is a JSON document that describes a small Excel workbook for testing purposes. It includes:

- **Metadata**: Locale, decimal separator, list separator
- **Sheets**: One or more worksheets
- **Cells**: Values, formulas, and expected results

## Example Spec

```json
{
  "version": 1,
  "description": "Basic arithmetic operations",
  "meta": {
    "locale": "en-US",
    "decimal_separator": ".",
    "list_separator": ","
  },
  "sheets": [
    {
      "name": "Sheet1",
      "cells": {
        "A1": { "value": 10 },
        "A2": { "value": 20 },
        "A3": { "formula": "=A1+A2", "value": 30 }
      }
    }
  ]
}
```

## Available Specs

### Basic

- **[arithmetic.json](basic/arithmetic.json)** - Basic arithmetic operations (+, -, *, /)
- **[multi_sheet.json](basic/multi_sheet.json)** - Cross-sheet references and SUMPRODUCT

### Functions

- **[sum.json](functions/sum.json)** - SUM function with ranges and multiple arguments
- **[if.json](functions/if.json)** - IF function with logical conditions

### Errors

- **[error_types.json](errors/error_types.json)** - All Excel error types (#DIV/0!, #N/A, #VALUE!, etc.)

### Edge Cases

- **[type_coercion.json](edge_cases/type_coercion.json)** - Type conversion (string to number, boolean arithmetic, null handling)

## Creating New Specs

Use the `/create-spec` workflow:

```bash
# See workflow for details
cat .agent/workflows/create-spec.md
```

Or manually create a spec file following the template in [RFC_004_TEXT_FORMAT.md](../../specs/RFC_004_TEXT_FORMAT.md).

## Using Specs in Tests

```rust
use spec_format::SpecParser;

#[test]
fn test_from_spec() {
    // Load spec file
    let spec = SpecParser::parse_file("test-cases/specs/basic/arithmetic.json").unwrap();
    
    // Convert to workbook
    let workbook = spec.to_workbook();
    
    // Run calculation engine
    let engine = FormulaEngine::new();
    engine.calculate(&mut workbook);
    
    // Validate results
    for sheet in &spec.sheets {
        for (addr, cell_spec) in &sheet.cells {
            let actual = workbook.get_value(&format!("{}!{}", sheet.name, addr));
            let expected = cell_spec.expected_value();
            assert_eq!(actual, expected);
        }
    }
}
```

## Benefits

- **Human-readable**: Easy to read and edit in any text editor
- **Version control friendly**: Line-oriented diffs show exactly what changed
- **No Excel dependency**: Create tests without needing Excel installed
- **Deterministic**: Same input always produces same output
- **Reviewable**: Easy to review in pull requests

## Validation

All spec files should be valid JSON and conform to the schema defined in RFC 004.

Validate a spec file:

```bash
# Check JSON syntax
python -m json.tool test-cases/specs/basic/arithmetic.json

# Or use jq
jq . test-cases/specs/basic/arithmetic.json
```

## See Also

- [RFC 004 Specification](../../specs/RFC_004_TEXT_FORMAT.md) - Complete format specification
- [Testing Strategy](../../docs/TESTING_STRATEGY.md) - Overall testing approach
- [Create Spec Workflow](../../.agent/workflows/create-spec.md) - Step-by-step guide
