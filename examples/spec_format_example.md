# Spec Format Example

This example demonstrates how to use the RFC 004 spec format to load test cases and validate formula calculations.

## Overview

The spec format allows you to define Excel workbooks in JSON for testing purposes. This example shows:

1. Loading a spec file
2. Converting it to a workbook
3. Running the calculation engine
4. Validating results

## Example Spec File

See `test-cases/specs/basic/arithmetic.json` for a complete example.

## Usage Example (Pseudocode)

```rust
use spec_format::{SpecParser, WorkbookSpec};
use formula_engine::FormulaEngine;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 1. Load spec from file
    let spec = SpecParser::parse_file("test-cases/specs/basic/arithmetic.json")?;
    
    println!("Testing: {}", spec.description.unwrap_or_default());
    println!("Locale: {}", spec.meta.locale);
    
    // 2. Convert spec to workbook
    let mut workbook = spec.to_workbook();
    
    // 3. Create and run calculation engine
    let engine = FormulaEngine::new();
    engine.calculate(&mut workbook)?;
    
    // 4. Validate results
    let mut passed = 0;
    let mut failed = 0;
    
    for sheet in &spec.sheets {
        println!("\nValidating sheet: {}", sheet.name);
        
        for (addr, cell_spec) in &sheet.cells {
            let cell_ref = format!("{}!{}", sheet.name, addr);
            let actual = workbook.get_value(&cell_ref);
            let expected = cell_spec.expected_value();
            
            if values_match(&actual, &expected) {
                println!("  ✓ {} = {:?}", addr, actual);
                passed += 1;
            } else {
                println!("  ✗ {} expected {:?}, got {:?}", addr, expected, actual);
                failed += 1;
            }
        }
    }
    
    println!("\nResults: {} passed, {} failed", passed, failed);
    
    if failed > 0 {
        std::process::exit(1);
    }
    
    Ok(())
}

// Helper function to compare values with tolerance for floating point
fn values_match(actual: &CellValue, expected: &CellValue) -> bool {
    match (actual, expected) {
        (CellValue::Number(a), CellValue::Number(e)) => {
            (a - e).abs() < 1e-10
        }
        (CellValue::String(a), CellValue::String(e)) => a == e,
        (CellValue::Boolean(a), CellValue::Boolean(e)) => a == e,
        (CellValue::Null, CellValue::Null) => true,
        (CellValue::Error(a), CellValue::Error(e)) => a == e,
        _ => false,
    }
}
```

## Creating a Spec Programmatically

```rust
use spec_format::{WorkbookSpec, MetaData, SheetSpec, CellSpec, CellValue};
use std::collections::BTreeMap;

fn create_test_spec() -> WorkbookSpec {
    let mut cells = BTreeMap::new();
    
    // Add values
    cells.insert("A1".to_string(), CellSpec::Value(CellValue::Number(10.0)));
    cells.insert("A2".to_string(), CellSpec::Value(CellValue::Number(20.0)));
    
    // Add formula
    cells.insert("A3".to_string(), CellSpec::Formula {
        formula: "=A1+A2".to_string(),
        value: CellValue::Number(30.0),
    });
    
    let sheet = SheetSpec {
        name: "Sheet1".to_string(),
        cells,
    };
    
    WorkbookSpec {
        version: 1,
        description: Some("Programmatically created spec".to_string()),
        meta: MetaData {
            locale: "en-US".to_string(),
            decimal_separator: ".".to_string(),
            list_separator: ",".to_string(),
            timezone: None,
        },
        sheets: vec![sheet],
    }
}

fn main() {
    let spec = create_test_spec();
    
    // Serialize to JSON
    let json = SpecWriter::write(&spec).unwrap();
    println!("{}", json);
    
    // Save to file
    SpecWriter::write_file(&spec, "my_test.json").unwrap();
}
```

## Running Tests

```bash
# Run all spec-based tests
cargo test spec_

# Run specific test
cargo test test_arithmetic_spec

# Run with output
cargo test spec_ -- --nocapture
```

## See Also

- [RFC 004 Specification](../specs/RFC_004_TEXT_FORMAT.md)
- [Example Spec Files](../test-cases/specs/)
- [Testing Strategy](../docs/TESTING_STRATEGY.md)
