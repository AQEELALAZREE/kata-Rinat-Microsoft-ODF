# ODF Formula Specifications

This directory contains formula specifications in ODF (OpenDocument Formula) format.

## Purpose

These specs define:
1. Formula expressions to test
2. Expected results
3. Excel behavior documentation
4. Edge cases and quirks

## Format

Each spec file is XML following this structure:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<formula-specs category="math">
  <formula name="SUM_basic">
    <expression>=SUM(1,2,3)</expression>
    <expected type="number">6</expected>
    <description>Basic SUM with literal numbers</description>
  </formula>
  
  <formula name="SUM_range">
    <expression>=SUM(A1:A3)</expression>
    <expected type="number">6</expected>
    <description>SUM with range reference</description>
    <setup>
      <cell address="A1" value="1"/>
      <cell address="A2" value="2"/>
      <cell address="A3" value="3"/>
    </setup>
  </formula>
  
  <formula name="DIV_by_zero">
    <expression>=1/0</expression>
    <expected type="error">#DIV/0!</expected>
    <description>Division by zero produces #DIV/0! error</description>
  </formula>
</formula-specs>
```

## Categories

Organize specs by function category:

- `math.xml` - Mathematical functions (SUM, AVERAGE, ROUND, etc.)
- `logical.xml` - Logical functions (IF, AND, OR, etc.)
- `text.xml` - Text functions (CONCATENATE, LEFT, RIGHT, etc.)
- `lookup.xml` - Lookup functions (VLOOKUP, INDEX, MATCH, etc.)
- `date.xml` - Date/time functions (TODAY, DATE, YEAR, etc.)
- `operators.xml` - Operators (+, -, *, /, ^, &, etc.)
- `edge-cases.xml` - Edge cases and quirks

## Expected Value Types

- `type="number"` - Numeric value
- `type="string"` - Text value
- `type="boolean"` - TRUE or FALSE
- `type="error"` - Excel error (#DIV/0!, #N/A, etc.)
- `type="auto"` - Auto-detect (default)

## Workflow

### 1. Create Spec

```bash
# Create new spec file
cat > specs/odf-formulas/math.xml << 'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<formula-specs category="math">
  <formula name="SUM_basic">
    <expression>=SUM(1,2,3)</expression>
    <expected type="number">6</expected>
    <description>Basic SUM with literal numbers</description>
  </formula>
</formula-specs>
EOF
```

### 2. Generate Excel Test File

```bash
python scripts/generate_excel_tests.py specs/odf-formulas/ -o test-cases/generated/formulas.xlsx
```

### 3. Validate with Excel

```bash
python scripts/validate_with_excel.py test-cases/generated/formulas.xlsx
```

### 4. Update Specs Based on Results

If Excel behavior differs from expected:
- Update the `<expected>` value
- Add notes about Excel quirks
- Document version-specific behavior

## Example Specs

### Basic Arithmetic

```xml
<formula name="addition">
  <expression>=1+2</expression>
  <expected>3</expected>
</formula>

<formula name="precedence">
  <expression>=1+2*3</expression>
  <expected>7</expected>
  <description>Multiplication before addition</description>
</formula>
```

### Type Coercion

```xml
<formula name="text_to_number">
  <expression>="5"+3</expression>
  <expected>8</expected>
  <description>Text "5" coerced to number in arithmetic</description>
</formula>

<formula name="boolean_arithmetic">
  <expression>=TRUE+1</expression>
  <expected>2</expected>
  <description>TRUE = 1 in arithmetic</description>
</formula>
```

### Error Handling

```xml
<formula name="error_propagation">
  <expression>=1+#DIV/0!</expression>
  <expected type="error">#DIV/0!</expected>
  <description>Errors propagate through formulas</description>
</formula>
```

## Best Practices

1. **One concept per spec** - Each formula tests one specific behavior
2. **Clear names** - Use descriptive names like `SUM_with_text_values`
3. **Document quirks** - Add notes about unexpected Excel behavior
4. **Include edge cases** - Empty cells, errors, type mismatches
5. **Group by category** - Keep related formulas together

## Validation Reports

After running validation, check:
- `validation-report.json` - Full results
- `mismatches.txt` - Specs that don't match Excel
- Update specs based on actual Excel behavior

## Contributing

When adding new specs:
1. Test in real Excel first
2. Document the actual behavior
3. Include edge cases
4. Add to appropriate category file
5. Run validation to confirm
