---
description: Create a new Excel spec file for testing
---

# Create New Excel Spec File

This workflow guides you through creating a new JSON-based spec file for testing Excel formulas.

## Steps

1. **Determine the test category**
   - `basic/` - Simple arithmetic, basic formulas
   - `functions/` - Function-specific tests (SUM, IF, VLOOKUP, etc.)
   - `errors/` - Error handling and propagation
   - `edge_cases/` - Precision, type coercion, circular references

2. **Create the spec file**
   ```bash
   # Choose appropriate category and name
   touch test-cases/specs/<category>/<name>.json
   ```

3. **Use the template**
   ```json
   {
     "version": 1,
     "description": "Description of what this spec tests",
     "meta": {
       "locale": "en-US",
       "decimal_separator": ".",
       "list_separator": ","
     },
     "sheets": [
       {
         "name": "Sheet1",
         "cells": {
         }
       }
     ]
   }
   ```

4. **Add cells**
   - For values: `"A1": { "value": 10 }`
   - For formulas: `"A2": { "formula": "=A1*2", "value": 20 }`
   - For errors: `"A3": { "formula": "=A1/0", "error": "#DIV/0!" }`

5. **Validate the spec**
   ```bash
   # Validate JSON syntax
   python -m json.tool test-cases/specs/<category>/<name>.json
   
   # Or use jq
   jq . test-cases/specs/<category>/<name>.json
   ```

6. **Create corresponding test**
   - Add test case in `tests/` directory
   - Load the spec and validate results

## Example

Creating a test for AVERAGE function:

```bash
# Create file
cat > test-cases/specs/functions/average.json << 'EOF'
{
  "version": 1,
  "description": "AVERAGE function with range",
  "meta": {
    "locale": "en-US",
    "decimal_separator": ".",
    "list_separator": ","
  },
  "sheets": [
    {
      "name": "Data",
      "cells": {
        "A1": { "value": 10 },
        "A2": { "value": 20 },
        "A3": { "value": 30 },
        "B1": { "formula": "=AVERAGE(A1:A3)", "value": 20 }
      }
    }
  ]
}
EOF

# Validate
jq . test-cases/specs/functions/average.json
```

## Tips

- Keep specs small (dozens of cells, not hundreds)
- Use descriptive names for sheets and clear descriptions
- Include edge cases in your test data
- Test both successful calculations and error conditions
- For multi-locale tests, create separate specs with different `meta` settings
