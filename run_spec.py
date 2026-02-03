import json
import sys
from pathlib import Path
from excel_engine.engine import FormulaEngine

def run_spec(spec_data):
    engine = FormulaEngine("excel_engine/parser/grammar.lark")
    
    # Build global context: map "SheetName!Addr" and "Addr" (for current sheet) to values
    global_context = {}
    for sheet in spec_data["sheets"]:
        sheet_name = sheet["name"]
        for addr, cell in sheet["cells"].items():
            # Add sheet-qualified address
            global_context[f"{sheet_name}!{addr}"] = cell
    
    passed = 0
    failed = 0
    total_specs = 0
    
    for sheet in spec_data["sheets"]:
        sheet_name = sheet["name"]
        cells = sheet["cells"]
        
        # Create a local context for the current sheet that includes unqualified addresses
        local_context = global_context.copy()
        for addr, cell in cells.items():
            local_context[addr] = cell
            
        print(f"\nRunning specs for sheet: {sheet_name}")
        print("-" * 40)
        
        # Sort addresses for consistent output
        sorted_addresses = sorted(cells.keys(), key=lambda x: (x[0], int(x[1:]) if x[1:].isdigit() else 0))
        
        for addr in sorted_addresses:
            cell = cells[addr]
            if "formula" in cell:
                total_specs += 1
                formula = cell["formula"]
                expected_value = cell.get("value")
                expected_error = cell.get("error")
                
                try:
                    actual = engine.run_formula(formula, context=local_context, current_addr=addr)
                    
                    # Comparison logic
                    match = False
                    if expected_error is not None:
                        match = actual == expected_error
                        expected_str = expected_error
                    else:
                        if isinstance(actual, (int, float)) and isinstance(expected_value, (int, float)):
                            match = abs(actual - expected_value) < 1e-10
                        elif isinstance(actual, list) and isinstance(expected_value, list):
                            match = actual == expected_value
                        else:
                            match = str(actual) == str(expected_value)
                        expected_str = expected_value
                    
                    if match:
                        print(f"✓ {addr}: {formula} -> {actual} (Expected: {expected_str})")
                        passed += 1
                    else:
                        print(f"✗ {addr}: {formula} -> {actual} (Expected: {expected_str})")
                        failed += 1
                except Exception as e:
                    expected_str = expected_error if expected_error else expected_value
                    print(f"✗ {addr}: {formula} -> ERROR: {e} (Expected: {expected_str})")
                    failed += 1
                    
    print("-" * 40)
    print(f"Results: {passed} passed, {failed} failed out of {total_specs} spec formulas.")
    return passed, failed

if __name__ == "__main__":
    if len(sys.argv) < 2:
        # Try reading from stdin
        try:
            input_data = sys.stdin.read()
            if not input_data:
                print("Usage: python run_spec.py <spec_file.json>")
                sys.exit(1)
            spec = json.loads(input_data)
        except Exception:
            print("Usage: python run_spec.py <spec_file.json>")
            sys.exit(1)
    else:
        with open(sys.argv[1], 'r') as f:
            spec = json.load(f)
            
    run_spec(spec)
