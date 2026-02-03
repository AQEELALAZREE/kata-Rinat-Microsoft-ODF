import json
import os
from pathlib import Path
from run_spec import run_spec

def run_all_specs(root_dir):
    results = []
    specs_dir = Path(root_dir)
    
    # Find all json files recursively
    json_files = sorted(list(specs_dir.glob("**/*.json")))
    
    print(f"Found {len(json_files)} spec files.")
    print("=" * 60)
    
    total_passed = 0
    total_failed = 0
    total_formulas = 0
    
    for json_file in json_files:
        rel_path = json_file.relative_to(specs_dir)
        print(f"\n[SPEC] {rel_path}")
        
        try:
            with open(json_file, 'r') as f:
                spec_data = json.load(f)
            
            # Simple check if it's a valid spec file (has sheets)
            if "sheets" not in spec_data:
                print(f"Skipping {rel_path}: Not a valid spec file (missing 'sheets')")
                continue
                
            passed, failed = run_spec(spec_data)
            results.append({
                "file": str(rel_path),
                "passed": passed,
                "failed": failed,
                "total": passed + failed
            })
            total_passed += passed
            total_failed += failed
            total_formulas += (passed + failed)
            
        except Exception as e:
            print(f"Error running {rel_path}: {e}")
            results.append({
                "file": str(rel_path),
                "error": str(e)
            })

    print("\n" + "=" * 60)
    print("SUMMARY REPORT")
    print("=" * 60)
    print(f"{'File':<50} | {'Pass':<5} | {'Fail':<5}")
    print("-" * 65)
    
    for res in results:
        if "error" in res:
            print(f"{res['file']:<50} | ERROR: {res['error']}")
        else:
            print(f"{res['file']:<50} | {res['passed']:<5} | {res['failed']:<5}")
            
    print("-" * 65)
    print(f"{'TOTAL':<50} | {total_passed:<5} | {total_failed:<5}")
    print(f"Total Formulas: {total_formulas}")
    print(f"Pass Rate: {total_passed/total_formulas*100:.1f}%" if total_formulas > 0 else "N/A")

if __name__ == "__main__":
    run_all_specs("specs")
