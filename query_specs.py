import json
import os
import argparse
import re
from pathlib import Path

def get_all_specs(root_dir):
    specs_dir = Path(root_dir)
    return list(specs_dir.glob("**/*.json"))

def query_specs(root_dir, func_search=None, sheet_search=None, error_search=None, list_funcs=False, stats=False):
    json_files = get_all_specs(root_dir)
    
    results = []
    unique_functions = set()
    total_cells = 0
    total_formulas = 0
    
    for json_file in json_files:
        try:
            with open(json_file, 'r') as f:
                data = json.load(f)
            
            if "sheets" not in data:
                continue
                
            file_rel = json_file.relative_to(root_dir)
            
            for sheet in data["sheets"]:
                sheet_name = sheet["name"]
                
                # Check for sheet search
                if sheet_search and sheet_search.lower() not in sheet_name.lower():
                    # We still need to check formulas within this sheet for the sheet name
                    pass

                for addr, cell in sheet["cells"].items():
                    total_cells += 1
                    formula = cell.get("formula", "")
                    error = cell.get("error")
                    
                    if formula:
                        total_formulas += 1
                        # Extract functions (approximate)
                        funcs = re.findall(r"([A-Z][A-Z0-9_.]+)\(", formula)
                        for f in funcs:
                            unique_functions.add(f.upper())
                            
                        # Matches
                        match = False
                        if func_search:
                            if any(func_search.upper() == f.upper() for f in funcs):
                                match = True
                        
                        if sheet_search:
                            # Search for "SheetName!" in formula
                            if f"{sheet_search}!" in formula or f"'{sheet_search}'!" in formula:
                                match = True
                                
                        if error_search:
                            if error and error_search.upper() in error.upper():
                                match = True
                                
                        if match or (not func_search and not sheet_search and not error_search and not list_funcs and not stats):
                            results.append({
                                "file": str(file_rel),
                                "sheet": sheet_name,
                                "address": addr,
                                "formula": formula,
                                "value": cell.get("value"),
                                "error": error
                            })
        except Exception:
            continue

    if list_funcs:
        print("\nUnique Functions Used Across Specs:")
        print("-" * 30)
        for f in sorted(list(unique_functions)):
            print(f"- {f}")
        return

    if stats:
        print("\nSpec Suite Statistics:")
        print("-" * 30)
        print(f"Total Spec Files:  {len(json_files)}")
        print(f"Total Cells:       {total_cells}")
        print(f"Total Formulas:    {total_formulas}")
        print(f"Unique Functions:  {len(unique_functions)}")
        return

    if results:
        print(f"\nFound {len(results)} matching cells:")
        print("-" * 80)
        for res in results:
            err_str = f" [Error: {res['error']}]" if res['error'] else ""
            print(f"[{res['file']}] {res['sheet']}!{res['address']}: {res['formula']} -> {res['value']}{err_str}")
    else:
        print("No matches found.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Query Excel Spec Files")
    parser.add_argument("--dir", default="specs", help="Directory containing specs")
    parser.add_argument("--func", help="Search for a specific function (e.g., SUM)")
    parser.add_argument("--sheet", help="Search for formulas referencing a specific sheet")
    parser.add_argument("--error", help="Search for specific error codes (e.g., #DIV/0!)")
    parser.add_argument("--list-functions", action="store_true", help="List all unique functions found")
    parser.add_argument("--stats", action="store_true", help="Show overview statistics")
    
    args = parser.parse_args()
    
    query_specs(args.dir, args.func, args.sheet, args.error, args.list_functions, args.stats)
