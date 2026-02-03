#!/usr/bin/env python3
"""
Validate formula specs against real Excel behavior.

This script:
1. Opens an Excel file with formulas
2. Lets Excel calculate them
3. Extracts the results
4. Compares with expected values from specs
5. Reports mismatches

Usage:
    python validate_with_excel.py test-cases/formulas.xlsx
    python validate_with_excel.py test-cases/formulas.xlsx --specs specs/odf-formulas/
"""

import sys
import json
import argparse
from pathlib import Path

# Platform-specific Excel automation
try:
    # Windows
    import win32com.client
    PLATFORM = 'windows'
except ImportError:
    try:
        # macOS
        from appscript import app, k
        PLATFORM = 'mac'
    except ImportError:
        print("Error: Install pywin32 (Windows) or py-appscript (Mac)")
        sys.exit(1)


class ExcelValidator:
    """Validate formulas using real Excel"""
    
    def __init__(self, excel_file):
        self.excel_file = Path(excel_file)
        self.results = {}
        self.mismatches = []
        
    def open_excel(self):
        """Open Excel application"""
        if PLATFORM == 'windows':
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.workbook = self.excel.Workbooks.Open(str(self.excel_file.absolute()))
        else:  # mac
            self.excel = app('Microsoft Excel')
            self.workbook = self.excel.open(str(self.excel_file.resolve()))
    
    def close_excel(self):
        """Close Excel application"""
        if PLATFORM == 'windows':
            self.workbook.Close(SaveChanges=False)
            self.excel.Quit()
        else:  # mac
            self.workbook.close(saving=k.no)
    
    def calculate_all(self):
        """Force Excel to calculate all formulas"""
        if PLATFORM == 'windows':
            self.excel.CalculateFull()
        else:  # mac
            # Excel for macOS recalculates automatically when opening/reading.
            # appscript does not expose an application-level 'calculate' command.
            pass
    
    def extract_results(self):
        """Extract all formula results from Excel"""
        results = {}
        
        if PLATFORM == 'windows':
            for sheet in self.workbook.Sheets:
                sheet_name = sheet.Name
                for row in range(1, sheet.UsedRange.Rows.Count + 1):
                    for col in range(1, sheet.UsedRange.Columns.Count + 1):
                        cell = sheet.Cells(row, col)
                        if cell.HasFormula:
                            address = f"{sheet_name}!{cell.Address(False, False)}"
                            formula = cell.Formula
                            value = cell.Value
                            results[address] = {
                                'formula': formula,
                                'value': value
                            }
        else:  # mac
            for sheet in self.workbook.sheets():
                sheet_name = sheet.name()
                used_range = sheet.used_range()
                for row in used_range.rows():
                    for cell in row.cells():
                        if cell.formula():
                            address = f"{sheet_name}!{cell.address()}"
                            results[address] = {
                                'formula': cell.formula(),
                                'value': cell.value()
                            }
        
        return results
    
    def validate(self, expected_results=None):
        """
        Validate Excel results against expected values
        
        Args:
            expected_results: Dict of {cell_address: expected_value}
        """
        print(f"Opening {self.excel_file}...")
        self.open_excel()
        
        print("Calculating formulas...")
        self.calculate_all()
        
        print("Extracting results...")
        self.results = self.extract_results()
        
        print(f"Found {len(self.results)} formulas")
        
        if expected_results:
            print("\nValidating against expected results...")
            self.compare_results(expected_results)
        
        self.close_excel()
        
        return self.results
    
    def compare_results(self, expected):
        """Compare Excel results with expected values"""
        for address, expected_value in expected.items():
            if address not in self.results:
                self.mismatches.append({
                    'cell': address,
                    'error': 'Formula not found in Excel file',
                    'expected': expected_value
                })
                continue
            
            actual_value = self.results[address]['value']
            
            if not self.values_match(actual_value, expected_value):
                self.mismatches.append({
                    'cell': address,
                    'formula': self.results[address]['formula'],
                    'expected': expected_value,
                    'actual': actual_value
                })
    
    def values_match(self, actual, expected, tolerance=1e-10):
        """Check if two values match (with tolerance for floats)"""
        # Handle None/null
        if actual is None and expected is None:
            return True
        if actual is None or expected is None:
            return False
        
        # Handle numbers
        if isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
            return abs(actual - expected) < tolerance
        
        # Handle strings (case-insensitive like Excel)
        if isinstance(actual, str) and isinstance(expected, str):
            return actual.upper() == expected.upper()
        
        # Exact match for other types
        return actual == expected
    
    def print_report(self):
        """Print validation report"""
        print("\n" + "="*60)
        print("VALIDATION REPORT")
        print("="*60)
        
        total = len(self.results)
        passed = total - len(self.mismatches)
        
        print(f"\nTotal formulas: {total}")
        print(f"Passed: {passed}")
        print(f"Failed: {len(self.mismatches)}")
        print(f"Success rate: {passed/total*100:.1f}%")
        
        if self.mismatches:
            print("\n" + "-"*60)
            print("MISMATCHES:")
            print("-"*60)
            for mismatch in self.mismatches:
                print(f"\nCell: {mismatch['cell']}")
                if 'formula' in mismatch:
                    print(f"Formula: {mismatch['formula']}")
                print(f"Expected: {mismatch['expected']}")
                print(f"Actual: {mismatch.get('actual', 'N/A')}")
                if 'error' in mismatch:
                    print(f"Error: {mismatch['error']}")
    
    def save_report(self, output_file):
        """Save validation report to JSON"""
        report = {
            'total_formulas': len(self.results),
            'passed': len(self.results) - len(self.mismatches),
            'failed': len(self.mismatches),
            'results': self.results,
            'mismatches': self.mismatches
        }
        
        with open(output_file, 'w') as f:
            json.dump(report, f, indent=2)
        
        print(f"\nReport saved to: {output_file}")


def load_expected_results(specs_dir):
    """Load expected results from spec files"""
    # TODO: Implement ODF spec parser
    # For now, return empty dict
    return {}


def main():
    parser = argparse.ArgumentParser(description='Validate formulas with Excel')
    parser.add_argument('excel_file', help='Excel file with formulas')
    parser.add_argument('--specs', help='Directory with ODF specs')
    parser.add_argument('--output', default='validation-report.json',
                       help='Output report file')
    
    args = parser.parse_args()
    
    # Load expected results if specs provided
    expected = {}
    if args.specs:
        expected = load_expected_results(args.specs)
    
    # Validate
    validator = ExcelValidator(args.excel_file)
    validator.validate(expected)
    
    # Print and save report
    validator.print_report()
    validator.save_report(args.output)
    
    # Exit with error code if there are mismatches
    sys.exit(1 if validator.mismatches else 0)


if __name__ == '__main__':
    main()
