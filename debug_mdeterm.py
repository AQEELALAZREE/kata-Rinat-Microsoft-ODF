from excel_engine.engine import FormulaEngine
import os

engine = FormulaEngine("excel_engine/parser/grammar.lark")
context = {
    "A1": {"value": 1}, "B1": {"value": 2},
    "A2": {"value": 3}, "B2": {"value": 4}
}
formula = "=MDETERM(A1:B2)"
try:
    res = engine.run_formula(formula, context)
    print(f"Result: {res}")
except Exception as e:
    print(f"Error: {e}")

from excel_engine.functions import FUNCTIONS
print(f"MDETERM in FUNCTIONS: {'MDETERM' in FUNCTIONS}")
print(f"MDETERM func: {FUNCTIONS.get('MDETERM')}")
