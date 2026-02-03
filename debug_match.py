from excel_engine.functions import FUNCTIONS
print(f"MATCH in FUNCTIONS: {'MATCH' in FUNCTIONS}")

from excel_engine.engine import FormulaEngine
engine = FormulaEngine("excel_engine/parser/grammar.lark")
res = engine.run_formula('=MATCH("B", {"A","B","C"}, 0)')
print(f"MATCH Result: {res}")
