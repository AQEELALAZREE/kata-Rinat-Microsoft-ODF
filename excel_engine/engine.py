from .parser.parser import FormulaParser
from .functions import FUNCTIONS

class FormulaEngine:
    def __init__(self, grammar_path):
        self.parser = FormulaParser(grammar_path)
    
    def evaluate(self, ast, context=None):
        if isinstance(ast, dict):
            if ast.get("type") == "function":
                func_name = ast["name"]
                args = [self.evaluate(arg, context) for arg in ast["args"]]
                func = FUNCTIONS.get(func_name)
                if not func:
                    raise ValueError(f"#NAME? (Unknown function: {func_name})")
                return func(*args)
            elif ast.get("type") == "cell_reference":
                address = ast["address"]
                if context and address in context:
                    cell = context[address]
                    if "formula" in cell:
                        # Simple non-recursive for now, or just return value if already computed
                        return cell.get("value")
                    return cell.get("value")
                return 0 # Default for empty cell
            elif ast.get("type") == "unary":
                op = ast["op"]
                val = self.evaluate(ast["value"], context)
                try:
                    num_val = float(val) if val is not None else 0
                    return -num_val if op == '-' else num_val
                except (ValueError, TypeError):
                    return "#VALUE!"
            elif ast.get("type") == "op":
                left = self.evaluate(ast["left"], context)
                right = self.evaluate(ast["right"], context)
                op = ast["op"]
                
                # Coerce to numeric if possible for arithmetic ops
                if op in ('+', '-', '*', '/', '^'):
                    try:
                        left = float(left) if left is not None else 0
                        right = float(right) if right is not None else 0
                    except (ValueError, TypeError):
                        return "#VALUE!"

                if op == '+': return left + right
                if op == '-': return left - right
                if op == '*': return left * right
                if op == '/': 
                    if right == 0: return "#DIV/0!"
                    return left / right
                if op == '^': return left ** right
                
                if op == '=': return left == right
                if op == '<>': return left != right
                if op == '<': return left < right
                if op == '>': return left > right
                if op == '<=': return left <= right
                if op == '>=': return left >= right
                
                return None
            else:
                return None
        elif isinstance(ast, list):
            # Array literal - elements might be expressions
            return [[self.evaluate(cell, context) for cell in row] for row in ast]
        else:
            return ast

    def run_formula(self, formula_str, context=None):
        if not formula_str or not str(formula_str).startswith("="):
            return formula_str
        try:
            ast = self.parser.parse(formula_str)
            return self.evaluate(ast, context)
        except Exception as e:
            return f"#ERROR: {str(e)}"
