from .parser.parser import FormulaParser
from .functions import FUNCTIONS, ERR_VALUE, ERR_DIV0, ERR_NUM, ERR_NA, ERR_REF, ERR_NAME, ERR_NULL, to_num

class FormulaEngine:
    def __init__(self, grammar_path):
        self.parser = FormulaParser(grammar_path)
    
    def _expand_range(self, start_ref, end_ref):
        def parse_addr(addr):
            import re
            m = re.match(r"([A-Z]+)([0-9]+)", addr)
            if not m: return "A", 1
            return m.group(1), int(m.group(2))
        
        start_col, start_row = parse_addr(start_ref["address"])
        end_col, end_row = parse_addr(end_ref["address"])
        
        def col_to_int(col):
            num = 0
            for char in col: num = num * 26 + (ord(char) - ord('A') + 1)
            return num
        
        def int_to_col(n):
            name = ""
            while n > 0: n, rem = divmod(n - 1, 26); name = chr(65 + rem) + name
            return name
        
        c1, c2 = col_to_int(start_col), col_to_int(end_col)
        r1, r2 = start_row, end_row
        if c1 > c2: c1, c2 = c2, c1
        if r1 > r2: r1, r2 = r2, r1
        
        range_cells = []
        for r in range(r1, r2 + 1):
            row_cells = []
            for c in range(c1, c2 + 1):
                sheet = start_ref.get("sheet")
                addr = f"{int_to_col(c)}{r}"
                row_cells.append({"type": "cell_reference", "sheet": sheet, "address": addr})
            range_cells.append(row_cells)
        return range_cells

    def evaluate(self, ast, context=None, current_addr=None):
        ERROR_VALUES = {ERR_VALUE, ERR_DIV0, ERR_NUM, ERR_NA, ERR_REF, ERR_NAME, ERR_NULL}
        if context is None: context = {}

        if isinstance(ast, dict):
            t = ast.get("type")
            if t == "function":
                func_name = ast["name"]
                args_ast = ast["args"]

                if func_name == "IF":
                    if len(args_ast) < 2: return ERR_VALUE
                    condition = self.evaluate(args_ast[0], context, current_addr)
                    if condition in ERROR_VALUES: return condition
                    is_true = bool(to_num(condition)) if not isinstance(condition, bool) else condition
                    if is_true: return self.evaluate(args_ast[1], context, current_addr)
                    return self.evaluate(args_ast[2], context, current_addr) if len(args_ast) > 2 else False
                
                elif func_name == "IFERROR":
                    val = self.evaluate(args_ast[0], context, current_addr)
                    return self.evaluate(args_ast[1], context, current_addr) if val in ERROR_VALUES else val

                # Eager evaluation for others, but special for context-aware
                if func_name in ("ROW", "COLUMN", "ROWS", "COLUMNS", "OFFSET", "INDIRECT", "ADDRESS", "ISREF"):
                    args = []
                    for arg in args_ast:
                        if isinstance(arg, dict) and arg.get("type") in ("cell_reference", "range_reference"):
                            args.append(arg)
                        else:
                            args.append(self.evaluate(arg, context, current_addr))
                else:
                    args = [self.evaluate(arg, context, current_addr) for arg in args_ast]

                if func_name not in ("ISERROR", "ISERR", "ISNA", "ISREF", "TYPE", "ERROR.TYPE"):
                    for arg in args:
                        if isinstance(arg, str) and arg in ERROR_VALUES: return arg
                
                func = FUNCTIONS.get(func_name)
                if not func: raise ValueError(f"{ERR_NAME} (Unknown function: {func_name})")
                
                if func_name in ("ROW", "COLUMN", "ROWS", "COLUMNS", "OFFSET", "INDIRECT", "ADDRESS"):
                    return func(*args, context=context, engine=self, current_addr=current_addr)
                return func(*args)

            elif t == "cell_reference":
                sheet, address = ast.get("sheet"), ast["address"]
                lookup_addr = f"{sheet}!{address}" if sheet else address
                if lookup_addr in context:
                    cell = context[lookup_addr]
                    if "error" in cell: return cell["error"]
                    return cell.get("value")
                return None
            
            elif t == "range_reference":
                expanded = self._expand_range(ast["start"], ast["end"])
                return [[self.evaluate(cell, context, current_addr) for cell in row] for row in expanded]
            
            elif t == "unary":
                val = self.evaluate(ast["value"], context, current_addr)
                if val in ERROR_VALUES: return val
                nv = to_num(val); return -nv if ast["op"] == '-' else nv
            
            elif t == "op":
                left = self.evaluate(ast["left"], context, current_addr)
                if left in ERROR_VALUES: return left
                right = self.evaluate(ast["right"], context, current_addr)
                if right in ERROR_VALUES: return right
                
                op = ast["op"]
                if op == '&': return str(left if left is not None else "") + str(right if right is not None else "")
                if op in ('+', '-', '*', '/', '^'):
                    lv, rv = to_num(left), to_num(right)
                    if op == '+': return lv + rv
                    if op == '-': return lv - rv
                    if op == '*': return lv * rv
                    if op == '/': return ERR_DIV0 if rv == 0 else lv / rv
                    if op == '^': return lv ** rv
                if op == '=': return left == right
                if op == '<>': return left != right
                try:
                    if op == '<': return left < right
                    if op == '>': return left > right
                    if op == '<=': return left <= right
                    if op == '>=': return left >= right
                except: return ERR_VALUE
            return None
        elif isinstance(ast, list): return [[self.evaluate(cell, context, current_addr) for cell in row] for row in ast]
        return ast

    def run_formula(self, formula_str, context=None, current_addr=None):
        if not formula_str or not str(formula_str).startswith("="): return formula_str
        try:
            ast = self.parser.parse(formula_str)
            return self.evaluate(ast, context, current_addr)
        except Exception as e:
            msg = str(e)
            for err in (ERR_NAME, ERR_VALUE, ERR_DIV0, ERR_NUM, ERR_NA, ERR_REF, ERR_NULL):
                if err in msg: return err
            return f"#ERROR: {msg}"
