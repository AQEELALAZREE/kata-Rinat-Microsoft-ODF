from lark import Lark, Transformer

class FormulaTransformer(Transformer):
    def _log(self, msg):
        with open("transformer_debug.log", "a") as f:
            f.write(str(msg) + "\n")

    def formula(self, items):
        self._log(f"formula: {items}")
        return items[0]
    
    def number(self, items):
        return float(items[0])
    
    def string(self, items):
        return items[0][1:-1].replace('""', '"')
    
    def true(self, _):
        return True
    
    def false(self, _):
        return False
    
    def array(self, items):
        return items[0]

    def array_literal(self, items):
        return items
    
    def array_row(self, items):
        return items
    
    def sheet_cell_ref(self, items):
        sheet = str(items[0]).strip("'")
        address = str(items[1]).replace("$", "")
        return {"type": "cell_reference", "sheet": sheet, "address": address}

    def cell_ref(self, items):
        address = str(items[0]).replace("$", "")
        return {"type": "cell_reference", "sheet": None, "address": address}

    def range_reference(self, items):
        if len(items) == 3:
            return {"type": "range_reference", "start": items[0], "end": items[2]}
        return {"type": "range_reference", "start": items[0], "end": items[1]}

    def function_call(self, items):
        func_name = str(items[0]).upper()
        args = items[1] if len(items) > 1 and items[1] is not None else []
        return {"type": "function", "name": func_name, "args": args}

    def arguments(self, items):
        return items

    def unary_expr(self, items):
        if len(items) == 1:
            return items[0]
        return {"type": "unary", "op": str(items[0]), "value": items[1]}

    def add_expr(self, items):
        if len(items) == 1:
            return items[0]
        res = items[0]
        for i in range(1, len(items), 2):
            res = {"type": "op", "op": str(items[i]), "left": res, "right": items[i+1]}
        return res

    def concat(self, items):
        if len(items) == 1:
            return items[0]
        res = items[0]
        for i in range(1, len(items), 2):
            res = {"type": "op", "op": "&", "left": res, "right": items[i+1]}
        return res

    def mult_expr(self, items):
        if len(items) == 1:
            return items[0]
        res = items[0]
        for i in range(1, len(items), 2):
            res = {"type": "op", "op": str(items[i]), "left": res, "right": items[i+1]}
        return res

    def power_expr(self, items):
        if len(items) == 1:
            return items[0]
        res = items[0]
        for i in range(1, len(items), 2):
            res = {"type": "op", "op": "^", "left": res, "right": items[i+1]}
        return res

    def comparison(self, items):
        if len(items) == 1:
            return items[0]
        res = items[0]
        for i in range(1, len(items), 2):
            res = {"type": "op", "op": str(items[i]), "left": res, "right": items[i+1]}
        return res

class FormulaParser:
    def __init__(self, grammar_path):
        with open(grammar_path, 'r') as f:
            grammar = f.read()
        self.lark = Lark(grammar, start='formula', parser='earley')
        self.transformer = FormulaTransformer()
    
    def parse(self, formula):
        tree = self.lark.parse(formula)
        return self.transformer.transform(tree)
