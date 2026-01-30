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
    
    def cell_reference(self, items):
        return {"type": "cell_reference", "address": str(items[0]).replace("$", "")}

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
        self.lark = Lark(grammar, start='formula', parser='lalr')
        self.transformer = FormulaTransformer()
    
    def parse(self, formula):
        tree = self.lark.parse(formula)
        return self.transformer.transform(tree)
