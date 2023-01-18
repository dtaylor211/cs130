from lark import Tree, Transformer, Lark, Token
from decimal import Decimal

parser = Lark.open('formulas.lark', start='formula', rel_to=__file__)

class Evaluator(Transformer):
    '''
    tbd

    '''
    def __init__(self):
        Transformer.__init__(self)

    def NUMBER(self, token: Token):
        return Decimal(token)
    
    def STRING(self, token: Token):
        return str(token)[0][1:-1]

    def ERROR_VALUE(self, token: Token):
        pass
        # return CellError(token)

    def add_expr(self, args):
        operator = args[1]
        a = args[0].children[-1]
        b = args[2].children[-1]
        if operator == '+':
            return a + b
        else:
            return a - b
    
    def mul_expr(self, args):
        operator = args[1]
        a = args[0].children[-1]
        b = args[2].children[-1]
        if operator == '*':
            return a * b
        else:
            return a / b