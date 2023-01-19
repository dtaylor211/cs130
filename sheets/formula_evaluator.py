from lark import Tree, Transformer, Lark, Token
from decimal import Decimal
import sys # remove

parser = Lark.open('formulas.lark', start='formula', rel_to=__file__)

class Evaluator(Transformer):
    '''
    Evaluate an input string as a formula, based on the described Lark grammar
    in ./formulas.lark

    '''

    def __init__(self):
        # Initialize an evaluator, as well as the Transformer

        Transformer.__init__(self)

    def NUMBER(self, token: Token):
        # Evaluate a NUMBER type into a Decimal object   
        #
        # Make sure to normalize 
        norm = Decimal(token).normalize()
        _, _, exp = norm.as_tuple()
        if exp <= 0:
            return norm
        else:
            return norm.quantize(1)
    
    def STRING(self, token: Token):
        # Evaluate a STRING type

        return str(token)[0][1:-1]

    def ERROR_VALUE(self, token: Token):
        # Evaluate an ERROR_VALUE within the Tree
        pass
        # return CellError(token)

    def add_expr(self, args):
        # Evaluate an addition expression within the Tree
        #
        # Right now this does not work for complex expressions

        operator = args[1]
        a = args[0].children[-1]
        b = args[2].children[-1]

        if operator == '+':
            result =  a + b
        else:
            result = a - b
        return Tree('number', [Decimal(result)])

    def expr(self, args):
        # to-do

        return eval(args[0])
    
    def mul_expr(self, args):
        # Evaluate a multiplication expression within the Tree
        #
        # Right now this does not work for complex expressions

        operator = args[1]
        a = self.transform(args[0]).children[-1]
        b = self.transform(args[2]).children[-1]
        if operator == '*':
            result = a * b
        else:
            result = a / b
        return Tree('number', [Decimal(result)])

    def unary_op(self, args):
        # Evaluate a unary operation within the Tree
        # 
        # Need to check if this works

        operator = args[0]
        a = args[1].children[-1]
        result = -1 * a if operator == '-' else a
        return Tree('number', [Decimal(result)])