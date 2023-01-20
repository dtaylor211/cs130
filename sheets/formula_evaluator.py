from lark import Tree, Transformer, Lark, Token
from decimal import Decimal, DecimalException
from .cell_error import CellError, CellErrorType, cell_error_dict
from typing import *
import re
# from workbook import Workbook
import sys # remove



class Evaluator(Transformer):
    '''
    Evaluate an input string as a formula, based on the described Lark grammar
    in ./formulas.lark

    '''

    def __init__(self, workbook, sheet_name: str):
        '''
        Initialize an evaluator, as well as the Transformer

        '''

        Transformer.__init__(self)
        self.workbook = workbook
        self.working_sheet = sheet_name


    ########################################################################
    # Terminals and Bases?
    ########################################################################


    def NUMBER(self, token: Token):
        '''
        Evaluate a NUMBER type into a Decimal object   
        This method also normalizes the value stored in Decimal

        Arguments:
        - token: Token - contains data about the number

        Returns:
        - Decimal object storing the given number

        '''

        return self.normalize_number(Decimal(token))


    def STRING(self, token: Token):
        '''
        Evaluate a STRING type  

        Arguments:
        - token: Token - contains data about the string

        Returns:
        - String object with beginning and ending quotations removed
        
        '''

        return (str(token)[1:-1])
    

    def CELLREF(self, token: Token):
        '''
        Evaluate a CELLREF type

        Arguments:
        - token: Token - contains data about the cell reference

        Returns:
        - the original token passed in
        I think unneeded

        '''

        return token


    ########################################################################
    # Expressions
    ########################################################################


    def expr(self, args: List) -> Tree:
        '''
        Evaluate an expression, such as add_expr or str_concat

        Arguments:
        - args: List - list with Tree and/or Token objects

        Returns:
        - the evaluated result of the expression

        '''

        return eval(args[0])


    def add_expr(self, args: List) -> Tree:
        '''
        Evaluate an addition/subtraction expression within the Tree

        Arguments:
        - args: List - list with Tree and/or Token objects of format [x, +/-, y]

        Returns:
        - Tree holding the result as a Decimal object

        '''

        try:
            operator = args[1]
            x = args[0].children[-1]
            y = args[-1].children[-1]

            # Check for propogating errors
            if isinstance(x, CellError):
                return Tree('cell_error', [x])
            if isinstance(y, CellError):
                return Tree('cell_error', [y])
            
            if x in cell_error_dict.values():
                e_type = [i for i in cell_error_dict if cell_error_dict[i]==x]
                c = CellErrorType(e_type[0])
                return Tree('cell_error', [CellError(c, '', None)])
            if y in cell_error_dict.values():
                e_type = [i for i in cell_error_dict if cell_error_dict[i]==y]
                c = CellErrorType(e_type[0])
                return Tree('cell_error', [CellError(c, '', None)])

            # Check for compatible types, deal with empty case
            x = Decimal(0) if x is None else Decimal(x)
            y = Decimal(0) if y is None else Decimal(y)

            result = x + y if operator == '+' else x - y
            return Tree('number', [self.normalize_number(Decimal(result))])

        except Exception as e:
            return self.process_exceptions(e, detail='adding/subtracting')
    

    def mul_expr(self, args: List) -> Tree:
        '''
        Evaluate a multiplication/divison expression within the Tree

        Arguments:
        - args: List - list with Tree and/or Token objects of format [x, *//, y]

        Returns:
        - Tree holding the result as a Decimal object
        
        '''
        
        try:
            operator = args[1]
            x = self.transform(args[0]).children[-1]
            y = self.transform(args[-1]).children[-1]

            # Check for propogating errors
            if isinstance(x, CellError):
                return Tree('cell_error', [x])
            if isinstance(y, CellError):
                return Tree('cell_error', [y])

            # Check for compatible types, deal with empty case
            x = Decimal(0) if x is None else Decimal(x)
            y = Decimal(0) if y is None else Decimal(y)

            # Check for division by zero
            if y == Decimal(0) and operator == '/':
                raise ZeroDivisionError
                
            result = x * y if operator == '*' else x / y
            return Tree('number', [self.normalize_number(Decimal(result))])

        except Exception as e:
            return self.process_exceptions(e, detail='multiplication/division')    


    def unary_op(self, args: List) -> Tree:
        '''
        Evaluate an unary operation expression within the Tree

        Arguments:
        - args: List - list with Tree and/or Token objects of format [+/-, x]

        Returns:
        - Tree holding the result as a Decimal object
        
        '''

        try:
            operator = args[0]
            x = args[1].children[-1]

            # Do i need to check for errors?

            # Check for compatible types, deal with empty case
            x = Decimal(0) if x is None else Decimal(x)

            result = -1 * x if operator == '-' else x
            return Tree('number', [Decimal(result)])

        except Exception as e:
            return self.process_exceptions(e, detail='unary operations')


    def concat_expr(self, args: List) -> Tree:
        '''
        Evaluate a string concatenation expression within the Tree

        Arguments:
        - args: List - list with Tree and/or Token objects of format [s1, &, s2]

        Returns:
        - Tree holding the result as a String object

        '''
        
        try:
            s1 = self.transform(args[0]).children[-1]
            s2 = self.transform(args[-1]).children[-1]

            # Check for propogating errors
            if isinstance(s1, CellError):
                return Tree('cell_error', [s1])
            if isinstance(s2, CellError):
                return Tree('cell_error', [s2])

            # Check for compatible types, deal with empty case
            s1 = '' if s1 is None else str(s1)
            s2 = '' if s2 is None else str(s2)

            return Tree('string', [s1+s2])

        except Exception as e:
            return self.process_exceptions(e, 'string concatenation')


    def cell(self, args: List) -> Tree:
        '''
        Evaluate a cell expression

        Arguments:
        - args: List - list of Tree and/or Token objects of format [sheet, cell]

        Returns:
        - Tree holding the resulting cell as a

        Might not need this

        '''
        try:
            if len(args) == 2:
                self.working_sheet = args[0]
                cell_name = args[1]
            else:
                cell_name = args[0]

            # Check that cell location is within bounds
            if not re.match(r"^[A-Z]{1,4}[1-9][0-9]{0,3}$", cell_name.upper()):
                raise KeyError('Cell location out of bounds')
            
            result = self.workbook.get_cell_value(self.working_sheet, cell_name)

            # Check for propogating errors
            if isinstance(result, CellError):
                return Tree('cell_error', [result])

            # Deal with empty cases
            result = [Decimal(0)] if result is None else [result]
            return Tree('cell_ref', result)
        
        except KeyError as k:
            return self.process_exceptions(k)

        except Exception as e:
            return self.process_exceptions(e, 'cell operations')
        
    def parens(self, args: List):
        '''
        Evaluate an expression enclosed in parenthesis

        Arguments:
        - args: List - list of Tree and/or Token objects of format

        Returns:
        - Tree holding the expression inside the parenthesis

        '''

        return args[0] # removes the outer brackets


    ########################################################################
    # Exception Processing
    ########################################################################       


    def process_exceptions(self, ex: Exception, detail: str = '') -> Tree:
        '''
        Process the occurrence of an exception

        Arguments:
        - exception: Exception - the exception that was raised during evaluation
        - detail: str (default '') - any information to be added, usually gives
            the name of the operation being performed at the time of the error

        Returns:
        - Tree holding the exception as a CellError object
        
        '''

        error_type = None
        if isinstance(ex, DecimalException):
            detail = f'Attempting to perform {detail} on incompatible or '\
                'incorrect types'
            error_type = CellErrorType.TYPE_ERROR
        elif isinstance(ex, ZeroDivisionError):
            detail = 'Attempting to perform division with 0'
            error_type = CellErrorType.DIVIDE_BY_ZERO
        elif isinstance(ex, KeyError):
            detail = f'Attempting to access unknown sheet'
            error_type = CellErrorType.BAD_REFERENCE
        else:
            detail = f'Unknown error occurred while performing {detail}'
        return Tree('cell_error', 
            [CellError(error_type, detail=detail, exception = ex)])
    

    ########################################################################
    # Helper Functions
    ######################################################################## 

    def normalize_number(self, num: Decimal) -> Decimal:
        '''
        Normalize a Decimal object such that we remove all leading and 
        trailing zeros, while not simplifying using exponent notation

        Arguments:
        - num: Decimal - decimal to be normalized

        Returns:
        - normalized Decimal

        '''

        norm = num.normalize()
        _, _, exp = norm.as_tuple()
        return norm if exp <= 0 else norm.quantize(1)
