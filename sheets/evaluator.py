import re
from lark import Tree, Transformer, Lark, Token
from decimal import Decimal, DecimalException, InvalidOperation
from typing import *

from .cell_error import CellError, CellErrorType, CELL_ERRORS


class Evaluator(Transformer):
    '''
    Evaluate an input string as a formula, based on the described Lark grammar
    in ./formulas.lark

    '''

    def __init__(self, workbook, sheet_name: str):
        '''
        Initialize an evaluator, as well as the Transformer

        Store the current workbook object as well as the current working
        sheet name

        '''

        Transformer.__init__(self)
        self._workbook = workbook
        self._working_sheet = sheet_name

    ########################################################################
    # Getters and Setters
    ########################################################################

    def get_working_sheet(self) -> str:
        '''
        Get the name of the current working sheet

        Returns:
        - string sheet name

        '''

        return self._working_sheet
    
    def set_working_sheet(self, sheet_name: str) -> None:
        '''
        Set the name of the current working sheet

        Arguments:
        - sheet_name: str - string sheet name

        '''

        self._working_sheet = sheet_name

    def get_workbook(self) -> None:
        '''
        Set the name of the current working sheet

        Arguments:
        - sheet_name: str - string sheet name

        '''

        return self._workbook

    ########################################################################
    # Bases
    ########################################################################

    def NUMBER(self, token: Token) -> Decimal:
        '''
        Evaluate a NUMBER type into a Decimal object   
        This method also normalizes the value stored in Decimal

        Arguments:
        - token: Token - contains data about the number

        Returns:
        - Decimal object storing the given number

        '''

        return self.__normalize_number(Decimal(token))

    def STRING(self, token: Token) -> str:
        '''
        Evaluate a STRING type  

        Arguments:
        - token: Token - contains data about the string

        Returns:
        - String object with beginning and ending quotations removed
        
        '''

        return (str(token)[1:-1])
    
    def CELLREF(self, token: Token) -> Token:
        '''
        Evaluate a CELLREF type

        Arguments:
        - token: Token - contains data about the cell reference

        Returns:
        - the original token passed in

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

            # Check for compatible types, deal with empty case
            x = Decimal(0) if x is None else Decimal(x)
            y = Decimal(0) if y is None else Decimal(y)

            result = x + y if operator == '+' else x - y
            return Tree('number', [self.__normalize_number(Decimal(result))])

        except Exception as e:
            return self.__process_exceptions(e, detail='addition/subtraction')

    def mul_expr(self, args: List) -> Tree:
        '''
        Evaluate a multiplication/divison expression within the Tree

        Throw a ZeroDivisionError if trying to divide by 0

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

            # Check for compatible types, deal with empty cases
            x = Decimal(0) if x is None else Decimal(x)
            y = Decimal(0) if y is None else Decimal(y)

            # Check for division by zero
            if y == Decimal(0) and operator == '/':
                raise ZeroDivisionError
            
            result = x * y if operator == '*' else x / y
            return Tree('number', [self.__normalize_number(Decimal(result))])

        except Exception as e:
            return self.__process_exceptions(e, detail='multiplication/division')    

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

            # Check for propogating errors
            if isinstance(x, CellError):
                return Tree('cell_error', [x])

            # Check for compatible types, deal with empty case
            x = Decimal(0) if x is None else Decimal(x)

            result = -1 * x if operator == '-' else x
            return Tree('number', [Decimal(result)])

        except Exception as e:
            return self.__process_exceptions(e, detail='unary operations')

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
            return self.__process_exceptions(e, 'string concatenation')

    def cell(self, args: List) -> Tree:
        '''
        Evaluate a cell expression

        Throw a KeyError if the cell location in the cell_ref is out of bounds

        Arguments:
        - args: List - list of Tree and/or Token objects of format [sheet, cell]

        Returns:
        - Tree holding the resulting cell as a Cell object

        '''

        try:
            if len(args) == 2:
                working_sheet = args[0]
                if working_sheet[0] == "'":
                    working_sheet = working_sheet[1:-1]
                cell_name = args[1]
            else:
                working_sheet = self.get_working_sheet()
                cell_name = args[0]

            # Check that cell location is within bounds
            if not re.match(r"^[A-Z]{1,4}[1-9][0-9]{0,3}$", cell_name.upper()):
                raise KeyError('Cell location out of bounds')
            
            workbook = self.get_workbook()
            result = workbook.get_cell_value(working_sheet, cell_name)

            # Check for propogating errors
            if isinstance(result, CellError):
                return Tree('cell_error', [result])

            # Deal with empty cases
            return Tree('cell_ref', [result])
        
        except KeyError as k:
            if k.args[0] == 'Cell location out of bounds':
                detail = 'cell'
            else: detail = 'sheet'
            return self.__process_exceptions(k, detail)

        except Exception as e:
            return self.__process_exceptions(e, 'cell operations')
        
    def parens(self, args: List) -> Tree:
        '''
        Evaluate an expression enclosed in parenthesis

        Arguments:
        - args: List - expression enclosed in brackets [expr]

        Returns:
        - Tree holding the expression inside the parenthesis

        '''

        return args[0] # removes the outer brackets
    
    def error(self, args: List) -> Tree:
        '''
        Evaluate an error expression

        Arguments:
        - args: List - error Tree enclosed in brackets [err]

        Returns:
        - Tree holding the error, now converted to a CellError type

        '''

        x = args[0]
        e_type = [i for i in CELL_ERRORS if CELL_ERRORS[i]==x.upper()]
        c = CellErrorType(e_type[0])
        return Tree('cell_error', [CellError(c, '', None)])

    ########################################################################
    # Exception Processing
    ########################################################################       

    def __process_exceptions(self, ex: Exception, detail: str = '') -> Tree:
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

        if isinstance(ex, DecimalException) or isinstance(ex, InvalidOperation):
            detail = f'Attempting to perform {detail} on incompatible or '\
                'incorrect types'
            error_type = CellErrorType.TYPE_ERROR

        elif isinstance(ex, ZeroDivisionError):
            detail = 'Attempting to perform division with 0'
            error_type = CellErrorType.DIVIDE_BY_ZERO

        elif isinstance(ex, KeyError):
            detail = f'Attempting to access unknown {detail}'
            error_type = CellErrorType.BAD_REFERENCE

        else:
            detail = f'Unknown error occurred while performing {detail}'

        return Tree('cell_error', 
            [CellError(error_type, detail=detail, exception = ex)])
    
    ########################################################################
    # Helper Functions
    ######################################################################## 

    def __normalize_number(self, num: Decimal) -> Decimal:
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
