'''
Evaluator

This module holds the functionality for evaluating a parsed formula based on the
described Lark grammar found at formulas.Lark.

See the Workbook, Sheet, and Cell modules for implementation.

Global Variables:
- COMP_OPERATORS (Dict[str, built_in_function]) - converts string of operator
    to the operator function
- EMPTY_SUBS (Dict[type, Any]) - converts type of not None expression to the
    correct empty value

Classes:
- Evaluator

    Attributes:
    - workbook (Workbook)
    - function_handler (FunctionHandler)

    Methods:
    - get_working_sheet(object) -> str
    - set_working_sheet(object, str) -> None
    - NUMBER(object, Token) -> Decimal
    - STRING(object, Token) -> str
    - BOOL(object, Token) -> bool
    - add_expr(object, List) -> Tree
    - mul_expr(object, List) -> Tree
    - unary_op(object, List) -> Tree
    - concat_expr(object, List) -> Tree
    - comp_expr(object, List) -> Tree
    - cell(object, List) -> Tree
    - func_expr(object, List) -> Tree
    - args_expr(object, List) -> Tree
    - parens(object, List) -> Tree
    - error(object, List) -> Tree

'''


import re
from decimal import Decimal, DecimalException, InvalidOperation
from typing import List, Tuple, Any
import operator

from lark import Tree, Transformer, Token

from .cell_error import CellError, CellErrorType, CELL_ERRORS
from .function_handler import FunctionHandler
from .utils import convert_to_bool


COMP_OPERATORS = {
    ">": operator.gt,
    "<": operator.lt,
    "<=": operator.le,
    ">=": operator.ge,
    "=": operator.eq,
    "==": operator.eq,
    "!=": operator.ne,
    "<>": operator.ne
}

EMPTY_SUBS = {
    str: '',
    Decimal: Decimal(0),
    bool: False
}

class Evaluator(Transformer):
    '''
    Evaluate an input string as a formula, based on the described Lark grammar
    in formulas.lark

    '''

    def __init__(self, workbook, sheet_name: str):
        '''
        Initialize an evaluator, as well as the Transformer

        Store the current workbook object as well as the current working
        sheet name

        '''

        Transformer.__init__(self)
        self.workbook = workbook
        self.function_handler = FunctionHandler()
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

    ########################################################################
    # Bases
    ########################################################################

    # pylint: disable=invalid-name

    # we disable snake case checkers here for the following bases, as they
    # must use uppercase to match the Lark grammar and all variables are
    # pre-existing and are thus previously checked

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

        return str(token)[1:-1]

    def BOOL(self, token: Token) -> bool:
        '''
        Evaluate a BOOL type

        Arguments:
        - token: Token - contains data about the boolean

        Returns:
        - boolean object

        '''

        return convert_to_bool(token, str)

    ########################################################################
    # Expressions
    ########################################################################

    # pylint: enable=invalid-name

    # we enable the checking for snake-cases again

    # pylint: disable=broad-exception-caught

    # We disable the checking for broad exceptions for the next few functions.
    # Although this is dangerous, we explicitly handle all exceptions in the
    # function __process_exceptions. Even if an unspecified exception occurs
    # in the following code, the message will be propagated through the result,
    # and the value of the formula will evaluate to a CellError with a None
    # error type

    def add_expr(self, args: List) -> Tree:
        '''
        Evaluate an addition/subtraction expression within the Tree

        Arguments:
        - args: List - list with Tree and/or Token objects of format
            [x, +/-, y]

        Returns:
        - Tree holding the result as a Decimal object

        '''

        try:
            oper = args[1]
            x = args[0].children[-1]
            y = args[-1].children[-1]

            # Check for propogating errors
            if isinstance(x, CellError):
                return Tree('cell_error', [x])
            if isinstance(y, CellError):
                return Tree('cell_error', [y])

            # Check for compatible types, deal with empty case
            # Decimal already accounts for passing in boolean values
            x = Decimal(0) if x is None else Decimal(x)
            y = Decimal(0) if y is None else Decimal(y)

            result = x + y if oper == '+' else x - y
            return Tree('number', [self.__normalize_number(Decimal(result))])

        except Exception as e:
            return self.__process_exceptions(e, detail='addition/subtraction')


    def mul_expr(self, args: List) -> Tree:
        '''
        Evaluate a multiplication/divison expression within the Tree

        Throw a ZeroDivisionError if trying to divide by 0

        Arguments:
        - args: List - list with Tree and/or Token objects of format
            [x, *//, y]

        Returns:
        - Tree holding the result as a Decimal object

        '''

        try:

            oper = args[1]
            x = args[0].children[-1]
            y = args[-1].children[-1]

            # Check for propogating errors
            if isinstance(x, CellError):
                return Tree('cell_error', [x])
            if isinstance(y, CellError):
                return Tree('cell_error', [y])

            # Check for compatible types, deal with empty cases
            # Decimal already accounts for passing in boolean values
            x = Decimal(0) if x is None else Decimal(x)
            y = Decimal(0) if y is None else Decimal(y)

            # Check for division by zero
            if y == Decimal(0) and oper == '/':
                raise ZeroDivisionError

            result = x * y if oper == '*' else x / y
            return Tree('number', [self.__normalize_number(Decimal(result))])

        except Exception as e:
            return self.__process_exceptions(e,
                                             detail='multiplication/division')

    def unary_op(self, args: List) -> Tree:
        '''
        Evaluate an unary operation expression within the Tree

        Arguments:
        - args: List - list with Tree and/or Token objects of format
            [+/-, x]

        Returns:
        - Tree holding the result as a Decimal object

        '''

        try:

            oper = args[0]
            x = args[1].children[-1]

            # Check for propogating errors
            if isinstance(x, CellError):
                return Tree('cell_error', [x])

            # Check for compatible types, deal with empty case
            # Decimal already accounts for passing in boolean values
            x = Decimal(0) if x is None else Decimal(x)

            result = -1 * x if oper == '-' else x
            return Tree('number', [Decimal(result)])

        except Exception as e:
            return self.__process_exceptions(e, detail='unary operations')

    def concat_expr(self, args: List) -> Tree:
        '''
        Evaluate a string concatenation expression within the Tree

        Arguments:
        - args: List - list with Tree and/or Token objects of format
            [s1, &, s2]

        Returns:
        - Tree holding the result as a String object

        '''

        try:
            str1 = args[0].children[-1]
            str2 = args[-1].children[-1]

            # Check for propogating errors
            if isinstance(str1, CellError):
                return Tree('cell_error', [str1])
            if isinstance(str2, CellError):
                return Tree('cell_error', [str2])

            # Handle when either input is a boolean
            if isinstance(str1, bool):
                str1 = str(str1).upper()
            if isinstance(str2, bool):
                str2 = str(str2).upper()

            # Check for compatible types, deal with empty case
            str1 = '' if str1 is None else str(str1)
            str2 = '' if str2 is None else str(str2)

            return Tree('string', [str1+str2])

        except Exception as e:
            return self.__process_exceptions(e, 'string concatenation')

    def comp_expr(self, args: List) -> Tree:
        '''
        Evaluate a comparison expression within the Tree

        Arguments:
        - args: List - List with Tree and/or Token objects of
            format [x, </<=/>/>=/=/==/!=/<>, y]

        Returns:
        - Tree holding the result as a boolean object

        '''

        try:
            oper = args[1]
            x = args[0].children[-1]
            y = args[-1].children[-1]
            x_type = type(x)
            y_type = type(y)

            # Check for propogating errors
            if isinstance(x, CellError):
                return Tree('cell_error', [x])
            if isinstance(y, CellError):
                return Tree('cell_error', [y])

            result = self.__compare_values(x, y, (x_type, y_type), oper)

            return Tree('bool', [result])

        except Exception as e:
            return self.__process_exceptions(e, 'boolean comparison')

    def cell(self, args: List) -> Tree:
        '''
        Evaluate a cell expression

        Throw a KeyError if the cell location in the cell_ref is out of bounds

        Arguments:
        - args: List - list of Tree and/or Token objects of format
            [sheet, cell]

        Returns:
        - Tree holding the resulting cell as a Cell object

        '''

        try:
            if len(args) == 2:
                working_sheet = args[0]
                if working_sheet[0] == "'":
                    working_sheet = working_sheet[1:-1]
                cell_name = args[1].replace('$', '')
            else:
                working_sheet = self.get_working_sheet()
                cell_name = args[0].replace('$', '')
            # Check that cell location is within bounds
            if not re.match(r"^[A-Z]{1,4}[1-9][0-9]{0,3}$", cell_name.upper()):
                raise KeyError('Cell location out of bounds')

            result = self.workbook.get_cell_value(working_sheet, cell_name)

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

    def func_expr(self, args: List) -> Tree:
        '''
        Evaluate a function expression

        Arguments:
        - args: List - function name and list of arguments

        Returns:
        - Tree holding result of the function

        '''

        try:
            func_name = args[0].lower()
            args_list = args[-1].children

            result = self.function_handler.map_func(func_name)
            return result(args_list)

        except KeyError as e:
            detail = 'function'
            return self.__process_exceptions(e, detail)

        except Exception as e:
            detail = "function operations"
            return self.__process_exceptions(e, detail)

    # pylint: enable=broad-exception-caught

    # We enable the checks for catching broad exceptions.

    def args_expr(self, args: List) -> Tree:
        '''
        Evaluate an expression of function arguments:

        '''

        if args == []:
            return Tree('args_list', [])

        if len(args[-1].children) == 1:
            return Tree('args_list', [args[0]]+[args[-1]])

        return Tree('args_list', [args[0]]+args[-1].children)

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
        e_type = [i[0] for i in list(CELL_ERRORS.items()) if i[-1]==x.upper()]
        e_type = CellErrorType(e_type[0])
        return Tree('cell_error', [CellError(e_type, '', None)])

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

        if isinstance(ex, (DecimalException, InvalidOperation)):
            detail = f'Attempting to perform {detail} on incompatible or '\
                'incorrect types'
            error_type = CellErrorType.TYPE_ERROR

        elif isinstance(ex, ZeroDivisionError):
            detail = 'Attempting to perform division with 0'
            error_type = CellErrorType.DIVIDE_BY_ZERO

        elif isinstance(ex, KeyError):
            error_type = CellErrorType.BAD_REFERENCE
            if detail == 'function':
                error_type = CellErrorType.BAD_NAME
            detail = f'Attempting to access unknown {detail}'

        elif isinstance(ex, TypeError):
            detail = f'{detail}: {ex.args[0]}'
            error_type = CellErrorType.TYPE_ERROR

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

    def __compare_values(self, left: Any, right: Any, types: Tuple[type, type],
                         oper: str) -> bool:
        '''
        Get the boolean value for a comparison between types of bool, str,
        and/or Decimal

        Arguments:
        - left: Any - left side of comparison
        - right: Any - right side of comparison
        - types: Tuple[type, type] - types of the left and right sides of the
            comparison operator
        - oper: str - comparison operator

        Returns:
        - boolean result of comparison

        '''

        result = False
        if types[0] == types[-1]:
            if types == (str, str):
                left = left.lower()
                right = right.lower()
            result = COMP_OPERATORS[oper](left, right)

        elif types in [(bool, str), (str, Decimal), (bool, Decimal)]:
            if oper in ['>', '>=', '!=', '<>']:
                result = True

        elif types in [(str, bool), (Decimal, str), (Decimal, bool)]:
            if oper in ['<', '<=', '!=', '<>']:
                result = True

        else:
            if left is not None:
                result = COMP_OPERATORS[oper](left, EMPTY_SUBS[types[0]])
            else:
                result = COMP_OPERATORS[oper](EMPTY_SUBS[types[-1]], right)

        return result
