'''
Function Handler

This module holds the basic functionality of dealing with function calls.

See the Evaluator module for implementation.

Classes:
- FunctionHandler

    Methods:
    - map_func(object, str) -> Callable

'''

from typing import Callable, List, Tuple, Optional
from decimal import Decimal, InvalidOperation

import lark
from lark import Tree, Lark

# import sheets
from .utils import convert_to_bool, get_loc_from_coords
from .cell_error import CellError, CellErrorType
from .configs import VERSION



# pylint: disable=too-few-public-methods

class FunctionHandler:
    '''
    Handles internal function calls

    '''

    PARSER = Lark.open('formulas.lark', start='formula',
                rel_to=__file__, parser='lalr')

    def __init__(self):
        '''
        Initialize a new function handler

        '''

        self._recognized_funcs = {
            'and': self.__and,
            'or': self.__or,
            'not': self.__not,
            'xor': self.__xor,
            'exact': self.__exact,
            'if': self.__if,
            'iferror': self.__iferror,
            'choose': self.__choose,
            'isblank': self.__isblank,
            'iserror': self.__iserror,
            'version': self.__version,
            'indirect': self.__indirect,
            'min': self.__min,
            'max': self.__max,
            'sum': self.__sum,
            'average': self.__average,
            'hlookup': self.__hlookup,
            'vlookup': self.__vlookup
        }

    def map_func(self, func_name: str) -> Callable:
        '''
        Map a function name to its Callable counterpart

        Arguments:
        - func_name: str - name of desired function

        Returns:
        - Callable function with name func_name

        '''

        return self._recognized_funcs[func_name]

    def __and(self, args: List) -> Tree:
        '''
        AND logic functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing boolean result

        '''

        if len(args) < 1:
            raise TypeError('Invalid number of arguments')

        bool_result = True
        for expression in args:
            arg = expression.children[0]
            if isinstance(arg, CellError):
                return Tree('cell_error', [arg])
            res = convert_to_bool(arg, type(arg)) if arg is not None else False
            bool_result = bool_result and res

        return Tree('bool', [bool_result])

    def __or(self, args: List) -> Tree:
        '''
        OR logic functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing boolean result

        '''

        if len(args) < 1:
            raise TypeError('Invalid number of arguments')

        bool_result = False
        for expression in args:
            arg = expression.children[0]
            if isinstance(arg, CellError):
                return Tree('cell_error', [arg])
            res = convert_to_bool(arg, type(arg))
            bool_result = bool_result or res

        return Tree('bool', [bool_result])

    def __not(self, args: List) -> Tree:
        '''
        NOT logic functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing boolean result

        '''

        if len(args) != 1:
            raise TypeError('Invalid number of arguments')

        arg = args[0].children[0]
        if isinstance(arg, CellError):
            return Tree('cell_error', [arg])
        res = convert_to_bool(arg, type(arg))
        bool_result = not res

        return Tree('bool', [bool_result])

    def __xor(self, args: List) -> Tree:
        '''
        XOR logic functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing boolean result

        '''

        if len(args) < 1:
            raise TypeError('Invalid number of arguments')

        bool_result = False
        for expression in args:
            arg = expression.children[0]
            if isinstance(arg, CellError):
                return Tree('cell_error', [arg])
            res = convert_to_bool(arg, type(arg))
            # only flips when res true, will be true on odd number of true args
            bool_result = bool_result != res

        return Tree('bool', [bool_result])

    def __exact(self, args: List) -> Tree:
        '''
        EXACT logic functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing boolean result

        '''

        if len(args) != 2:
            raise TypeError('Invalid number of arguments')

        # pylint: disable=duplicate-code

        # we have same code called once more in evaluator, but it would not
        # really make sense to have a helper as this is a specific case that
        # only happens twice
        str1 = args[0].children[0]
        str2 = args[1].children[0]

        if isinstance(str1, CellError):
            return Tree('cell_error', [str1])
        if isinstance(str2, CellError):
            return Tree('cell_error', [str2])

        if isinstance(str1, bool):
            str1 = str(str1).upper()
        if isinstance(str2, bool):
            str2 = str(str2).upper()

        # Check for compatible types, deal with empty case
        str1 = '' if str1 is None else str(str1)
        str2 = '' if str2 is None else str(str2)

        bool_result = str1 == str2

        return Tree('bool', [bool_result])

    def __if(self, args: List) -> Tree:
        '''
        IF logic functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing result value

        '''

        if len(args) != 2 and len(args) != 3:
            raise TypeError('Invalid number of arguments')

        arg = args[0].children[0]
        if isinstance(arg, CellError):
            return Tree('cell_error', [arg])
        res = convert_to_bool(arg, type(arg))

        if res:
            if args[1].children[-1] is None:
                args[1].children[-1] = Decimal('0')
            return args[1]
        if len(args) == 3:
            if args[2].children[-1] is None:
                args[2].children[-1] = Decimal('0')
            return args[2]

        return Tree('bool', [False])

    def __iferror(self, args: List) -> Tree:
        '''
        IFERROR logic functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing result value

        '''

        if len(args) != 1 and len(args) != 2:
            raise TypeError('Invalid number of arguments')

        arg = args[0].children[0]
        if not isinstance(arg, CellError):
            if args[0].children[-1] is None:
                args[0].children[-1] = Decimal('0')
            return args[0]
        if len(args) == 2:
            if args[1].children[-1] is None:
                args[1].children[-1] = Decimal('0')
            return args[1]

        return Tree('string', [""])

    def __choose(self, args: List) -> Tree:
        '''
        CHOOSE logic functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing result value

        '''

        if len(args) < 2:
            raise TypeError('Invalid number of arguments')

        arg = args[0].children[0]
        if isinstance(arg, CellError):
            return Tree('cell_error', [arg])
        try:
            arg = Decimal(0) if arg is None else Decimal(arg)
            if arg % 1 != 0 or arg < 1 or arg > len(args)-1:
                raise TypeError('Invalid CHOOSE index')
        except InvalidOperation as e:
            raise TypeError('Invalid CHOOSE index') from e

        if args[int(arg)].children[-1] is None:
            args[int(arg)].children[-1] = Decimal('0')
        return args[int(arg)]

    def __isblank(self, args: List) -> Tree:
        '''
        ISBLANK logic functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing boolean result

        '''

        if len(args) != 1:
            raise TypeError('Invalid number of arguments')

        arg = args[0].children[0]
        if arg is None:
            return Tree('bool', [True])

        return Tree('bool', [False])

    def __iserror(self, args: List) -> Tree:
        '''
        ISERROR logic functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing boolean result

        '''

        if len(args) != 1:
            raise TypeError('Invalid number of arguments')

        arg = args[0].children[0]
        if isinstance(arg, CellError):
            return Tree('bool', [True])

        return Tree('bool', [False])

    def __version(self, args: List) -> Tree:
        '''
        VERSION functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing result value

        '''

        if len(args) != 0:
            raise TypeError('Invalid number of arguments')

        return Tree('string', [f"{VERSION}"])

    def __indirect(self, args: List) -> Tuple[Tree or str, Optional[str]]:
        '''
        INDIRECT functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing result value

        '''

        if len(args) != 1:
            raise TypeError('Invalid number of arguments')

        if args[0].data == 'cell_ref':
            return args[0]

        if args[0].data == 'cell_error':
            return args[0]

        try:
            tree = self.PARSER.parse(f'={str(args[0].children[-1])}')
            if tree.data != 'cell':
                raise lark.exceptions.LarkError
            return tree, 'Y'

        except lark.exceptions.LarkError:
            return Tree('cell_error', [
                CellError(CellErrorType.BAD_REFERENCE, '')])

    def __min(self, args: List) -> Tree:
        '''
        MIN functionality

        Arguments:
        - args: List - list of arguments to given function

        Returns:
        - Tree containing Decimal result

        '''

        all_values = self.__populate_all_range_values(args)

        min_val = min(all_values) if all_values else Decimal(0)
        return Tree('number', [min_val])

    def __max(self, args: List) -> Tree:
        '''
        MAX functionality

        Arguments:
        - args: List - list of arguments for given function

        Returns:
        - Tree with Decimal result

        '''

        all_values = self.__populate_all_range_values(args)

        max_val = max(all_values) if all_values else Decimal(0)
        return Tree('number', [max_val])

    def __sum(self, args: List) -> Tree:
        '''
        SUM functionality
        
        Arguments:
        - args: List - list of arguments for given function

        Returns:
        - Tree with Decimal result

        '''

        all_values = self.__populate_all_range_values(args)

        sum_val = sum(all_values) if all_values else Decimal(0)
        return Tree('number', [sum_val])

    def __average(self, args: List) -> Tree:
        '''
        AVERAGE functionality
        
        Arguments:
        - args: List - list of arguments for given function
        
        Returns:
        - Tree containing Decimal result

        '''

        all_values = self.__populate_all_range_values(args)

        if not all_values:
            raise ZeroDivisionError('All given values are empty')

        avg_val = sum(all_values) / len(all_values)
        return Tree('number', [avg_val])

    def __hlookup(self, args: List) -> Tree:
        '''
        HLOOKUP functionality
        
        Arguments:
        - args: List - list of arguments for given function
        
        Returns:
        - Tree containing cell reference result

        '''

        key = args[0].children[-1]
        tl_br_corners = args[1].children[-1]['tl_br_corners']
        ran = args[1].children[-1]['cells']
        index = args[-1].children[-1]

        if index < 1:
            raise TypeError(f'invalid index given: {index}')
        curr_col, curr_row = tl_br_corners[0]
        max_col, max_row = tl_br_corners[-1]

        while ran[get_loc_from_coords((curr_col, curr_row))] != key:
            if curr_col > max_col:
                raise TypeError('No column found')
            curr_col += 1

        if curr_row + index - 1 > max_row:
            raise TypeError(f'invalid index given: {index}')
        loc = get_loc_from_coords((curr_col, curr_row+index-1))
        res = [ran[loc]]

        return Tree('cell_ref', res)

    def __vlookup(self, args: List) -> Tree:
        '''
        VLOOKUP functionality
        
        Arguments:
        - args: List - list of arguments for given function
        
        Returns:
        - Tree containing cell reference result

        '''
        key = args[0].children[-1]
        tl_br_corners = args[1].children[-1]['tl_br_corners']
        ran = args[1].children[-1]['cells']
        index = args[-1].children[-1]

        if index < 1:
            raise TypeError(f'invalid index given: {index}')
        curr_col, curr_row = tl_br_corners[0]
        max_col, max_row = tl_br_corners[-1]

        while ran[get_loc_from_coords((curr_col, curr_row))] != key:
            if curr_row > max_row:
                raise TypeError('No row found')
            curr_row += 1

        if curr_col + index - 1 > max_col:
            raise TypeError(f'invalid index given: {index}')
        loc = get_loc_from_coords((curr_col+index-1, curr_row))
        res = [ran[loc]]

        return Tree('cell_ref', res)

    ########################################################################
    # Helper Functions
    ########################################################################

    def __populate_all_range_values(self, args: List) -> List:
        '''
        Populate a list of all values in a range of cells

        Arguments:
        - args: List - list of arguments for a given common math function

        Returns:
        - List of relevant, non-empty values, converted to Decimals

        '''

        all_values = []
        if len(args) < 1:
            raise TypeError('Invalid number of arguments')

        for arg in args:
            try:
                if arg.data == 'cell_range':
                    for cell_val in list(arg.children[-1]['cells'].values()):
                        if cell_val is None:
                            continue
                        all_values.append(Decimal(cell_val))
                else:
                    all_values.append(arg.children[-1])

            except InvalidOperation as e:
                raise TypeError(f'Cannot convert argument {arg} to Decimal') \
                    from e
        return all_values
