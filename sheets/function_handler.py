'''
Function Handler

This module holds the basic functionality of dealing with function calls.

See the Evaluator module for implementation.

Classes:
- FunctionHandler

    Methods:
    - map_func(object, str) -> Callable

'''

from typing import Callable, List
from lark import Tree

from .utils import convert_to_bool
from .cell_error import CellError


class FunctionHandler:
    '''
    Handles internal function calls

    '''

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
            # 'choose': self.__choose,
            # 'isblank': self.__isblank,
            # 'iserror': self.__iserror,
            # 'version': self.__version,
            # 'indirect': self.__indirect
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
            res = convert_to_bool(arg, type(arg))
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

        bool_result = True
        for expression in args:
            arg = expression.children[0]
            res = convert_to_bool(arg, type(arg))
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

        str1 = args[0].children[0]
        str2 = args[1].children[0]

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
        res = convert_to_bool(arg, type(arg))
        if res:
            return args[1]
        if len(args) == 3:
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
            return args[0]
        if len(args) == 2:
            return args[1]

        return Tree('string', [""])
