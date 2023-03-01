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


class FunctionHandler:
    '''
    Handles internal function calls

    '''

    def __init__(self):
        '''
        Initialize a new function handler

        '''

        self._recognized_funcs = {
            'and': self.__and
            # 'or': self.__or,
            # 'not': self.__not,
            # 'xor': self.__xor,
            # 'exact': self.__exact,
            # 'if': self.__if,
            # 'iferror': self.__iferror,
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

        for expression in args:
            arg = expression.children[0]
            res = convert_to_bool(arg, type(arg))
            if not res:
                return Tree('bool', [False])

        return Tree('bool', [True])
