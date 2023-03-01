'''
Function Handler

This module holds the basic functionality of dealing with function calls.

See the Evaluator module for implementation.

Global Variables:


Classes:
- FunctionHandler

    Methods:
    - map_func(object, str) -> Callable

'''

from typing import Callable, List
from lark import Tree

from .utils import convert_to_bool


class FunctionHandler:

    def __init__(self):

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
        return self._recognized_funcs[func_name]

    def __and(self, args: List) -> Tree:
        '''
        AND logic functionality

        Arguments:
        - args: List

        Returns:
        - Tree containing boolean result

        '''

        if len(args) < 1:
            raise TypeError('Invalid number of arguments')

        for expression in args:
            a = expression.children[0]
            res = convert_to_bool(a, type(a))
            if not res:
                return Tree('bool', [False])

        return Tree('bool', [True])
