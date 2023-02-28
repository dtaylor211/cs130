'''
Function Handler

This module holds the basic functionality of dealing with function calls.

See the Evaluator module for implementation.

Global Variables:


Classes:
- FunctionHandler

    Methods:
    - and

'''

from decimal import Decimal
from lark import Tree


class FunctionHandler:

    def __init__(self):

        self.recognized_funcs = {
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

    def map_func(self, func_name: str):
        return self.recognized_funcs[func_name]
    
    def convert_to_bool(self, input, input_type: type) -> bool:
        '''
        '''
        
        result = None
        if input_type == bool:
            return input
        elif input_type == str:
            if input.lower() == 'true':
                result = True
            elif input.lower() == 'false':
                result = False
            else:
                raise TypeError('Cannot convert given string to boolean')
        else:
            if input != Decimal(0):
                result = True
            else:
                result = False
        return result
    
    #  def __convert_from_bool(bool_input: bool, target_type: type) -> Any:
    #     '''
    #     '''

    #     result = None
    #     if bool_input:
    #         if target_type == Decimal:
    #             result = Decimal(1)
    #         else:
    #             result = 'TRUE'
    #     else:
    #         if target_type == Decimal:
    #             result = Decimal(0)
    #         else:
    #             result = 'FALSE'
    #     return result
    
    def __and(self, args):
        '''
        '''

        if len(args.children) < 1:
            raise TypeError
        
        for expression in args.children:
            a = expression.children[0]
            res = self.convert_to_bool(a, type(a))
            if res != True:
                return Tree('bool', [False])
        return Tree('bool', [True])

    

