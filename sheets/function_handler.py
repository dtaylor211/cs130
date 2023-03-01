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


from lark import Tree

from .utils import convert_to_bool


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
            raise TypeError('Invalid number of arguments')
        print('d',args.children)
        for expression in args.children:
            print(expression)
            a = expression.children[0]
            print('a', a, type(a))
            res = convert_to_bool(a, type(a))
            print(res)
            if not res:
                print('ay')
                return Tree('bool', [False])

        return Tree('bool', [True])

    

