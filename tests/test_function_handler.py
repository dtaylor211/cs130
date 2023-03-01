'''
Test Function Handler

Tests the Function Handler module found at ../sheets/function_handler.py. 
Focuses on functionality verificiation of the supported function names.

GLOBAL_VARIABLES:
- WB (Workbook) - the Workbook used for this test suite
- EVALUATOR (Evaluator) - the Evaluator used for this test suite
- PARSER (Any) - the Parser used for this test suite

Classes:
- TestFunctionHandler

    Methods:
    - test_num_literals(object) -> None

'''

from decimal import Decimal

from lark import Lark, Tree
import pytest

# pylint: disable=unused-import, import-error
import context
from sheets.evaluator import Evaluator
from sheets import Workbook, CellError, CellErrorType


WB = Workbook()
WB.new_sheet('Test')
EVALUATOR = Evaluator(WB, 'Test')
PARSER = Lark.open('../sheets/formulas.lark', start='formula', rel_to=__file__)

class TestFunctionHandler:
    '''
    Tests the Function Handler and internal function supports

    '''

    def test_and(self) -> None:
        '''
        Test AND logic

        '''

        WB.set_cell_contents('Test', 'A1', '=True')
        WB.set_cell_contents('Test', 'A2', '=False')
        WB.set_cell_contents('Test', 'A3', '=1')

        tree = PARSER.parse('=AND()')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=AND(1, "string")')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR
        
        tree = PARSER.parse('=AND(True, 4)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=AND("true", 7==7, A1, A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])
        print('start')
        tree = PARSER.parse('=AND("true", 7==7, A2)')
        print(tree)
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])
