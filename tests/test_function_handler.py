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
    - test_and(object) -> None

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

        tree = PARSER.parse('=AND(0, "string")')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=AND(True, 4)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=AND("true", 7==7, A1, A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=AND("true", 7==7, A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

    def test_or(self) -> None:
        '''
        Test OR logic

        '''

        WB.set_cell_contents('Test', 'A1', '=True')
        WB.set_cell_contents('Test', 'A2', '=False')
        WB.set_cell_contents('Test', 'A3', '=1')

        tree = PARSER.parse('=OR()')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=OR(1, "string")')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=OR(False, 4)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=OR("false", 7==8, A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=OR("FaLSe", 7==8, A2, AND(A1, A3))')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

    def test_not(self) -> None:
        '''
        Test NOT logic

        '''

        WB.set_cell_contents('Test', 'A1', '=True')
        WB.set_cell_contents('Test', 'A2', '=False')
        WB.set_cell_contents('Test', 'A3', '=0')

        tree = PARSER.parse('=NOT()')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=NOT("string")')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=NOT(False, 4)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=NOT(False)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=NOT(7==7)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=NOT(AND("FaLSe", 7==8, A1, A2, A3))')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

    def test_xor(self) -> None:
        '''
        Test XOR logic

        '''

        WB.set_cell_contents('Test', 'A1', '=True')
        WB.set_cell_contents('Test', 'A2', '=False')
        WB.set_cell_contents('Test', 'A3', '=1')

        tree = PARSER.parse('=XOR()')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=XOR(1, "string")')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=XOR(False, 4)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=XOR("false", 7==8, A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=XOR("FaLSe", 7==8, A2, AND(A1, A3))')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=XOR("FaLSe", 7==7, A2, AND(A1, A3))')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=XOR("tRUe", 7==7, NOT(A2), XOR(A1, A3))')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

    def test_exact(self) -> None:
        '''
        Test EXACT logic

        '''

        WB.set_cell_contents('Test', 'A1', '=True')
        WB.set_cell_contents('Test', 'A2', '')
        WB.set_cell_contents('Test', 'A3', '=0')

        tree = PARSER.parse('=EXACT()')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=EXACT("string")')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=EXACT(False, 4, 10)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=EXACT(False, "FALSE")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=EXACT(A1, "True")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=EXACT(A2, "")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=EXACT(A3, "0")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

    def test_if(self) -> None:
        '''
        Test IF logic

        '''

        WB.set_cell_contents('Test', 'A1', '=True')
        WB.set_cell_contents('Test', 'A2', '=False')
        WB.set_cell_contents('Test', 'A3', '=0')

        tree = PARSER.parse('=IF()')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=IF("string", A1)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=IF(False)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=iF(False, True, "false", 12)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=IF(False, 12)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=IF(7==7, 12)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('12')])

        tree = PARSER.parse('=IF(EXACT(7==8, A2), "string1", 0)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ["string1"])

        tree = PARSER.parse('=IF(AND("FaLSe", 7==8, A1, A2, A3), "0", A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal('0')])

    def test_iferror(self) -> None:
        '''
        Test IFERROR logic

        '''

        WB.set_cell_contents('Test', 'A1', '=IFERROR()')
        WB.set_cell_contents('Test', 'A2', '=False')
        WB.set_cell_contents('Test', 'A3', '=0')

        tree = PARSER.parse('=IFERROR()')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=IFERROR("string", A1, 0)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=IFERROR(A1)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', [""])

        tree = PARSER.parse('=IFERROR(A1, 12)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('12')])

        tree = PARSER.parse('=IFERROR(A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [False])

        tree = PARSER.parse('=IFERROR(A2, A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [False])
