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
    - test_or(object) -> None
    - test_not(object) -> None
    - test_xor(object) -> None
    - test_exact(object) -> None
    - test_if(object) -> None
    - test_iferror(object) -> None
    - test_choose(object) -> None
    - test_isblank(object) -> None
    - test_iserror(object) -> None
    - test_version(object) -> None
    - test_indirect(object) -> None
    - test_indirect2(object) -> None
    - test_common_math(object) -> None

'''

from decimal import Decimal

from lark import Lark, Tree

# pylint: disable=unused-import, import-error
import context
from sheets.evaluator import Evaluator
from sheets import Workbook, CellError, CellErrorType, version


WB = Workbook()
WB.new_sheet('Test')
EVALUATOR = Evaluator(WB, 'Test')
PARSER = Lark.open('../sheets/formulas.lark', start='formula', rel_to=__file__, parser='lalr')

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

        tree = PARSER.parse('=and("true", 7==7, A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=and("true")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=and(Test!A4)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=and("Test!A3")')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=and(False, #REF!)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

        WB.set_cell_contents('Test', 'A1', '\'tru')
        tree = PARSER.parse('=A1&"e"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['true'])

        tree = PARSER.parse('=AND(A1&"e")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=AND (A1&"e")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

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

        tree = PARSER.parse('=or(True, #REF!)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

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

        tree = PARSER.parse('=not(#REF!)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

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

        tree = PARSER.parse('=xor(True, #REF!)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

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

        tree = PARSER.parse('=EXACT(#REF!, #REF!)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

        WB.set_cell_contents('Test', 'A2', '=A3')
        WB.set_cell_contents('Test', 'A3', '=EXACT(#REF!, A2)')
        tree = PARSER.parse('=A3')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.CIRCULAR_REFERENCE

    def test_if(self) -> None:
        '''
        Test IF logic

        '''

        WB.set_cell_contents('Test', 'A1', '=True')
        WB.set_cell_contents('Test', 'A2', '=False')
        WB.set_cell_contents('Test', 'A3', '=0')

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

        tree = PARSER.parse('=IF(EXACT(7==8, A2), "string1", #REF!)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ["string1"])

        tree = PARSER.parse('=IF(AND("FaLSe", 7==8, A1, A2, A3), "0", A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal('0')])

        WB.set_cell_contents('Test', 'A1', '=A2+1')
        WB.set_cell_contents('Test', 'A2', '=IF(OR(True, 0), A1+1, A3+1)')
        WB.set_cell_contents('Test', 'A3', '=1+2')
        tree = PARSER.parse('=A2')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.CIRCULAR_REFERENCE

        WB.set_cell_contents('Test', 'A2', '=IF(AND(True, 0), A1+1, A3+1)')
        tree = PARSER.parse('=A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal('4')])

        WB.set_cell_contents('Test', 'A2', '=IF(AND(True, 0), A1+1)')
        tree = PARSER.parse('=A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [False])

        tree = PARSER.parse('=IF(#REF!, A1)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

        tree = PARSER.parse('=IF(True, Aaa122)')
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

        tree = PARSER.parse('=IFERROR("#REF!")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ["#REF!"])

        tree = PARSER.parse('=IFERROR(A2, #REF!)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [False])

        WB.set_cell_contents('Test', 'A1', '=A2+1')
        WB.set_cell_contents('Test', 'A2', '=IFERROR(A1+1, A3+1)')
        WB.set_cell_contents('Test', 'A3', '=1+2')
        tree = PARSER.parse('=A2')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.CIRCULAR_REFERENCE

        WB.set_cell_contents('Test', 'A2', '=IFERROR(A1+#REF!, A3+1)')
        tree = PARSER.parse('=A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal('4')])

        WB.set_cell_contents('Test', 'A2', '=IFERROR(A1+#REF!)')
        tree = PARSER.parse('=A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [""])

        tree = PARSER.parse('=IFERROR(ZZ201)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal('0')])

    def test_choose(self) -> None:
        '''
        Test CHOOSE logic

        '''

        WB.set_cell_contents('Test', 'A1', '2')
        WB.set_cell_contents('Test', 'A2', '=True')
        WB.set_cell_contents('Test', 'A3', '=0')

        tree = PARSER.parse('=CHOOSE(0, A1, 0)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=CHOOSE(1.5, A1, 0)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=CHOOSE(3, A1, 0)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=CHOOSE(A1, 1, 12)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('12')])

        tree = PARSER.parse('=CHOOSE(A2, "string1", #REF!)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ["string1"])

        tree = PARSER.parse('=CHOOSE(A3+1, A2, A1)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [True])

        WB.set_cell_contents('Test', 'A1', '=A2+1')
        WB.set_cell_contents('Test', 'A2', '=CHOOSE("1", A1+1, 2+1, A3+1)')
        WB.set_cell_contents('Test', 'A3', '=1+2')
        tree = PARSER.parse('=A2')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.CIRCULAR_REFERENCE

        WB.set_cell_contents('Test', 'A2', '=CHOOSE(2+1, A1+1, 2+1, A3+1)')
        tree = PARSER.parse('=A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal('4')])

        WB.set_cell_contents('Test', 'A1', '=A1')
        WB.set_cell_contents('Test', 'A2', '=CHOOSE(3, 0, A1)')
        tree = PARSER.parse('=A2')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=CHOOSE(#REF!, A1)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

        tree = PARSER.parse('=CHOOSE(1, Abcd1233)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal('0')])

    def test_isblank(self) -> None:
        '''
        Test ISBLANK logic

        '''

        WB.set_cell_contents('Test', 'A1', '')
        WB.set_cell_contents('Test', 'A2', '=False')
        WB.set_cell_contents('Test', 'A3', '=0')

        tree = PARSER.parse('=ISBLANK()')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=ISBLANK("string", A1)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=ISBLANK("")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=ISBLANK(A1)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=ISBLANK(A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=ISBLANK(A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        WB.set_cell_contents('Test', 'A3', '#REF!')
        tree = PARSER.parse('=ISBLANK(A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        WB.set_cell_contents('Test', 'A2', '=A3')
        WB.set_cell_contents('Test', 'A3', '=ISBLANK(A3)')
        tree = PARSER.parse('=A3')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.CIRCULAR_REFERENCE

    def test_iserror(self) -> None:
        '''
        Test ISERROR logic

        '''

        WB.set_cell_contents('Test', 'A1', '')
        WB.set_cell_contents('Test', 'A2', '=A1+')
        WB.set_cell_contents('Test', 'A3', '=1/0')

        tree = PARSER.parse('=ISERROR()')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=ISERROR("string", A1)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=ISERROR("A1+")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=ISERROR(A1)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=ISERROR(A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=ISERROR(A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        WB.set_cell_contents('Test', 'A1', '=A2')
        WB.set_cell_contents('Test', 'A2', '=A1')
        WB.set_cell_contents('Test', 'A3', '=ISERROR(A2)')
        tree = PARSER.parse('=ISERROR(A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=ISERROR(A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        WB.set_cell_contents('Test', 'A1', '=ISERROR(A2)')
        WB.set_cell_contents('Test', 'A2', '=ISERROR(A1)')
        WB.set_cell_contents('Test', 'A3', '=ISERROR(A2)')
        tree = PARSER.parse('=ISERROR(A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('=ISERROR(A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

    def test_version(self) -> None:
        '''
        test VERSION functionality

        '''

        tree = PARSER.parse('=VERSION()')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', [version])

        tree = PARSER.parse('=VERSION(arg1)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=VERSION("",arg1)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

    def test_indirect(self) -> None:
        '''
        test INDIRECT functionality

        '''

        tree = PARSER.parse('=INDIRECT()')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        tree = PARSER.parse('=INDIRECT(A1, A2)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        WB.set_cell_contents('Test', 'A1', '=1')
        tree = PARSER.parse('=INDIRECT(A1)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal(1)])

        tree = PARSER.parse('=INDIRECT("A1")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal(1)])

        WB.set_cell_contents('Test', 'A2', 'True')
        tree = PARSER.parse('=INDIRECT("Test!A2")')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [True])

        tree = PARSER.parse('=INDIRECT(Test!A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [True])

        tree = PARSER.parse('=INDIRECT(A4)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

        tree = PARSER.parse('=INDIRECT("A5")')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

        tree = PARSER.parse('=INDIRECT(Sheet2!A1)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

        tree = PARSER.parse('=INDIRECT("Sheet2!A1")')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

        tree = PARSER.parse('=INDIRECT("Sheet2!!A1")')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

    def test_indirect2(self) -> None:
        '''
        test INDIRECT functionality - Part 2

        '''

        tree = PARSER.parse('=INDIRECT(123)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

        tree = PARSER.parse('=INDIRECT(True)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

        tree = PARSER.parse('=INDIRECT(AND(1))')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.BAD_REFERENCE

    def test_common_math(self) -> None:
        '''
        test MIN, MAX, SUM, AVG functionality

        '''

        WB.set_cell_contents('Test', 'A1', '=0')
        WB.set_cell_contents('Test', 'A2', '=-100')
        WB.set_cell_contents('Test', 'B2', '=130')
        WB.set_cell_contents('Test', 'D1', '=True')
        WB.set_cell_contents('Test', 'D2', '="string"')
        tree = PARSER.parse('=MIN(1, 3, A1:B2)')
        result = EVALUATOR.transform(tree).children[-1]
        assert result == Decimal(-100)

        tree = PARSER.parse('=MAX(1, 3, A1:B2)')
        result = EVALUATOR.transform(tree).children[-1]
        assert result == Decimal(130)

        tree = PARSER.parse('=SUM(1, 3, A1:B2)')
        result = EVALUATOR.transform(tree).children[-1]
        assert result == Decimal(34)

        tree = PARSER.parse('=AVERAGE(1, 3, A1:B2)')
        result = EVALUATOR.transform(tree).children[-1]
        assert result == Decimal('6.8')

        tree = PARSER.parse('=SUM(C1:C2)')
        result = EVALUATOR.transform(tree).children[-1]
        assert result == Decimal(0)

        tree = PARSER.parse('=MAX(20, D1:D2)')
        result = EVALUATOR.transform(tree).children[-1]
        assert isinstance(result, CellError)
        assert result.get_type() == CellErrorType.TYPE_ERROR

        WB.set_cell_contents('Test', 'B1', '=True')
        WB.set_cell_contents('Test', 'C1', '=1')
        WB.set_cell_contents('Test', 'C2', '=1')
        WB.set_cell_contents('Test', 'C3', '=1')
        WB.set_cell_contents('Test', 'C4', '=1')
        WB.set_cell_contents('Test', 'C5', '=1')

        tree = PARSER.parse('=SUM(IF(B1, C1:C5, D1:D10))')
        result = EVALUATOR.transform(tree).children[-1]
        assert result == Decimal(5)

    def test_lookups(self) -> None:
        '''
        test HLOOKUP and VLOOKUP functionality

        '''

        WB.set_cell_contents('Test', 'A1', '=0')
        WB.set_cell_contents('Test', 'A2', '=-100')
        WB.set_cell_contents('Test', 'B1', '="sparkles"')
        WB.set_cell_contents('Test', 'B2', '=130')
        WB.set_cell_contents('Test', 'C1', '="sparkles"')
        WB.set_cell_contents('Test', 'C2', '=134')
        WB.set_cell_contents('Test', 'D1', '=True')
        WB.set_cell_contents('Test', 'D2', '="string"')

        tree = PARSER.parse('=HLOOKUP("sparkles", A1:D2, 2)')
        result = EVALUATOR.transform(tree).children[-1]
        assert result == Decimal(130)

        tree = PARSER.parse('=VLOOKUP(0, D2:A1, 2)')
        result = EVALUATOR.transform(tree).children[-1]
        assert result == "sparkles"
