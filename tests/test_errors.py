'''
Test Errors

Tests the Evaluator and Workbook modules found in ../sheets/ with invalid
inputs.
Also tests for error propogation.

GLOBAL_VARIABLES:
- WB (Workbook) - the Workbook used for this test suite

Classes:
- TestErrors

    Methods:
    - test_parse_error(object) -> None
    - test_circular_reference(object) -> None
    - test_bad_reference(object) -> None
    - test_bad_name(object) -> None
    - test_type_error(object) -> None
    - test_divide_by_zero(object) -> None
    - test_error_as_literal(object) -> None
    - test_error_in_complex_literal(object) -> None
    - test_reference_cell_with_error(object) -> None
    - test_error_ordering(object) -> None

'''


# pylint: disable=unused-import, import-error
import context
from sheets.workbook import Workbook
from sheets.cell_error import CellError, CellErrorType


WB = Workbook()
WB.new_sheet('Test')

class TestErrors:
    '''
    Tests the formula parser and evaluator using invalid inputs

    '''

    def test_parse_error(self) -> None:
        '''
        Test when given a formula with a parse error

        '''

        WB.set_cell_contents('Test', 'A1', '=1E+4')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=1E+4'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A1', '=A1A2')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=A1A2'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A1', '=1**2')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=1**2'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A1', '=A1((A2)')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=A1((A2)'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A1', '=A1+(A2-A1')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=A1+(A2-A1'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A1', '=A1+A2-A1)')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=A1+A2-A1)'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A1', '=A 1')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=A 1'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A1', 'Sheet2')
        WB.set_cell_contents('Test', 'A2', '=Test!A1 & \'Sheet2\'')
        contents = WB.get_cell_contents('Test', 'A2')
        result_value = WB.get_cell_value('Test', 'A2')
        assert contents == '=Test!A1 & \'Sheet2\''
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A3', '=VERSION(')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=VERSION('
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A3', '=any_func(arg,)')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=any_func(arg,)'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A3', '=__invalid_func(arg)')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=__invalid_func(arg)'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A3', '=invalid?_func(arg)')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=invalid?_func(arg)'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A3', '=valid_func(arg')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=valid_func(arg'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A3', '=INDIRECT(Sheet2!!A1)')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=INDIRECT(Sheet2!!A1)'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.PARSE_ERROR

    def test_circular_reference(self) -> None:
        '''
        Test when given a formula with a circular reference error

        '''

        WB.set_cell_contents('Test', 'A1', '12')
        WB.set_cell_contents('Test', 'A2', '13')
        WB.set_cell_contents('Test', 'A3', '=A1+A2')
        WB.set_cell_contents('Test', 'A4', '=A1+A2+A3')

        WB.set_cell_contents('Test', 'A1', '=A4')
        result_contents = WB.get_cell_contents('Test','A4')
        result_value = WB.get_cell_value('Test', 'A4')
        assert result_contents == '=A1+A2+A3'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE

        WB.set_cell_contents('Test', 'A1', '=A1')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=A1'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE

        WB.set_cell_contents('Test', 'A1', '=1')
        WB.set_cell_contents('Test', 'B1', '=A1')
        WB.set_cell_contents('Test', 'A1', '=B1')
        value = WB.get_cell_value('Test', 'A1')
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE

        WB.set_cell_contents('Test', 'A1', '=B1')
        WB.set_cell_contents('Test', 'B1', '=INDIRECT("Test!A1")')
        value = WB.get_cell_value('Test', 'B1')
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE

        WB.set_cell_contents('Test', 'A1', '=B1')
        WB.set_cell_contents('Test', 'B1', '=INDIRECT(Test!B1)')
        value = WB.get_cell_value('Test', 'B1')
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE

    def test_bad_reference(self) -> None:
        '''
        Test when given a formula with a bad reference error

        '''

        WB.set_cell_contents('Test', 'A1', '=Test2!A1')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=Test2!A1'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE

        WB.set_cell_contents('Test', 'A1', '=AAAAA1')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=AAAAA1'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE

        WB.set_cell_contents('Test', 'A1', '=AAAA10000')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=AAAA10000'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE

        WB.new_sheet('Del')
        WB.set_cell_contents('Del', 'A1', '1')
        WB.set_cell_contents('Test', 'A1', '=Del!A1')
        WB.del_sheet('Del')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=Del!A1'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE

        WB.new_sheet('Del sheet')
        WB.set_cell_contents('Test', 'A1', "='Del sheet'!A1+1")
        WB.del_sheet('Del sheet')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == "='Del sheet'!A1+1"
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE

        WB.new_sheet('Sheet1')
        WB.set_cell_contents('Sheet1', 'A1', 'Sheet2')
        WB.set_cell_contents('Sheet1', 'A2', '=Sheet1!A1 & Sheet2')
        contents = WB.get_cell_contents('Sheet1', 'A2')
        value = WB.get_cell_value('Sheet1', 'A2')
        assert contents == '=Sheet1!A1 & Sheet2'
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.BAD_REFERENCE

    def test_bad_name(self) -> None:
        '''
        Test when given a formula with a bad name

        '''

        WB.set_cell_contents('Test', 'A1', '=UNKNOWN()')
        result_contents = WB.get_cell_contents('Test', 'A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=UNKNOWN()'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.BAD_NAME

        WB.set_cell_contents('Test', 'A1', '=a__afunc(arg1, arg2)')
        result_contents = WB.get_cell_contents('Test', 'A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=a__afunc(arg1, arg2)'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.BAD_NAME

        WB.set_cell_contents('Test', 'A3', '=A1(A2)')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=A1(A2)'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.BAD_NAME

        WB.set_cell_contents('Test', 'A3', '=and_or(A2)')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=and_or(A2)'
        assert isinstance(result_value, CellError)
        assert result_value.get_type() == CellErrorType.BAD_NAME

    def test_type_error(self) -> None:
        '''
        Test when given a formula with a type error

        '''

        WB.set_cell_contents('Test', 'A1', '="string"+123')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '="string"+123'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=123+"string"')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=123+"string"'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=123*"string"')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=123*"string"'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=-"string"')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=-"string"'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', 'string')
        WB.set_cell_contents('Test', 'A2', '2')
        WB.set_cell_contents('Test', 'A3', '=A1-A2')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=A1-A2'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A3', '=A1/A2')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=A1/A2'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A3', '=AND(True, "string")')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=AND(True, "string")'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A3', '=VERSION(True)')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=VERSION(True)'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A3', '=AND()')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert result_contents == '=AND()'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

    def test_divide_by_zero(self) -> None:
        '''
        Test when given a formula with a divide by zero error

        '''

        WB.set_cell_contents('Test', 'A1', '=12/0')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=12/0'
        assert result_value.get_type() == CellErrorType.DIVIDE_BY_ZERO
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=12')
        WB.set_cell_contents('Test', 'A2', '0')
        WB.set_cell_contents('Test', 'A4', '=A1/A2')
        result_contents = WB.get_cell_contents('Test','A4')
        result_value = WB.get_cell_value('Test', 'A4')
        assert result_contents == '=A1/A2'
        assert result_value.get_type() == CellErrorType.DIVIDE_BY_ZERO
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A3', None)
        WB.set_cell_contents('Test', 'A4', '=A1/A3')
        result_contents = WB.get_cell_contents('Test','A4')
        result_value = WB.get_cell_value('Test', 'A4')
        assert result_contents == '=A1/A3'
        assert result_value.get_type() == CellErrorType.DIVIDE_BY_ZERO
        assert isinstance(result_value, CellError)

    def test_error_as_literal(self) -> None:
        '''
        Test when given a formula with an error as a literal

        '''

        WB.set_cell_contents('Test', 'A1', '=#REF!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=#REF!'
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=#ERROR!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=#ERROR!'
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=#VALUE!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=#VALUE!'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=#CIRCREF!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=#CIRCREF!'
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=#DIV/0!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=#DIV/0!'
        assert result_value.get_type() == CellErrorType.DIVIDE_BY_ZERO
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=#NAME?')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=#NAME?'
        assert result_value.get_type() == CellErrorType.BAD_NAME
        assert isinstance(result_value, CellError)

    def test_error_in_complex_literal(self) -> None:
        '''
        Test when given a formula containing a complex operation containing a
        literal error

        '''

        WB.set_cell_contents('Test', 'A1', '=#REF!+5')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=#REF!+5'
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=#NAME?*5')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=#NAME?*5'
        assert result_value.get_type() == CellErrorType.BAD_NAME
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=5*#NAME?')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=5*#NAME?'
        assert result_value.get_type() == CellErrorType.BAD_NAME
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '=-#CIRCREF!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '=-#CIRCREF!'
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'A1', '="string"&#CIRCREF!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert result_contents == '="string"&#CIRCREF!'
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(result_value, CellError)

    def test_reference_cell_with_error(self) -> None:
        '''
        Test when given a formula that references a cell with an error

        '''

        WB.new_sheet('Test3')
        WB.set_cell_contents('Test3', 'A1', '=#REF!')
        WB.set_cell_contents('Test3', 'A2', '=A1')
        result_contents = WB.get_cell_contents('Test3','A2')
        result_value = WB.get_cell_value('Test3', 'A2')
        assert result_contents == '=A1'
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test3', 'A1', '=#ERROR!')
        WB.set_cell_contents('Test3', 'A2', '=A1')
        result_contents = WB.get_cell_contents('Test3','A2')
        result_value = WB.get_cell_value('Test3', 'A2')
        assert result_contents == '=A1'
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test3', 'A1', '=#NAME?')
        WB.set_cell_contents('Test3', 'A2', '=A1')
        result_contents = WB.get_cell_contents('Test3','A2')
        result_value = WB.get_cell_value('Test3', 'A2')
        assert result_contents == '=A1'
        assert result_value.get_type() == CellErrorType.BAD_NAME
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test3', 'A1', '=#CIRCREF!')
        WB.set_cell_contents('Test3', 'A2', '=A1')
        result_contents = WB.get_cell_contents('Test3','A2')
        result_value = WB.get_cell_value('Test3', 'A2')
        assert result_contents == '=A1'
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test3', 'A1', '=#VALUE!')
        WB.set_cell_contents('Test3', 'A2', '=A1')
        result_contents = WB.get_cell_contents('Test3','A2')
        result_value = WB.get_cell_value('Test3', 'A2')
        assert result_contents == '=A1'
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert isinstance(result_value, CellError)

        WB.set_cell_contents('Test', 'a1', '#div/0!')
        WB.set_cell_contents('Test', 'a2', '=a1+5')
        value = WB.get_cell_value('Test', 'a2')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.DIVIDE_BY_ZERO

    def test_error_ordering(self) -> None:
        '''
        Test when given a formula where multiple errors could propagate

        '''

        WB.set_cell_contents('Test', 'A1', '=B1')
        WB.set_cell_contents('Test', 'B1', '=C1')
        WB.set_cell_contents('Test', 'C1', '=B1/0')
        value = WB.get_cell_value('Test', 'C1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

        WB.set_cell_contents('Test', 'A1', '1')
        WB.set_cell_contents('Test', 'A2', '=AA99 +&* A2')
        value = WB.get_cell_value('Test', 'A2')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

        WB.set_cell_contents('Test', 'A1', '=Sheet99!A1 / 0')
        value = WB.get_cell_value('Test', 'A1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.BAD_REFERENCE

        WB.set_cell_contents('Test', 'A1', 'word')
        WB.set_cell_contents('Test', 'A2', '=A1 / 0')
        value = WB.get_cell_value('Test', 'A2')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.TYPE_ERROR

        WB.set_cell_contents('Test', 'D1', '=D2')
        WB.set_cell_contents('Test', 'D2', '=BADNAME(D1)')
        value = WB.get_cell_value('Test', 'D2')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.BAD_NAME

        value = WB.get_cell_value('Test', 'D1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.BAD_NAME

        WB.set_cell_contents('Test', 'E1', '=E2')
        WB.set_cell_contents('Test', 'E2', '=INDIRECT(E1, "bad")')
        value = WB.get_cell_value('Test', 'E2')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

        WB.set_cell_contents('Test', 'E2', '=INDIRECT("bad", E1)')
        value = WB.get_cell_value('Test', 'E2')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE
