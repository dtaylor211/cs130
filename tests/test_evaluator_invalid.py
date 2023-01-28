import context

import pytest
from lark import Lark, Tree
from decimal import Decimal

from sheets.workbook import Workbook
from sheets.cell_error import CellError, CellErrorType

WB = Workbook()
WB.new_sheet('Test')

class TestEvaluatorInvalid:
    '''
    Tests the formula parser and evaluator using invalid inputs

    '''

    def test_parse_errors(self):
        WB.set_cell_contents('Test', 'A1', '=1E+4')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=1E+4')
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=A1A2')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=A1A2')
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=1**2')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=1**2')
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=A1(A2)')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=A1(A2)')
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=A1+(A2-A1')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=A1+(A2-A1')
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=A1+A2-A1)')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=A1+A2-A1)')
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=A 1')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=A 1')
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert(isinstance(result_value, CellError))


    def test_circular_reference(self):
        WB.set_cell_contents('Test', 'A1', '12')
        WB.set_cell_contents('Test', 'A2', '13')
        WB.set_cell_contents('Test', 'A3', '=A1+A2')
        WB.set_cell_contents('Test', 'A4', '=A1+A2+A3')

        WB.set_cell_contents('Test', 'A1', '=A4')
        result_contents = WB.get_cell_contents('Test','A4')
        result_value = WB.get_cell_value('Test', 'A4')
        assert(result_contents == '=A1+A2+A3')
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=A1')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=A1')
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert(isinstance(result_value, CellError))


    def test_bad_reference(self):
        WB.set_cell_contents('Test', 'A1', '=Test2!A1')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=Test2!A1')
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=AAAAA1')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=AAAAA1')
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=AAAA10000')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=AAAA10000')
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE
        assert(isinstance(result_value, CellError))

        WB.new_sheet('Del')
        WB.set_cell_contents('Del', 'A1', '1')
        WB.set_cell_contents('Test', 'A1', '=Del!A1')
        WB.del_sheet('Del')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=Del!A1')
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE
        assert(isinstance(result_value, CellError))

# [-5] If A refers to a cell in some sheet S, but then sheet S is deleted, A
# should be updated to be a BAD_REFERENCE error. 
# ALREADY TEST FOR THIS AND IT PASSES???????
# THINK ABOUT POSSIBLE CORNER CASES
# Possibly the handling sheet names with spaces

    def test_bad_name(self):
        # to be implemented in later projects
        pass


    def test_type_errors(self):
        WB.set_cell_contents('Test', 'A1', '="string"+123')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '="string"+123')
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=123+"string"')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=123+"string"')
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=123*"string"')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=123*"string"')
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=-"string"')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=-"string"')
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', 'string')
        WB.set_cell_contents('Test', 'A2', '2')
        WB.set_cell_contents('Test', 'A3', '=A1-A2')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert(result_contents == '=A1-A2')
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A3', '=A1/A2')
        result_contents = WB.get_cell_contents('Test','A3')
        result_value = WB.get_cell_value('Test', 'A3')
        assert(result_contents == '=A1/A2')
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert(isinstance(result_value, CellError))

    
    def test_divide_by_zero(self):
        WB.set_cell_contents('Test', 'A1', '=12/0')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=12/0')
        assert result_value.get_type() == CellErrorType.DIVIDE_BY_ZERO
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=12')
        WB.set_cell_contents('Test', 'A2', '0')
        WB.set_cell_contents('Test', 'A4', '=A1/A2')
        result_contents = WB.get_cell_contents('Test','A4')
        result_value = WB.get_cell_value('Test', 'A4')
        assert(result_contents == '=A1/A2')
        assert result_value.get_type() == CellErrorType.DIVIDE_BY_ZERO
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A3', None)
        WB.set_cell_contents('Test', 'A4', '=A1/A3')
        result_contents = WB.get_cell_contents('Test','A4')
        result_value = WB.get_cell_value('Test', 'A4')
        assert(result_contents == '=A1/A3')
        assert result_value.get_type() == CellErrorType.DIVIDE_BY_ZERO
        assert(isinstance(result_value, CellError))

    
    def test_errors_as_literals(self):
        WB.set_cell_contents('Test', 'A1', '=#REF!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=#REF!')
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=#ERROR!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=#ERROR!')
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=#VALUE!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=#VALUE!')
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=#CIRCREF!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=#CIRCREF!')
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=#DIV/0!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=#DIV/0!')
        assert result_value.get_type() == CellErrorType.DIVIDE_BY_ZERO
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=#NAME?')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=#NAME?')
        assert result_value.get_type() == CellErrorType.BAD_NAME
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=#REF!+5')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=#REF!+5')
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=#NAME?*5')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=#NAME?*5')
        assert result_value.get_type() == CellErrorType.BAD_NAME
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=5*#NAME?')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=5*#NAME?')
        assert result_value.get_type() == CellErrorType.BAD_NAME
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '=-#CIRCREF!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '=-#CIRCREF!')
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'A1', '="string"&#CIRCREF!')
        result_contents = WB.get_cell_contents('Test','A1')
        result_value = WB.get_cell_value('Test', 'A1')
        assert(result_contents == '="string"&#CIRCREF!')
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert(isinstance(result_value, CellError))

    
    def test_reference_cells_with_errors(self):
        WB.new_sheet('Test3')
        WB.set_cell_contents('Test3', 'A1', '=#REF!')
        WB.set_cell_contents('Test3', 'A2', '=A1')
        result_contents = WB.get_cell_contents('Test3','A2')
        result_value = WB.get_cell_value('Test3', 'A2')
        assert(result_contents == '=A1')
        assert result_value.get_type() == CellErrorType.BAD_REFERENCE
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test3', 'A1', '=#ERROR!')
        WB.set_cell_contents('Test3', 'A2', '=A1')
        result_contents = WB.get_cell_contents('Test3','A2')
        result_value = WB.get_cell_value('Test3', 'A2')
        assert(result_contents == '=A1')
        assert result_value.get_type() == CellErrorType.PARSE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test3', 'A1', '=#NAME?')
        WB.set_cell_contents('Test3', 'A2', '=A1')
        result_contents = WB.get_cell_contents('Test3','A2')
        result_value = WB.get_cell_value('Test3', 'A2')
        assert(result_contents == '=A1')
        assert result_value.get_type() == CellErrorType.BAD_NAME
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test3', 'A1', '=#CIRCREF!')
        WB.set_cell_contents('Test3', 'A2', '=A1')
        result_contents = WB.get_cell_contents('Test3','A2')
        result_value = WB.get_cell_value('Test3', 'A2')
        assert(result_contents == '=A1')
        assert result_value.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test3', 'A1', '=#VALUE!')
        WB.set_cell_contents('Test3', 'A2', '=A1')
        result_contents = WB.get_cell_contents('Test3','A2')
        result_value = WB.get_cell_value('Test3', 'A2')
        assert(result_contents == '=A1')
        assert result_value.get_type() == CellErrorType.TYPE_ERROR
        assert(isinstance(result_value, CellError))

        WB.set_cell_contents('Test', 'a1', '#div/0!')
        WB.set_cell_contents('Test', 'a2', '=a1+5')
        value = WB.get_cell_value('Test', 'a2')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.DIVIDE_BY_ZERO


    def test_error_ordering(self):
        # add more to this later
        WB.set_cell_contents('Test', 'A1', '=B1')
        WB.set_cell_contents('Test', 'B1', '=C1')
        WB.set_cell_contents('Test', 'C1', '=B1/0')
        value = WB.get_cell_value('Test', 'C1')
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE