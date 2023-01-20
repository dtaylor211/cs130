import pytest
import context
from lark import Lark, Tree
from sheets.formula_evaluator import Evaluator
from sheets.workbook import Workbook
from decimal import Decimal
from sheets.cell_error import CellError, CellErrorType

wb = Workbook()
wb.new_sheet('Test')
evaluator = Evaluator(wb, 'Test')
parser = Lark.open('../sheets/formulas.lark', start='formula', rel_to=__file__)

class TestEvaluatorInvalid:
    '''
    Tests the formula parser and evaluator using invalid inputs

    '''

    def test_parse_errors(self):
        wb.set_cell_contents('Test', 'A1', '=1E+4')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#ERROR!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=A1A2')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#ERROR!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=A1A2')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#ERROR!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=1**2')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#ERROR!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=A1(A2)')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#ERROR!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=A1+(A2-A1')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#ERROR!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=A1+A2-A1)')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#ERROR!')
        assert(isinstance(result_value, CellError))


    def test_circular_reference(self):
        pass


    def test_bad_reference(self):
        wb.set_cell_contents('Test', 'A1', '=Test2!A1')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#REF!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=AAAA1')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#REF!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=AAA10000')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#REF!')
        assert(isinstance(result_value, CellError))


    def test_bad_name(self):
        pass


    def test_type_errors(self):
        wb.set_cell_contents('Test', 'A1', '="string"+123')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#VALUE!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=123+"string"')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#VALUE!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=123*"string"')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#VALUE!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=-"string"')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#VALUE!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', 'string')
        wb.set_cell_contents('Test', 'A2', '2')
        wb.set_cell_contents('Test', 'A3', '=A1-A2')
        tree = parser.parse('=A1-A2')
        print(evaluator.transform(tree))
        result_contents = wb.get_cell_contents('Test','A3')
        print(result_contents)
        result_value = wb.get_cell_value('Test', 'A3')
        print(result_value)
        assert(result_contents == '#VALUE!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A3', '=A1/A2')
        
        result_contents = wb.get_cell_contents('Test','A3')
        result_value = wb.get_cell_value('Test', 'A3')
        assert(result_contents == '#VALUE!')
        assert(isinstance(result_value, CellError))

    
    def test_divide_by_zero(self):
        wb.set_cell_contents('Test', 'A1', '=12/0')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#DIV/0!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=12/0')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#DIV/0!')
        assert(isinstance(result_value, CellError))

    
    def test_errors_as_literals(self):
        pass
