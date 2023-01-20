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

class TestEvaluator:
    '''
    Tests the formula parser and evaluator

    '''

    def test_num_literals(self):
        tree = parser.parse('=123')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('123')]))

        tree = parser.parse('=12.3')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('12.3')]))

        tree = parser.parse('=.2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('.2')]))

        tree = parser.parse('=0010.00200')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('10.002')]))

        tree = parser.parse('=   0010.')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('10')]))

        tree = parser.parse('=0010.      ')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('10')]))

        tree = parser.parse('=   0010.    ')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('10')]))

        tree = parser.parse('=0.2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('0.2')]))

        tree = parser.parse('=000000000.2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('0.2')]))

        tree = parser.parse('=1000000')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('1000000')]))

        tree = parser.parse('=12.00000000')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('12')]))

        tree = parser.parse('=12.000000001')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('12.000000001')]))


    def test_string_literals(self):
        tree = parser.parse('="\'"')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['\'']))

        tree = parser.parse('=""')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['']))

        tree = parser.parse('="string"')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['string']))

        tree = parser.parse('="123"')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['123']))

        tree = parser.parse('="this is a string with spaces"')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['this is a string with spaces']))

        tree = parser.parse('=      "Jar Jar Binks was always supposed to'\
                            ' be the real phantom menace"     ')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['Jar Jar Binks was always supposed '\
                                         'to be the real phantom menace']))


    def test_cell_references(self):
        wb.set_cell_contents('Test', 'A1', '1')
        wb.set_cell_contents('Test', 'A2', '2')
        wb.set_cell_contents('Test', 'A3', 'string')
        wb.set_cell_contents('Test', 'A4', '12string')
        wb.set_cell_contents('Test', 'a5', 'DarthJarJar')

        tree = parser.parse('=A1')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', [Decimal(1)]))

        tree = parser.parse('=a1')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', [Decimal(1)]))

        tree = parser.parse('=A2')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', [Decimal(2)]))
        
        tree = parser.parse('=A3')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', ["string"]))

        tree = parser.parse('=A4')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', ["12string"]))

        tree = parser.parse('=        A4')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', ["12string"]))

        tree = parser.parse('=A4   ')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', ["12string"]))

        tree = parser.parse('=    A4   ')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', ["12string"]))

        tree = parser.parse('=A5')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', ["DarthJarJar"]))


    def test_invalid_literals(self):
        wb.set_cell_contents('Test', 'A1', '=1E+4')
        result_contents = wb.get_cell_contents('Test','A1')
        result_value = wb.get_cell_value('Test', 'A1')
        assert(result_contents == '#ERROR!')
        assert(isinstance(result_value, CellError))

        wb.set_cell_contents('Test', 'A1', '=A1A2')
        result = wb.get_cell_contents('Test','A1')
        assert(result == '#ERROR!')

        # add more

    
    def test_string_concatenation(self):
        wb.set_cell_contents('Test', 'A1', 'string1')
        wb.set_cell_contents('Test', 'A2', 'string2')
        wb.set_cell_contents('Test', 'A3', '="Donnie Pinkston is a goat"')
        wb.set_cell_contents('Test', 'A4', '"string3"')
        wb.set_cell_contents('Test', 'A5', "'Anakin")

        tree = parser.parse('="this is "&"a string"')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ["this is a string"]))

        tree = parser.parse('="this is "&A1')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ["this is string1"]))

        tree = parser.parse('=A1&A2')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ["string1string2"]))

        tree = parser.parse('=A1&A3')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ["string1Donnie Pinkston is a goat"]))

        tree = parser.parse('=A1&A4')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['string1"string3"']))

        tree = parser.parse('=A1&A5')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['string1Anakin']))

        tree = parser.parse('= A1   &  A5    ')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['string1Anakin']))
    

    def test_unary_operations(self):
        wb.set_cell_contents('Test', 'A1', '2')
        wb.set_cell_contents('Test', 'A2', '2.2')
        wb.set_cell_contents('Test', 'A3', '-4')
        wb.set_cell_contents('Test', 'A4', '+25')

        tree = parser.parse('=-34')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-34)]))

        tree = parser.parse('=-A1')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-2)]))

        tree = parser.parse('=-A2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('-2.2')]))

        tree = parser.parse('=-A3')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(4)]))

        tree = parser.parse('=+A3')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-4)]))

        tree = parser.parse('=-A4')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-25)]))

        tree = parser.parse('= - A4  ')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-25)]))
    

    def test_addition_subtraction(self):
        wb.set_cell_contents('Test', 'A1', '1')
        wb.set_cell_contents('Test', 'A2', '=2')
        wb.set_cell_contents('Test', 'A3', '-3.25')
        wb.set_cell_contents('Test', 'A4', "'123")

        tree = parser.parse('=A1+A2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(3)]))

        tree = parser.parse('=34+A1')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(35)]))

        tree = parser.parse('=-34+A1')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-33)]))

        tree = parser.parse('=A1-A2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-1)]))

        tree = parser.parse('=A3-A2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-5.25)]))

        tree = parser.parse('= A3 - A2   ')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-5.25)]))

        tree = parser.parse('=A4-A2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(121)]))


    def test_multiplication_division(self):
        wb.set_cell_contents('Test', 'A1', '1')
        wb.set_cell_contents('Test', 'A2', '2')
        wb.set_cell_contents('Test', 'A3', '-3.25')
        wb.set_cell_contents('Test', 'A4', "'123")

        tree = parser.parse('=A1*A2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(2)]))

        tree = parser.parse('=A1/A2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(0.5)]))

        tree = parser.parse('=A2*A3')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-6.5)]))

        tree = parser.parse('=A3/A2')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-1.625)]))

        tree = parser.parse('=A3/A3')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(1)]))

        tree = parser.parse('=A3*A4')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(-399.75)]))


    def test_complex_formula(self):
        wb.set_cell_contents('Test', 'A1', 's1')
        wb.set_cell_contents('Test', 'A2', 's2')
        wb.set_cell_contents('Test', 'A3', '-3.25')
        wb.set_cell_contents('Test', 'A4', "'123")
        wb.set_cell_contents('Test', 'A5', "=A3*2")
        wb.set_cell_contents('Test', 'A6', '1.7')

        tree = parser.parse('=A1&A2&A1&A2&(A1&A2)')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['s1s2s1s2s1s2']))

        tree = parser.parse('=A2&(1+2)')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['s23']))

        tree = parser.parse('=-(2+4)')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('-6')]))

        tree = parser.parse('=A1&A4&(A2&A3)')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['s1123s2-3.25']))

        tree = parser.parse('=A1&A2&(A4*A3)')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['s1s2-399.75']))

        # add sheet name tests for passing into cell field

        tree = parser.parse('=Test!A1&A2&(A4*A3)')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['s1s2-399.75']))

        tree = parser.parse('=A3*2.00000*10')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('-65')]))

        tree = parser.parse('=-A3*2+(-A4)')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('-116.5')]))

        tree = parser.parse('=A3+A5*(A6/2)+((82-A3)+7*2.04+A6)')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal('92.455')]))
