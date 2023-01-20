import pytest
import context
from lark import Lark, Tree
from sheets.evaluator import Evaluator
from sheets.workbook import Workbook
from decimal import Decimal

wb = Workbook()
wb.new_sheet('Test')
evaluator = Evaluator(wb, 'Test')
parser = Lark.open('../sheets/formulas.lark', start='formula', rel_to=__file__)

class TestEvaluator:
    '''
    Tests the formula parser and evaluator using valid inputs

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
        wb.set_cell_contents('Test', 'A6', '12.0000000')
        wb.set_cell_contents('Test', 'A7', '\'    123')

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

        tree = parser.parse('=A6')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', [Decimal(12)]))

        tree = parser.parse('=A7')
        result = evaluator.transform(tree)
        assert(result == Tree('cell_ref', ['123']))

    
    def test_string_concatenation(self):
        wb.set_cell_contents('Test', 'A1', 'string1')
        wb.set_cell_contents('Test', 'A2', 'string2')
        wb.set_cell_contents('Test', 'A3', '="Donnie Pinkston is a goat"')
        wb.set_cell_contents('Test', 'A4', '"string3"')
        wb.set_cell_contents('Test', 'A5', "'Anakin")
        wb.set_cell_contents('Test', 'A6', None)

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

        tree = parser.parse('=7&9')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['79']))

        tree = parser.parse('=7&9&"string"')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['79string']))

        tree = parser.parse('=7.1&9.2&"string"')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['7.19.2string']))

        tree = parser.parse('=A1&A6')
        result = evaluator.transform(tree)
        assert(result == Tree('string', ['string1']))
    

    def test_unary_operations(self):
        wb.set_cell_contents('Test', 'A1', '2')
        wb.set_cell_contents('Test', 'A2', '2.2')
        wb.set_cell_contents('Test', 'A3', '-4')
        wb.set_cell_contents('Test', 'A4', '+25')
        wb.set_cell_contents('Test', 'A5', None)

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

        tree = parser.parse('=+A5')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(0)]))

        tree = parser.parse('=-A5')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(0)])) # should be 0?
    

    def test_addition_subtraction(self):
        wb.set_cell_contents('Test', 'A1', '1')
        wb.set_cell_contents('Test', 'A2', '=2')
        wb.set_cell_contents('Test', 'A3', '-3.25')
        wb.set_cell_contents('Test', 'A4', "'123")
        wb.set_cell_contents('Test', 'A5', None)

        tree = parser.parse('=1+1')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(2)]))

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
        
        tree = parser.parse('=A1+A5')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(1)]))

        tree = parser.parse('=A1-A5')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(1)]))


    def test_multiplication_division(self):
        wb.set_cell_contents('Test', 'A1', '1')
        wb.set_cell_contents('Test', 'A2', '2')
        wb.set_cell_contents('Test', 'A3', '-3.25')
        wb.set_cell_contents('Test', 'A4', "'123")
        wb.set_cell_contents('Test', 'A5', '3')
        wb.set_cell_contents('Test', 'A6', None)

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

        tree = parser.parse('=A1/A5')
        result = evaluator.transform(tree)
        assert(result == Tree('number', 
                              [Decimal('0.3333333333333333333333333333')]))

        tree = parser.parse('=A1*A6')
        result = evaluator.transform(tree)
        assert(result == Tree('number', [Decimal(0)]))


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


    def test_reference_same_sheet(self):
        (index, name) = wb.new_sheet("Sheet1")
        assert name == "Sheet1"
        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "B2", "=Sheet1!A1")
        contents = wb.get_cell_contents(name, "A1")
        assert contents == "1"
        contents = wb.get_cell_contents(name, "B2")
        assert contents == "=Sheet1!A1"
        value = wb.get_cell_value(name, "A1")
        assert value == Decimal(1)
        value = wb.get_cell_value(name, "B2")
        assert value == Decimal(1)

        (index, name) = wb.new_sheet("Sheet1\\")
        assert name == "Sheet1\\"
        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "B2", "='Sheet1\\'!A1")
        contents = wb.get_cell_contents(name, "A1")
        assert contents == "1"
        contents = wb.get_cell_contents(name, "B2")
        assert contents == "=Sheet1!A1"
        value = wb.get_cell_value(name, "A1")
        assert value == Decimal(1)
        value = wb.get_cell_value(name, "B2")
        assert value == Decimal(1)

        (index, name) = wb.new_sheet("Sheet2")
        assert name == "Sheet2"
        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "B2", "=shEet2!A1")
        contents = wb.get_cell_contents(name, "A1")
        assert contents == "1"
        contents = wb.get_cell_contents(name, "B2")
        assert contents == "=shEet2!A1"
        value = wb.get_cell_value(name, "A1")
        assert value == Decimal(1)
        value = wb.get_cell_value(name, "B2")
        assert value == Decimal(1)

        (index, name) = wb.new_sheet("Other Totals")
        assert name == "Other Totals"
        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "B2", "='Other Totals'!A1")
        contents = wb.get_cell_contents(name, "A1")
        assert contents == "1"
        contents = wb.get_cell_contents(name, "B2")
        assert contents == "='Other Totals'!A1"
        value = wb.get_cell_value(name, "A1")
        assert value == Decimal(1)
        value = wb.get_cell_value(name, "B2")
        assert value == Decimal(1)

    def test_reference_other_sheet(self):
        (index, name) = wb.new_sheet("June Totals")
        assert name == "June Totals"
        (index, name) = wb.new_sheet("July Totals")
        assert name == "July Totals"

        wb.set_cell_contents("June Totals", "A1", "1")
        wb.set_cell_contents("June Totals", "A2", "2")
        wb.set_cell_contents("July Totals", "B2", "='June Totals'!A1")
        contents = wb.get_cell_contents("June Totals", "A1")
        assert contents == "1"
        contents = wb.get_cell_contents("June Totals", "A2")
        assert contents == "2"
        contents = wb.get_cell_contents("July Totals", "B2")
        assert contents == "='June Totals'!A1"
        value = wb.get_cell_value("June Totals", "A1")
        assert value == Decimal(1)
        value = wb.get_cell_value("June Totals", "A2")
        assert value == Decimal(2)
        value = wb.get_cell_value("July Totals", "B2")
        assert value == Decimal(1)

        # is the user expected to put 'June Totals'! after accessing other sheet?
        wb.set_cell_contents("June Totals", "B2", "=A2")
        contents = wb.get_cell_contents("June Totals", "B2")
        assert contents == "=A2"
        value = wb.get_cell_value("June Totals", "B2")
        assert value == Decimal(2)