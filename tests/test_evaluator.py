'''
Test Evaluator

Tests the Evaluator module found at ../sheets/evaluator.py with
valid inputs.

GLOBAL_VARIABLES:
- WB (Workbook) - the Workbook used for this test suite
- EVALUATOR (Evaluator) - the Evaluator used for this test suite
- PARSER (Any) - the Parser used for this test suite

Classes:
- TestEvaluator

    Methods:
    - test_num_literals(object) -> None
    - test_string_literals(object) -> None
    - test_cell_references(object) -> None
    - test_string_concatenation(object) -> None
    - test_unary_operations(object) -> None
    - test_addition_subtraction(object) -> None
    - test_multiplication_division(object) -> None
    - test_comparison(object) -> None
    - test_complex_formula(object) -> None
    - test_reference_same_sheet(object) -> None
    - test_reference_other_sheet(object) -> None

'''


from decimal import Decimal

from lark import Lark, Tree

# pylint: disable=unused-import, import-error
import context
from sheets.evaluator import Evaluator
from sheets.workbook import Workbook


WB = Workbook()
WB.new_sheet('Test')
EVALUATOR = Evaluator(WB, 'Test')
PARSER = Lark.open('../sheets/formulas.lark', start='formula', rel_to=__file__)

class TestEvaluator:
    '''
    Tests the formula parser and evaluator using valid inputs

    '''

    def test_num_literals(self) -> None:
        '''
        Test when given a formula of numeric literals

        '''

        tree = PARSER.parse('=123')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('123')])

        tree = PARSER.parse('=12.3')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('12.3')])

        tree = PARSER.parse('=.2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('.2')])

        tree = PARSER.parse('=0010.00200')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('10.002')])

        tree = PARSER.parse('=   0010.')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('10')])

        tree = PARSER.parse('=0010.      ')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('10')])

        tree = PARSER.parse('=   0010.    ')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('10')])

        tree = PARSER.parse('=0.2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('0.2')])

        tree = PARSER.parse('=000000000.2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('0.2')])

        tree = PARSER.parse('=1000000')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('1000000')])

        tree = PARSER.parse('=12.00000000')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('12')])

        tree = PARSER.parse('=12.000000001')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('12.000000001')])

    def test_string_literals(self) -> None:
        '''
        Test when given a formula of string literals

        '''

        tree = PARSER.parse('="\'"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['\''])

        tree = PARSER.parse('=""')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', [''])

        tree = PARSER.parse('="string"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['string'])

        tree = PARSER.parse('="123"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['123'])

        tree = PARSER.parse('="this is a string with spaces"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['this is a string with spaces'])

        tree = PARSER.parse('=      "Jar Jar Binks was always supposed to'\
                            ' be the real phantom menace"     ')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['Jar Jar Binks was always supposed '\
                                         'to be the real phantom menace'])

    def test_cell_references(self) -> None:
        '''
        Test when given a formula including cell references

        '''

        WB.set_cell_contents('Test', 'A1', '1')
        WB.set_cell_contents('Test', 'A2', '2')
        WB.set_cell_contents('Test', 'A3', 'string')
        WB.set_cell_contents('Test', 'A4', '12string')
        WB.set_cell_contents('Test', 'a5', 'DarthJarJar')
        WB.set_cell_contents('Test', 'A6', '12.0000000')
        WB.set_cell_contents('Test', 'A7', '\'    123')
        WB.set_cell_contents('Test', 'A8', '=A9')

        tree = PARSER.parse('=A1')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal(1)])

        tree = PARSER.parse('=a1')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal(1)])

        tree = PARSER.parse('=A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal(2)])

        tree = PARSER.parse('=A3')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', ["string"])

        tree = PARSER.parse('=A4')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', ["12string"])

        tree = PARSER.parse('=        A4')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', ["12string"])

        tree = PARSER.parse('=A4   ')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', ["12string"])

        tree = PARSER.parse('=    A4   ')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', ["12string"])

        tree = PARSER.parse('=A5')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', ["DarthJarJar"])

        tree = PARSER.parse('=A6')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal(12)])

        tree = PARSER.parse('=A7')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', ['    123'])

        tree = PARSER.parse('=A8')
        result = EVALUATOR.transform(tree)
        assert result == Tree('cell_ref', [Decimal(0)])

    def test_string_concatenation(self) -> None:
        '''
        Test when given a formula with the string concatenation operator

        '''

        WB.set_cell_contents('Test', 'A1', 'string1')
        WB.set_cell_contents('Test', 'A2', 'string2')
        WB.set_cell_contents('Test', 'A3', '="Donnie Pinkston is a goat"')
        WB.set_cell_contents('Test', 'A4', '"string3"')
        WB.set_cell_contents('Test', 'A5', "'Anakin")
        WB.set_cell_contents('Test', 'A6', None)

        tree = PARSER.parse('="this is "&"a string"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ["this is a string"])

        tree = PARSER.parse('="this is "&A1')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ["this is string1"])

        tree = PARSER.parse('=A1&A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ["string1string2"])

        tree = PARSER.parse('=A1&A3')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ["string1Donnie Pinkston is a goat"])

        tree = PARSER.parse('=A1&A4')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['string1"string3"'])

        tree = PARSER.parse('=A1&A5')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['string1Anakin'])

        tree = PARSER.parse('= A1   &  A5    ')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['string1Anakin'])

        tree = PARSER.parse('=7&9')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['79'])

        tree = PARSER.parse('=7&9&"string"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['79string'])

        tree = PARSER.parse('=7.1&9.2&"string"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['7.19.2string'])

        tree = PARSER.parse('=A1&A6')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['string1'])

        tree = PARSER.parse('=Test!A1&"test"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['string1test'])

    def test_unary_operations(self) -> None:
        '''
        Test when given a formula with a unary operator (+/-)

        '''

        WB.set_cell_contents('Test', 'A1', '2')
        WB.set_cell_contents('Test', 'A2', '2.2')
        WB.set_cell_contents('Test', 'A3', '-4')
        WB.set_cell_contents('Test', 'A4', '+25')
        WB.set_cell_contents('Test', 'A5', None)

        tree = PARSER.parse('=-34')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-34)])

        tree = PARSER.parse('=-A1')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-2)])

        tree = PARSER.parse('=-A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('-2.2')])

        tree = PARSER.parse('=-A3')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(4)])

        tree = PARSER.parse('=+A3')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-4)])

        tree = PARSER.parse('=-A4')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-25)])

        tree = PARSER.parse('= - A4  ')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-25)])

        tree = PARSER.parse('=+A5')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(0)])

        tree = PARSER.parse('=-A5')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(0)])

    def test_addition_subtraction(self) -> None:
        '''
        Test when given a formula with an addition operator (+/-)

        '''

        WB.set_cell_contents('Test', 'A1', '1')
        WB.set_cell_contents('Test', 'A2', '=2')
        WB.set_cell_contents('Test', 'A3', '-3.25')
        WB.set_cell_contents('Test', 'A4', "'123")
        WB.set_cell_contents('Test', 'A5', None)
        WB.set_cell_contents('Test', 'A6', '=A1+A2')
        WB.set_cell_contents('Test', 'A7', '=A1+A6')

        tree = PARSER.parse('=1+1')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(2)])

        tree = PARSER.parse('=A1+A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(3)])

        tree = PARSER.parse('=34+A1')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(35)])

        tree = PARSER.parse('=-34+A1')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-33)])

        tree = PARSER.parse('=A1-A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-1)])

        tree = PARSER.parse('=A3-A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-5.25)])

        tree = PARSER.parse('= A3 - A2   ')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-5.25)])

        tree = PARSER.parse('=A4-A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(121)])

        tree = PARSER.parse('=A1+A5')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(1)])

        tree = PARSER.parse('=A1-A5')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(1)])

        tree = PARSER.parse('=A1+A6+A7')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(8)])

    def test_multiplication_division(self) -> None:
        '''
        Test when given a formula with a multiplication operator (*//)

        '''

        WB.set_cell_contents('Test', 'A1', '1')
        WB.set_cell_contents('Test', 'A2', '2')
        WB.set_cell_contents('Test', 'A3', '-3.25')
        WB.set_cell_contents('Test', 'A4', "'123")
        WB.set_cell_contents('Test', 'A5', '3')
        WB.set_cell_contents('Test', 'A6', None)

        tree = PARSER.parse('=A1*A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(2)])

        tree = PARSER.parse('=A1/A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(0.5)])

        tree = PARSER.parse('=A2*A3')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-6.5)])

        tree = PARSER.parse('=A3/A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-1.625)])

        tree = PARSER.parse('=A3/A3')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(1)])

        tree = PARSER.parse('=A3*A4')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(-399.75)])

        tree = PARSER.parse('=A1/A5')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number',
                              [Decimal('0.3333333333333333333333333333')])

        tree = PARSER.parse('=A1*A6')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal(0)])

    def test_comparison(self) -> None:
        '''
        Test when given a formula with a comparison operator
        (</<=/>/>=/=/==/!=/<>)

        '''

        WB.set_cell_contents('Test', 'A1', 'string1')
        WB.set_cell_contents('Test', 'A2', 'string2')
        WB.set_cell_contents('Test', 'A3', '=False')
        WB.set_cell_contents('Test', 'A4', '=True')

        tree = PARSER.parse('=A1<A2')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('="a"<"["')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('="a"<"["')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('="BLUE"="blue"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        tree = PARSER.parse('="BLUE"<"blue"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('="BLUE">"blue"')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [False])

        tree = PARSER.parse('=AND(True, 4)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('bool', [True])

        # tree = PARSER.parse('=A3<A4')
        # result = EVALUATOR.transform(tree)
        # assert result == Tree('bool', [True])

    def test_complex_formula(self) -> None:
        '''
        Test when given a complex formula with parenthesis, multiple
        operators, etc.

        '''

        WB.set_cell_contents('Test', 'A1', 's1')
        WB.set_cell_contents('Test', 'A2', 's2')
        WB.set_cell_contents('Test', 'A3', '-3.25')
        WB.set_cell_contents('Test', 'A4', "'123")
        WB.set_cell_contents('Test', 'A5', "=A3*2")
        WB.set_cell_contents('Test', 'A6', '1.7')

        tree = PARSER.parse('=A1&A2&A1&A2&(A1&A2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['s1s2s1s2s1s2'])

        tree = PARSER.parse('=A2&(1+2)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['s23'])

        tree = PARSER.parse('=-(2+4)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('-6')])

        tree = PARSER.parse('=A1&A4&(A2&A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['s1123s2-3.25'])

        tree = PARSER.parse('=A1&A2&(A4*A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['s1s2-399.75'])

        tree = PARSER.parse('=Test!A1&A2&(A4*A3)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('string', ['s1s2-399.75'])

        tree = PARSER.parse('=A3*2.00000*10')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('-65')])

        tree = PARSER.parse('=-A3*2+(-A4)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('-116.5')])

        tree = PARSER.parse('=A3+A5*(A6/2)+((82-A3)+7*2.04+A6)')
        result = EVALUATOR.transform(tree)
        assert result == Tree('number', [Decimal('92.455')])

    def test_reference_same_sheet(self) -> None:
        '''
        Test when given a formula that references the same sheet

        '''

        _, name = WB.new_sheet("Sheet1")
        assert name == "Sheet1"
        WB.set_cell_contents(name, "A1", "1")
        WB.set_cell_contents(name, "B2", "=Sheet1!A1")
        contents = WB.get_cell_contents(name, "A1")
        assert contents == "1"
        contents = WB.get_cell_contents(name, "B2")
        assert contents == "=Sheet1!A1"
        value = WB.get_cell_value(name, "A1")
        assert value == Decimal(1)
        value = WB.get_cell_value(name, "B2")
        assert value == Decimal(1)

        _, name = WB.new_sheet("Sheet2")
        WB.set_cell_contents(name, "A1", "1")
        WB.set_cell_contents(name, "B2", "=shEet2!A1")
        contents = WB.get_cell_contents(name, "A1")
        assert contents == "1"
        contents = WB.get_cell_contents(name, "B2")
        assert contents == "=shEet2!A1"
        value = WB.get_cell_value(name, "A1")
        assert value == Decimal(1)
        value = WB.get_cell_value(name, "B2")
        assert value == Decimal(1)

        _, name = WB.new_sheet("Other Totals")
        assert name == "Other Totals"
        WB.set_cell_contents(name, "A1", "1")
        WB.set_cell_contents(name, "B2", "='Other Totals'!A1")
        contents = WB.get_cell_contents(name, "A1")
        assert contents == "1"
        contents = WB.get_cell_contents(name, "B2")
        assert contents == "='Other Totals'!A1"
        value = WB.get_cell_value(name, "A1")
        assert value == Decimal(1)
        value = WB.get_cell_value(name, "B2")
        assert value == Decimal(1)

        WB.set_cell_contents(name, "C3", "='Other Totals'!A1+'Other Totals'!B2")
        contents = WB.get_cell_contents(name, "C3")
        assert contents == "='Other Totals'!A1+'Other Totals'!B2"
        value = WB.get_cell_value(name, "C3")
        assert value == Decimal(2)
        WB.set_cell_contents(name, "A1", "2")
        contents = WB.get_cell_contents(name, "A1")
        assert contents == "2"
        contents = WB.get_cell_contents(name, "C3")
        assert contents == "='Other Totals'!A1+'Other Totals'!B2"
        value = WB.get_cell_value(name, "A1")
        assert value == Decimal(2)
        value = WB.get_cell_value(name, "C3")
        assert value == Decimal(4)

    def test_reference_other_sheet(self) -> None:
        '''
        Test when given a formula that references another sheet

        '''

        _, name = WB.new_sheet("June Totals")
        assert name == "June Totals"
        _, name = WB.new_sheet("July Totals")
        assert name == "July Totals"

        WB.set_cell_contents("June Totals", "A1", "1")
        WB.set_cell_contents("June Totals", "A2", "2")
        WB.set_cell_contents("July Totals", "B2", "='June Totals'!A1")
        contents = WB.get_cell_contents("June Totals", "A1")
        assert contents == "1"
        contents = WB.get_cell_contents("June Totals", "A2")
        assert contents == "2"
        contents = WB.get_cell_contents("July Totals", "B2")
        assert contents == "='June Totals'!A1"
        value = WB.get_cell_value("June Totals", "A1")
        assert value == Decimal(1)
        value = WB.get_cell_value("June Totals", "A2")
        assert value == Decimal(2)
        value = WB.get_cell_value("July Totals", "B2")
        assert value == Decimal(1)

        WB.set_cell_contents("June Totals", "B2", "=A2")
        contents = WB.get_cell_contents("June Totals", "B2")
        assert contents == "=A2"
        value = WB.get_cell_value("June Totals", "B2")
        assert value == Decimal(2)

        WB.set_cell_contents("June Totals", "B1", "='August Totals'!A1+3")
        WB.new_sheet("August Totals")
        contents = WB.get_cell_contents("June Totals", "B1")
        assert contents == "='August Totals'!A1+3"
        value = WB.get_cell_value("June Totals", "B1")
        assert value == Decimal(3)
        WB.set_cell_contents("August Totals", "A1", "1")
        contents = WB.get_cell_contents("June Totals", "B1")
        assert contents == "='August Totals'!A1+3"
        value = WB.get_cell_value("June Totals", "B1")
        assert value == Decimal(4)

        WB.del_sheet("August Totals")
        WB.set_cell_contents("June Totals", "B3", "='June Totals'!B1+August!A1")
        WB.new_sheet("August Totals")
        WB.new_sheet("August")
        contents = WB.get_cell_contents("June Totals", "B3")
        assert contents == "='June Totals'!B1+August!A1"
        value = WB.get_cell_value("June Totals", "B3")
        assert value == Decimal(3)
        WB.set_cell_contents("August Totals", "A1", "1")
        contents = WB.get_cell_contents("June Totals", "B3")
        assert contents == "='June Totals'!B1+August!A1"
        value = WB.get_cell_value("June Totals", "B3")
        assert value == Decimal(4)
