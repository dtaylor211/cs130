'''
Test Workbook

Tests the Workbook module found at ../sheets/workbook.py with more complex
operations.  Tests for basic operations can be found at test_workbook.py

Classes:
- TestWorkbook

    Methods:
    - test_rename_sheet_update_complex(object) -> None
    - test_rename_sheet_apply_quotes(object) -> None
    - test_rename_sheet_remove_quotes(object) -> None
    - test_rename_sheet_parse_error(object) -> None
    - test_move_cells_same_sheet(object) -> None
    - test_copy_cells_same_sheet(object) -> None
    - test_move_cells_overlap_basic(object) -> None
    - test_copy_cells_overlap_basic(object) -> None
    - test_move_cells_overlap_complex(object) -> None
    - test_move_cells_overlap_abs_refs(object) -> None
    - test_move_cells_overlap_mix_refs(object) -> None
    - test_copy_cells_overlap_complex(object) -> None
    - test_copy_cells_overlap_mix_refs(object) -> None
    - test_move_cells_diff_sheets(object) -> None
    - test_copy_cells_diff_sheets(object) -> None
    - test_move_cells_with_references(object) -> None
    - test_copy_cells_with_references(object) -> None
    - test_move_cells_target_oob(object) -> None
    - test_copy_cells_target_oob(object) -> None
    - test_move_copy_with_error(object) -> None

'''


from decimal import Decimal

# pylint: disable=unused-import, import-error
import context
import pytest
from sheets.workbook import Workbook
from sheets.cell_error import CellError, CellErrorType


class TestWorkbook:
    ''' 
    Workbook tests with complex operations
    
    '''

    def test_rename_sheet_update_complex(self) -> None:
        '''
        Test updating complex references on sheet rename

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Sheet5')
        wb1.new_sheet('Sheet1Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '2')
        wb1.set_cell_contents('Sheet1Sheet1', 'A1', '1.5')
        wb1.set_cell_contents('Sheet5', 'A1', '=Sheet1Sheet1!A1 * Sheet1!A1')
        wb1.rename_sheet('Sheet1', 'Sheet99')
        contents = wb1.get_cell_contents('Sheet5', 'A1')
        value = wb1.get_cell_value('Sheet5', 'A1')
        assert contents == '=Sheet1Sheet1!A1 * Sheet99!A1'
        assert value == Decimal('3')

        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', 'ayo')
        wb1.set_cell_contents('Sheet1', 'A2', '=Sheet1!A1 & \"Sheet1\"')
        wb1.set_cell_contents('Sheet1', 'A3', '=Sheet1!A1 & \"Sheet1!\"')
        wb1.rename_sheet('Sheet1', 'Sheet98')
        contents = wb1.get_cell_contents('Sheet98', 'A2')
        value = wb1.get_cell_value('Sheet98', 'A2')
        assert contents == '=Sheet98!A1 & \"Sheet1\"'
        assert value == 'ayoSheet1'
        contents = wb1.get_cell_contents('Sheet98', 'A3')
        value = wb1.get_cell_value('Sheet98', 'A3')
        assert contents == '=Sheet98!A1 & \"Sheet1!\"'
        assert value == 'ayoSheet1!'

        wb1.set_cell_contents('Sheet98', 'C1', '=September!A1+3')
        wb1.rename_sheet('Sheet99', 'September')
        contents = wb1.get_cell_contents('Sheet98', 'C1')
        value = wb1.get_cell_value('Sheet98', 'C1')
        assert contents == '=September!A1+3'
        assert value == Decimal(5)

    def test_rename_sheet_apply_quotes(self) -> None:
        '''
        Test applying quotes to a sheet name on sheet rename

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Sheet2')
        wb1.set_cell_contents('Sheet1', 'A1', 'do or do not, there is no try')
        wb1.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1')
        wb1.rename_sheet('Sheet1', 'Sheet 1')
        contents = wb1.get_cell_contents('Sheet2', 'A1')
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert contents == '=\'Sheet 1\'!A1'
        assert value == 'do or do not, there is no try'

        wb1.new_sheet('Sheet3')
        wb1.rename_sheet('Sheet 1', 'Sheet1')
        wb1.set_cell_contents('Sheet1', 'A2', '\' roger')
        wb1.set_cell_contents('Sheet3', 'A1', '- Yoda (master shiesty)')
        wb1.set_cell_contents('Sheet2', 'A2', '=Sheet1!A1 & Sheet3!A1')
        wb1.set_cell_contents('Sheet2', 'A3', '=Sheet1!A2 & Sheet1!A2')
        wb1.rename_sheet('Sheet1', 'Sheet 1')
        wb1.rename_sheet('Sheet3', 'Sheet 3')
        contents = wb1.get_cell_contents('Sheet2', 'A2')
        value = wb1.get_cell_value('Sheet2', 'A2')
        assert contents == '=\'Sheet 1\'!A1 & \'Sheet 3\'!A1'
        assert value == 'do or do not, there is no try- Yoda (master shiesty)'
        contents = wb1.get_cell_contents('Sheet2', 'A3')
        value = wb1.get_cell_value('Sheet2', 'A3')
        assert contents == '=\'Sheet 1\'!A2 & \'Sheet 1\'!A2'
        assert value == ' roger roger'

        wb1.new_sheet('Sheet4')
        wb1.new_sheet('ShEet5')
        wb1.set_cell_contents('Sheet4', 'A1', 'good relations')
        wb1.set_cell_contents('Sheet4', 'A2', '\' with the wookies,')
        wb1.set_cell_contents('Sheet5', 'A1', '\' I have')
        wb1.set_cell_contents('Sheet2', 'A1', '=Sheet4!A1 & Sheet4!A2 & Sheet5!A1')
        wb1.rename_sheet('Sheet4', 'Sheet4?')
        contents = wb1.get_cell_contents('Sheet2', 'A1')
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert contents == '=\'Sheet4?\'!A1 & \'Sheet4?\'!A2 & Sheet5!A1'
        assert value == 'good relations with the wookies, I have'

    def test_rename_sheet_remove_quotes(self) -> None:
        '''
        Test removing quotes from a sheet name on sheet rename
    
        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Benjamin Juarez')
        wb1.set_cell_contents('Benjamin Juarez', 'A1', 'i heart jar jar binks')
        wb1.set_cell_contents('Sheet1', 'A1', '=\'Benjamin Juarez\'!A1')
        wb1.rename_sheet('Benjamin Juarez', 'BJ')
        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents == '=BJ!A1'
        assert value == 'i heart jar jar binks'

        wb1.new_sheet('Kyle McGraw')
        wb1.set_cell_contents('Kyle McGraw', 'A1', 'anakin skywalker')
        wb1.set_cell_contents('Sheet1',
                             'A2', '=BJ!A1&\" and "&\'Kyle McGraw\'!A1')
        wb1.rename_sheet('Kyle McGraw', 'KM')
        contents = wb1.get_cell_contents('Sheet1', 'A2')
        value = wb1.get_cell_value('Sheet1', 'A2')
        assert contents == '=BJ!A1&\" and \"&KM!A1'
        # this is not working anymore - assigned to Dallas to figure out
        assert value == 'i heart jar jar binks and anakin skywalker'

        wb1.new_sheet('Sheet2')
        wb1.set_cell_contents('Sheet2', 'A1', '=\'KM\'!A1 & KM!A1')
        wb1.rename_sheet('KM', 'DT')
        contents = wb1.get_cell_contents('Sheet2', 'A1')
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert contents == '=DT!A1 & DT!A1'
        assert value == 'anakin skywalkeranakin skywalker'

    def test_rename_sheet_parse_error(self) -> None:
        '''
        Test parse error on rename sheet

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Sheet2')
        wb1.set_cell_contents('Sheet1', 'A1', 'Dallas Taylor')
        wb1.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1 & Sheet1!!A1')
        wb1.rename_sheet('Sheet1', 'Sheet11')
        contents = wb1.get_cell_contents('Sheet2', 'A1')
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert contents == '=Sheet1!A1 & Sheet1!!A1'
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', 'Dallas Taylor')
        wb1.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1 &&')
        wb1.rename_sheet('Sheet1', 'Sheet12')
        contents = wb1.get_cell_contents('Sheet2', 'A1')
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert contents == '=Sheet1!A1 &&'
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', 'Dallas Taylor')
        wb1.set_cell_contents('Sheet1', 'A2', '=Sheet1!A1 & \'Sheet1\'')
        wb1.rename_sheet('Sheet1', 'Sheet13')
        contents = wb1.get_cell_contents('Sheet13', 'A2')
        value = wb1.get_cell_value('Sheet13', 'A2')
        assert contents == '=Sheet1!A1 & \'Sheet1\''
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

    def test_move_cells_same_sheet(self) -> None:
        '''
        Test moving a group of cells in the same sheet

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '1')
        wb1.move_cells('Sheet1', 'A1', 'A1', 'A2')
        contents = wb1.get_cell_contents('Sheet1', 'A2')
        value = wb1.get_cell_value('Sheet1', 'A2')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents is None
        assert value is None

        wb1.move_cells('shEEt1', 'A2', 'A2', 'A3')
        contents = wb1.get_cell_contents('Sheet1', 'A3')
        value = wb1.get_cell_value('Sheet1', 'A3')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet1', 'A2')
        value = wb1.get_cell_value('Sheet1', 'A2')
        assert contents is None
        assert value is None

    def test_copy_cells_same_sheet(self) -> None:
        '''
        Test copying a group of cells in the same sheet

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '1')
        wb1.copy_cells('Sheet1', 'A1', 'A1', 'A2')
        contents = wb1.get_cell_contents('Sheet1', 'A2')
        value = wb1.get_cell_value('Sheet1', 'A2')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents == '1'
        assert value == Decimal('1')

        wb1.copy_cells('shEEt1', 'A2', 'A2', 'A3')
        contents = wb1.get_cell_contents('Sheet1', 'A3')
        value = wb1.get_cell_value('Sheet1', 'A3')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet1', 'A2')
        value = wb1.get_cell_value('Sheet1', 'A2')
        assert contents == '1'
        assert value == Decimal('1')

    def test_move_cells_overlap_basic(self) -> None:
        '''
        Test moving a group of cells in the same sheet where the source area
        and target area are overlapping

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '1')
        wb1.set_cell_contents('Sheet1', 'A2', '1')
        wb1.set_cell_contents('Sheet1', 'A3', '1')
        wb1.set_cell_contents('Sheet1', 'B1', '2')
        wb1.set_cell_contents('Sheet1', 'B2', '2')
        wb1.set_cell_contents('Sheet1', 'B3', '2')

        wb1.move_cells('Sheet1', 'A1', 'A3', 'B3')

        contents = wb1.get_cell_contents('Sheet1', 'B1')
        value = wb1.get_cell_value('Sheet1', 'B1')
        assert contents == '2'
        assert value == Decimal('2')
        contents = wb1.get_cell_contents('Sheet1', 'B2')
        value = wb1.get_cell_value('Sheet1', 'B2')
        assert contents == '2'
        assert value == Decimal('2')
        contents = wb1.get_cell_contents('Sheet1', 'B3')
        value = wb1.get_cell_value('Sheet1', 'B3')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'B4')
        value = wb1.get_cell_value('Sheet1', 'B4')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents is None
        assert value is None
        contents = wb1.get_cell_contents('Sheet1', 'A2')
        value = wb1.get_cell_value('Sheet1', 'A2')
        assert contents is None
        assert value is None
        contents = wb1.get_cell_contents('Sheet1', 'A3')
        value = wb1.get_cell_value('Sheet1', 'A3')
        assert contents is None
        assert value is None

    def test_copy_cells_overlap_basic(self) -> None:
        '''
        Test copying a group of cells in the same sheet where the source area
        and target area are overlapping

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '1')
        wb1.set_cell_contents('Sheet1', 'A2', '1')
        wb1.set_cell_contents('Sheet1', 'A3', '1')
        wb1.set_cell_contents('Sheet1', 'B1', '2')
        wb1.set_cell_contents('Sheet1', 'B2', '2')
        wb1.set_cell_contents('Sheet1', 'B3', '2')

        wb1.copy_cells('Sheet1', 'A1', 'A3', 'B2')

        contents = wb1.get_cell_contents('Sheet1', 'B1')
        value = wb1.get_cell_value('Sheet1', 'B1')
        assert contents == '2'
        assert value == Decimal('2')
        contents = wb1.get_cell_contents('Sheet1', 'B2')
        value = wb1.get_cell_value('Sheet1', 'B2')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'B3')
        value = wb1.get_cell_value('Sheet1', 'B3')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'B4')
        value = wb1.get_cell_value('Sheet1', 'B4')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'A2')
        value = wb1.get_cell_value('Sheet1', 'A2')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'A3')
        value = wb1.get_cell_value('Sheet1', 'A3')
        assert contents == '1'
        assert value == Decimal('1')

    def test_move_cells_overlap_complex(self) -> None:
        '''
        Test copying a group of cells in the same sheet where the source area
        and target area are overlapping

        '''

        wb1 = Workbook()
        # Relative references
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '1')
        wb1.set_cell_contents('Sheet1', 'B1', '=A1')
        wb1.set_cell_contents('Sheet1', 'A2', '2')
        wb1.set_cell_contents('Sheet1', 'B2', '=A2+B1')

        wb1.move_cells('Sheet1', 'A1', 'B2', 'B2')

        contents = wb1.get_cell_contents('Sheet1', 'B2')
        value = wb1.get_cell_value('Sheet1', 'B2')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'C2')
        value = wb1.get_cell_value('Sheet1', 'C2')
        assert contents == '=B2'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'B3')
        value = wb1.get_cell_value('Sheet1', 'B3')
        assert contents == '2'
        assert value == Decimal('2')
        contents = wb1.get_cell_contents('Sheet1', 'C3')
        value = wb1.get_cell_value('Sheet1', 'C3')
        assert contents == '=B3 + C2'
        assert value == Decimal('3')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents is None
        assert value is None
        contents = wb1.get_cell_contents('Sheet1', 'B1')
        value = wb1.get_cell_value('Sheet1', 'B1')
        assert contents is None
        assert value is None
        contents = wb1.get_cell_contents('Sheet1', 'A2')
        value = wb1.get_cell_value('Sheet1', 'A2')
        assert contents is None
        assert value is None

    def test_move_cells_overlap_abs_refs(self) -> None:
        '''
        Test copying a group of cells in the same sheet where the source area
        and target area are overlapping, now including absolute cell references

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet2')
        wb1.set_cell_contents('Sheet2', 'A1', '1')
        wb1.set_cell_contents('Sheet2', 'B1', '=$A$1')
        wb1.set_cell_contents('Sheet2', 'A2', '2')
        wb1.set_cell_contents('Sheet2', 'B2', '=A2+B1')

        wb1.move_cells('Sheet2', 'A1', 'B2', 'B2')

        contents = wb1.get_cell_contents('Sheet2', 'B2')
        value = wb1.get_cell_value('Sheet2', 'B2')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet2', 'C2')
        value = wb1.get_cell_value('Sheet2', 'C2')
        assert contents == '=$A$1'
        assert value == Decimal('0')
        contents = wb1.get_cell_contents('Sheet2', 'B3')
        value = wb1.get_cell_value('Sheet2', 'B3')
        assert contents == '2'
        assert value == Decimal('2')
        contents = wb1.get_cell_contents('Sheet2', 'C3')
        value = wb1.get_cell_value('Sheet2', 'C3')
        assert contents == '=B3 + C2'
        assert value == Decimal('2')

        contents = wb1.get_cell_contents('Sheet2', 'A1')
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert contents is None
        assert value is None
        contents = wb1.get_cell_contents('Sheet2', 'B1')
        value = wb1.get_cell_value('Sheet2', 'B1')
        assert contents is None
        assert value is None
        contents = wb1.get_cell_contents('Sheet2', 'A2')
        value = wb1.get_cell_value('Sheet2', 'A2')
        assert contents is None
        assert value is None

    def test_move_cells_overlap_mix_refs(self) -> None:
        '''
        Test copying a group of cells in the same sheet where the source area
        and target area are overlapping, now including mixed cell references

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet3')
        wb1.set_cell_contents('Sheet3', 'A1', '1')
        wb1.set_cell_contents('Sheet3', 'A2', '=A$1')
        wb1.set_cell_contents('Sheet3', 'B1', '2')
        wb1.set_cell_contents('Sheet3', 'B2', '=A2+B1')

        wb1.move_cells('Sheet3', 'A1', 'B2', 'B2')

        contents = wb1.get_cell_contents('Sheet3', 'B2')
        value = wb1.get_cell_value('Sheet3', 'B2')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet3', 'C2')
        value = wb1.get_cell_value('Sheet3', 'C2')
        assert contents == '2'
        assert value == Decimal('2')
        contents = wb1.get_cell_contents('Sheet3', 'B3')
        value = wb1.get_cell_value('Sheet3', 'B3')
        assert contents == '=B$1'
        assert value == Decimal('0')
        contents = wb1.get_cell_contents('Sheet3', 'C3')
        value = wb1.get_cell_value('Sheet3', 'C3')
        assert contents == '=B3 + C2'
        assert value == Decimal('2')

        contents = wb1.get_cell_contents('Sheet3', 'A1')
        value = wb1.get_cell_value('Sheet3', 'A1')
        assert contents is None
        assert value is None
        contents = wb1.get_cell_contents('Sheet3', 'B1')
        value = wb1.get_cell_value('Sheet3', 'B1')
        assert contents is None
        assert value is None
        contents = wb1.get_cell_contents('Sheet3', 'A2')
        value = wb1.get_cell_value('Sheet3', 'A2')
        assert contents is None
        assert value is None

    def test_copy_cells_overlap_complex(self) -> None:
        '''
        Test copying a group of cells in the same sheet where the source area
        and target area are overlapping

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '=A2+B1')
        wb1.set_cell_contents('Sheet1', 'B1', '=B2')
        wb1.set_cell_contents('Sheet1', 'A2', '2')
        wb1.set_cell_contents('Sheet1', 'B2', '1')

        wb1.copy_cells('Sheet1', 'A1', 'B2', 'B2')

        contents = wb1.get_cell_contents('Sheet1', 'B2')
        value = wb1.get_cell_value('Sheet1', 'B2')
        assert contents == '=B3 + C2'
        assert value == Decimal('3')
        contents = wb1.get_cell_contents('Sheet1', 'C2')
        value = wb1.get_cell_value('Sheet1', 'C2')
        assert contents == '=C3'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'B3')
        value = wb1.get_cell_value('Sheet1', 'B3')
        assert contents == '2'
        assert value == Decimal('2')
        contents = wb1.get_cell_contents('Sheet1', 'C3')
        value = wb1.get_cell_value('Sheet1', 'C3')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents == "=A2+B1"
        assert value == Decimal('5')
        contents = wb1.get_cell_contents('Sheet1', 'B1')
        value = wb1.get_cell_value('Sheet1', 'B1')
        assert contents == "=B2"
        assert value == Decimal('3')
        contents = wb1.get_cell_contents('Sheet1', 'A2')
        value = wb1.get_cell_value('Sheet1', 'A2')
        assert contents == "2"
        assert value == Decimal('2')

    def test_copy_cells_overlap_mix_refs(self) -> None:
        '''
        Test copying a group of cells in the same sheet where the source area
        and target area are overlapping, now with mixed cell references

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet2')
        wb1.set_cell_contents('Sheet2', 'A1', '=A2+B1')
        wb1.set_cell_contents('Sheet2', 'B1', '=$B2')
        wb1.set_cell_contents('Sheet2', 'A2', '2')
        wb1.set_cell_contents('Sheet2', 'B2', '1')

        wb1.copy_cells('Sheet2', 'A1', 'B2', 'B2')

        contents = wb1.get_cell_contents('Sheet2', 'B2')
        value = wb1.get_cell_value('Sheet2', 'B2')
        assert contents == '=B3 + C2'
        assert value == Decimal('4')
        contents = wb1.get_cell_contents('Sheet2', 'C2')
        value = wb1.get_cell_value('Sheet2', 'C2')
        assert contents == '=$B3'
        assert value == Decimal('2')
        contents = wb1.get_cell_contents('Sheet2', 'B3')
        value = wb1.get_cell_value('Sheet2', 'B3')
        assert contents == '2'
        assert value == Decimal('2')
        contents = wb1.get_cell_contents('Sheet2', 'C3')
        value = wb1.get_cell_value('Sheet2', 'C3')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet2', 'A1')
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert contents == "=A2+B1"
        assert value == Decimal('6')
        contents = wb1.get_cell_contents('Sheet2', 'B1')
        value = wb1.get_cell_value('Sheet2', 'B1')
        assert contents == "=$B2"
        assert value == Decimal('4')
        contents = wb1.get_cell_contents('Sheet2', 'A2')
        value = wb1.get_cell_value('Sheet2', 'A2')
        assert contents == "2"
        assert value == Decimal('2')

    def test_move_cells_diff_sheets(self) -> None:
        '''
        Test moving a group of cells from one sheet to another sheet

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Sheet2')
        wb1.set_cell_contents('Sheet1', 'A1', '1')
        wb1.move_cells('Sheet1', 'A1', 'A1', 'A2', "Sheet2")
        contents = wb1.get_cell_contents('Sheet2', 'A2')
        value = wb1.get_cell_value('Sheet2', 'A2')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents is None
        assert value is None

        wb1.new_sheet('Sheet3')
        wb1.move_cells('shEEt2', 'A2', 'A2', 'A3', 'Sheet3')
        contents = wb1.get_cell_contents('Sheet3', 'A3')
        value = wb1.get_cell_value('Sheet3', 'A3')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet2', 'A2')
        value = wb1.get_cell_value('Sheet2', 'A2')
        assert contents is None
        assert value is None

        wb1.set_cell_contents('Sheet1', 'A2', '5')
        wb1.set_cell_contents('Sheet2', 'B2', '4')
        wb1.set_cell_contents('Sheet2', 'B4', '2')
        wb1.set_cell_contents('Sheet2', 'B3', '=(Sheet2!$B$2 / Sheet3!B3) + B4')
        wb1.set_cell_contents('Sheet3', 'A1', '2')
        wb1.set_cell_contents('Sheet3', 'B3', '1')
        value = wb1.get_cell_value('Sheet2', 'B3')
        assert value == Decimal('6')

        wb1.move_cells('Sheet2', 'B3', 'B3', 'A1', 'Sheet1')
        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents == '=(Sheet2!$B$2 / Sheet3!A1) + A2'
        assert value == Decimal('7')

    def test_copy_cells_diff_sheets(self) -> None:
        '''
        Test copying a group of cells from one sheet to another sheet

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Sheet2')
        wb1.set_cell_contents('Sheet1', 'A1', '1')
        wb1.copy_cells('Sheet1', 'A1', 'A1', 'A2', 'Sheet2')
        contents = wb1.get_cell_contents('Sheet2', 'A2')
        value = wb1.get_cell_value('Sheet2', 'A2')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents == '1'
        assert value == Decimal('1')

        wb1.new_sheet('Sheet3')
        wb1.copy_cells('shEEt2', 'A2', 'A2', 'A3', 'Sheet3')
        contents = wb1.get_cell_contents('Sheet3', 'A3')
        value = wb1.get_cell_value('Sheet3', 'A3')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'A3')
        value = wb1.get_cell_value('Sheet1', 'A3')
        assert contents is None
        assert value is None

        contents = wb1.get_cell_contents('Sheet2', 'A2')
        value = wb1.get_cell_value('Sheet2', 'A2')
        assert contents == '1'
        assert value == Decimal('1')

    def test_move_cells_with_references(self) -> None:
        '''
        Test moving a group of cells where contents involve formulas/refs

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A50', '=50')
        wb1.set_cell_contents('Sheet1', 'B51', '=60')
        wb1.set_cell_contents('Sheet1', 'C52', '=70')
        wb1.set_cell_contents('Sheet1', 'A1', '=A50+B51')
        wb1.move_cells('Sheet1', 'A1', 'A1', 'B2')
        contents = wb1.get_cell_contents('Sheet1', 'B2')
        value = wb1.get_cell_value('Sheet1', 'B2')
        assert contents == '=B51 + C52'
        assert value == Decimal('130')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents is None
        assert value is None

        wb1.move_cells('shEEt1', 'B2', 'B2', 'C3')
        contents = wb1.get_cell_contents('Sheet1', 'C3')
        value = wb1.get_cell_value('Sheet1', 'C3')
        assert contents == '=C52 + D53'
        assert value == Decimal('70')

        wb1.new_sheet('Sheet2')
        wb1.set_cell_contents('Sheet2', 'A1', 'test')
        wb1.set_cell_contents('Sheet2', 'B1', 'test2')
        wb1.set_cell_contents('Sheet2', 'A2', '=\'Sheet1\'!A$1 & "pass!A1"')
        wb1.move_cells('Sheet2', 'A1', 'A2', 'B2')
        contents = wb1.get_cell_contents('Sheet2', 'B3')
        value = wb1.get_cell_value('Sheet2', 'B3')
        assert contents ==  '=\'Sheet1\'!B$1 & "pass!A1"'
        assert value == 'pass!A1'

        # test moving up and to left
        wb1.new_sheet('Sheet3')
        wb1.set_cell_contents('Sheet3', 'A2', '1')
        wb1.set_cell_contents('Sheet3', 'A3', '1')
        wb1.set_cell_contents('Sheet3', 'B2', '2')
        wb1.set_cell_contents('Sheet3', 'B3', '4')
        wb1.set_cell_contents('Sheet3', 'C2', '3')
        wb1.set_cell_contents('Sheet3', 'C3', '=B2+$C2+B3')
        wb1.set_cell_contents('Sheet3', 'D2', '1')
        wb1.move_cells('Sheet3', 'B2', 'D3', 'A1')

        contents = wb1.get_cell_contents('Sheet3', 'A1')
        value = wb1.get_cell_value('Sheet3', 'A1')
        assert contents == '2'
        assert value == Decimal('2')

        contents = wb1.get_cell_contents('Sheet3', 'B2')
        value = wb1.get_cell_value('Sheet3', 'B2')
        assert contents == '=A1 + $C1 + A2'
        assert value == Decimal('7')

        contents = wb1.get_cell_contents('Sheet3', 'A3')
        value = wb1.get_cell_value('Sheet3', 'A3')
        assert value == Decimal('1')


    def test_copy_cells_with_references(self) -> None:
        '''
        Test copying a group of cells where contents involve formulas/refs

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A50', '=50')
        wb1.set_cell_contents('Sheet1', 'B51', '=60')
        wb1.set_cell_contents('Sheet1', 'C52', '=70')
        wb1.set_cell_contents('Sheet1', 'A1', '=A50+B51')
        wb1.copy_cells('Sheet1', 'A1', 'A1', 'B2')
        contents = wb1.get_cell_contents('Sheet1', 'B2')
        value = wb1.get_cell_value('Sheet1', 'B2')
        assert contents == '=B51 + C52'
        assert value == Decimal('130')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents == '=A50+B51'
        assert value == Decimal('110')

        wb1.copy_cells('shEEt1', 'B2', 'B2', 'C3')
        contents = wb1.get_cell_contents('Sheet1', 'C3')
        value = wb1.get_cell_value('Sheet1', 'C3')
        assert contents == '=C52 + D53'
        assert value == Decimal('70')

        contents = wb1.get_cell_contents('Sheet1', 'B2')
        value = wb1.get_cell_value('Sheet1', 'B2')
        assert contents == '=B51 + C52'
        assert value == Decimal('130')

        wb1.new_sheet('Sheet2')
        wb1.set_cell_contents('Sheet1', 'B1', 'test2')
        wb1.set_cell_contents('Sheet2', 'A2', '=\'Sheet1\'!A$1 & "pass!A1"')
        wb1.copy_cells('Sheet2', 'A1', 'A2', 'B2')
        contents = wb1.get_cell_contents('Sheet2', 'B3')
        value = wb1.get_cell_value('Sheet2', 'B3')
        assert contents ==  '=\'Sheet1\'!B$1 & "pass!A1"'
        assert value == 'test2pass!A1'

        wb1.new_sheet('Sheet3')
        wb1.set_cell_contents('Sheet3', 'A2', '1')
        wb1.set_cell_contents('Sheet3', 'B2', '2')
        wb1.set_cell_contents('Sheet3', 'B3', '4')
        wb1.set_cell_contents('Sheet3', 'C3', '=B2+$C2+B3')
        wb1.set_cell_contents('Sheet3', 'D2', '1')
        wb1.copy_cells('Sheet3', 'B2', 'D3', 'A1')

        contents = wb1.get_cell_contents('Sheet3', 'A1')
        assert contents == '2'

        contents = wb1.get_cell_contents('Sheet3', 'B2')
        value = wb1.get_cell_value('Sheet3', 'B2')
        assert contents == '=A1 + $C1 + A2'
        assert value == Decimal('7')

        contents = wb1.get_cell_contents('Sheet3', 'C3')
        value = wb1.get_cell_value('Sheet3', 'C3')
        assert contents == '=B2+$C2+B3'
        assert value == Decimal('11')

    def test_move_cells_target_oob(self) -> None:
        '''
        Test moving a group of cells where the target area would extend outside
        the valid area of the spreadsheet (no changes should be made)

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '1')
        wb1.set_cell_contents('Sheet1', 'B1', '1')
        wb1.set_cell_contents('Sheet1', 'A2', '1')
        wb1.set_cell_contents('Sheet1', 'B2', '1')

        with pytest.raises(ValueError):
            wb1.move_cells('Sheet1', 'A1', 'B2', 'A9999')

        contents = wb1.get_cell_contents('Sheet1', 'A1')
        value = wb1.get_cell_value('Sheet1', 'A1')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'B1')
        value = wb1.get_cell_value('Sheet1', 'B1')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'A2')
        value = wb1.get_cell_value('Sheet1', 'A2')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'B2')
        value = wb1.get_cell_value('Sheet1', 'B2')
        assert contents == '1'
        assert value == Decimal('1')
        contents = wb1.get_cell_contents('Sheet1', 'A9999')
        value = wb1.get_cell_value('Sheet1', 'A9999')
        assert contents is None
        assert value is None

    def test_copy_cells_target_oob(self) -> None:
        '''
        Test copying a group of cells where the target area would extend outside
        the valid area of the spreadsheet (no changes should be made)

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '1')
        wb1.set_cell_contents('Sheet1', 'B1', '1')
        wb1.set_cell_contents('Sheet1', 'A2', '1')
        wb1.set_cell_contents('Sheet1', 'B2', '1')

        with pytest.raises(ValueError):
            wb1.copy_cells('Sheet1', 'A1', 'B2', 'A9999')

        contents = wb1.get_cell_contents('Sheet1', 'A9999')
        value = wb1.get_cell_value('Sheet1', 'A9999')
        assert contents is None
        assert value is None

    def test_move_copy_with_error(self) -> None:
        '''
        Test moving/copying a group of cells with a parse error
        or value error

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '1')
        wb1.set_cell_contents('Sheet1', 'B2', '=Sheet1!A1 + SHH!SHH')
        wb1.move_cells('Sheet1', 'A1', 'B2', 'C1')
        contents = wb1.get_cell_contents('Sheet1', 'C1')
        value = wb1.get_cell_value('Sheet1', 'C1')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet1', 'D2')
        value = wb1.get_cell_value('Sheet1', 'D2')
        assert contents == '=Sheet1!A1 + SHH!SHH'
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

        wb1.new_sheet('Sheet2')
        wb1.copy_cells('Sheet1', 'C1', 'D2', 'A1', 'Sheet2')
        contents = wb1.get_cell_contents('Sheet2', 'A1')
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert contents == '1'
        assert value == Decimal('1')

        contents = wb1.get_cell_contents('Sheet2', 'B2')
        value = wb1.get_cell_value('Sheet2', 'B2')
        assert contents == '=Sheet1!A1 + SHH!SHH'
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

        wb1.set_cell_contents('Sheet2', 'A1', '1')
        wb1.set_cell_contents('Sheet2', 'A2', '2')
        wb1.set_cell_contents('Sheet2', 'A3', '=A1 + A2')
        wb1.move_cells('Sheet2', 'A3', 'A3', 'B2')
        contents = wb1.get_cell_contents('Sheet2', 'B2')
        value = wb1.get_cell_value('Sheet2', 'B2')
        assert contents == '=#REF! + B1'
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.BAD_REFERENCE

        wb1.set_cell_contents('Sheet2', 'A1', '=INDIRECT(B1)')
        wb1.move_cells('Sheet2', 'A1', 'A1', 'ZZZZ1')
        contents = wb1.get_cell_contents('Sheet2', 'ZZZZ1')
        value = wb1.get_cell_value('Sheet2', 'ZZZZ1')
        assert contents == '=INDIRECT(#REF!)'
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.BAD_REFERENCE
