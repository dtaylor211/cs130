'''
Test Workbook

Tests the Workbook module found at ../sheets/workbook.py with more complex
operations.  Tests for basic operations can be found at test_workbook.py

Classes:
- TestWorkbook

    Methods:
    - test_load_workbook(object) -> None
    - test_save_workbook(object) -> None
    - test_notify_cell(object) -> None
    - test_rename_sheet(object) -> None
    - test_rename_sheet_update(object) -> None
    - test_rename_sheet_update_complex(object) -> None
    - test_rename_sheet_apply_quotes(object) -> None
    - test_rename_sheet_remove_quotes(object) -> None
    - test_rename_sheet_parse_error(object) -> None
    - test_move_sheet(object) -> None
    - test_copy_sheet(object) -> None
    - test_mutate_returned_attributes(object) -> None
    - test_move_cells_same_sheet(object) -> None
    - test_move_copy_with_error(object) -> None

'''


from decimal import Decimal
import io
import json

# pylint: disable=unused-import, import-error
import context
import pytest
from sheets.workbook import Workbook
from sheets.cell_error import CellError, CellErrorType


class TestWorkbook:
    ''' 
    Workbook tests with complex operations
    
    '''

    def test_load_workbook(self) -> None:
        '''
        Test loading a workbook

        '''

        with open('tests/json_data/wb_data_valid.json', encoding="utf8") as fp:
            wb1 = Workbook.load_workbook(fp)
            assert wb1.num_sheets() == 2
            assert wb1.list_sheets() == ['Sheet1', 'Sheet2']
            assert wb1.get_cell_contents('Sheet1', 'A1') == '\'123'
            assert wb1.get_cell_contents('Sheet1', 'B1') == '5.3'
            assert wb1.get_cell_contents('Sheet1', 'C1') == '=A1*B1'
            assert wb1.get_cell_value('Sheet1', 'A1') == '123'
            assert wb1.get_cell_value('Sheet1', 'B1') == Decimal('5.3')
            assert wb1.get_cell_value('Sheet1', 'C1') == Decimal('651.9')

        with open('tests/json_data/wb_data_invalid_dup.json',
            encoding="utf8") as fp:
            with pytest.raises(ValueError):
                wb1 = Workbook.load_workbook(fp)

        with open('tests/json_data/wb_data_missing_sheets.json',
            encoding="utf8") as fp:
            with pytest.raises(KeyError):
                wb1 = Workbook.load_workbook(fp)
        with open('tests/json_data/wb_data_missing_name.json',
            encoding="utf8") as fp:
            with pytest.raises(KeyError):
                wb1 = Workbook.load_workbook(fp)
        with open('tests/json_data/wb_data_missing_contents.json',
            encoding="utf8") as fp:
            with pytest.raises(KeyError):
                wb1 = Workbook.load_workbook(fp)

        with open('tests/json_data/wb_data_bad_type_sheets.json',
            encoding="utf8") as fp:
            with pytest.raises(TypeError):
                wb1 = Workbook.load_workbook(fp)
        with open('tests/json_data/wb_data_bad_type_sheet.json',
            encoding="utf8") as fp:
            with pytest.raises(TypeError):
                wb1 = Workbook.load_workbook(fp)
        with open('tests/json_data/wb_data_bad_type_name.json',
            encoding="utf8") as fp:
            with pytest.raises(TypeError):
                wb1 = Workbook.load_workbook(fp)
        with open('tests/json_data/wb_data_bad_type_contents.json',
            encoding="utf8") as fp:
            with pytest.raises(TypeError):
                wb1 = Workbook.load_workbook(fp)
        with open('tests/json_data/wb_data_bad_type_cell_contents.json',
            encoding="utf8") as fp:
            with pytest.raises(TypeError):
                wb1 = Workbook.load_workbook(fp)

        with open('tests/json_data/wb_data_bad_type_location.json',
            encoding="utf8") as fp:
            with pytest.raises(json.JSONDecodeError):
                wb1 = Workbook.load_workbook(fp)

    def test_save_workbook(self) -> None:
        '''
        Test saving a workbook

        '''

        with io.StringIO('') as fp:
            wb1 = Workbook()
            wb1.new_sheet('Sheet1')
            wb1.new_sheet('Sheet2')
            wb1.set_cell_contents('Sheet1', 'A1', '1')
            wb1.set_cell_contents('Sheet2', 'B2', '2')
            wb1.save_workbook(fp)
            fp.seek(0)
            json_act = json.load(fp)
            json_exp = {
                'sheets':[
                    {
                        'name':'Sheet1',
                        'cell-contents':{
                            'A1':'1'
                        }
                    },
                    {
                        'name':'Sheet2',
                        'cell-contents':{
                            'B2':'2'
                        }
                    }
                ]
            }
            assert json_act == json_exp

        with io.StringIO('') as fp:
            wb1 = Workbook()
            wb1.new_sheet('Sheet1')
            wb1.set_cell_contents('Sheet1', 'A1', '1')
            wb1.new_sheet('Sheet2')
            wb1.set_cell_contents('Sheet2', 'B2', '=Sheet1!A1')
            wb1.save_workbook(fp)
            fp.seek(0)
            json_act = json.load(fp)
            json_exp = {
                'sheets':[
                    {
                        'name':'Sheet1',
                        'cell-contents':{
                            'A1':'1'
                        }
                    },
                    {
                        'name':'Sheet2',
                        'cell-contents':{
                            'B2':'=Sheet1!A1'
                        }
                    }
                ]
            }
            assert json_act == json_exp

    def test_notify_cell(self) -> None:
        '''
        Test cell notifications
        '''

        test_changed = []
        def on_cells_changed(_, changed_cells):
            '''
            This function gets called when cells change in the workbook that the
            function was registered on.  The changed_cells argument is an iterable
            of tuples; each tuple is of the form (sheet_name, cell_location).
            '''
            test_changed.append(changed_cells)
        wb1 = Workbook()
        wb1.notify_cells_changed(on_cells_changed)
        wb1.new_sheet('Sheet1')
        assert test_changed[-1] == []
        wb1.set_cell_contents('Sheet1', 'A1', '\'123')
        assert test_changed[-1] == [('Sheet1', 'A1')]
        wb1.set_cell_contents('Sheet1', 'C1', '=A1+B1')
        assert test_changed[-1] == [('Sheet1', 'C1')]
        wb1.set_cell_contents('Sheet1', 'B1', '5.3')
        assert test_changed[-1] == [('Sheet1', 'B1'), ('Sheet1', 'C1')]
        wb1.move_cells('Sheet1', 'C1', 'C1', 'C2')
        assert test_changed[-2] == [('Sheet1', 'C1')]
        assert test_changed[-1] == [('Sheet1', 'C2')]
        wb1.set_cell_contents('Sheet1', 'C2', None)
        assert test_changed[-1] == [('Sheet1', 'C2')]
        wb1.del_sheet('Sheet1')
        assert test_changed[-1] == []
        def on_cells_changed2(workbook, changed_cells):
            '''
            This function gets called when cells change in the workbook that the
            function was registered on.  The changed_cells argument is an iterable
            of tuples; each tuple is of the form (sheet_name, cell_location).
            '''
            test_changed.append([workbook.get_cell_value(sheet, cell) for
                sheet, cell in changed_cells])
        wb1.new_sheet('Test')
        wb1.notify_cells_changed(on_cells_changed2)
        wb1.set_cell_contents('Test', 'A1', '1')
        assert test_changed[-2] == [('Test', 'A1')]
        assert test_changed[-1] == [Decimal(1)]
        wb1.set_cell_contents('Test', 'B1', '=A1')
        assert test_changed[-2] == [('Test', 'B1')]
        assert test_changed[-1] == [Decimal(1)]
        wb1.set_cell_contents('Test', 'C1', '=A1+B1')
        assert test_changed[-2] == [('Test', 'C1')]
        assert test_changed[-1] == [Decimal(2)]
        wb1.set_cell_contents('Test', 'A1', '2')
        assert test_changed[-2] == [('Test', 'A1'), ('Test', 'B1'),
                                    ('Test', 'C1')]
        assert test_changed[-1] == [Decimal(2), Decimal(2), Decimal(4)]
        def on_cells_changed3(workbook, changed_cells):
            '''
            This function gets called when cells change in the workbook that the
            function was registered on.  The changed_cells argument is an iterable
            of tuples; each tuple is of the form (sheet_name, cell_location).
            '''
            raise ValueError('Whaaaat were raising a random error')
        wb1.notify_cells_changed(on_cells_changed3)
        wb1.set_cell_contents('Test', 'C1', '4')
        assert test_changed[-2] == []
        assert test_changed[-1] == []

    def test_rename_sheet(self) -> None:
        '''
        Test renaming a sheet

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '2')
        wb1.set_cell_contents('Sheet1', 'A2', '=A1')
        wb1.rename_sheet('Sheet1', 'SheEt2')
        sheet_names = wb1.list_sheets()
        sheet_objects = wb1.get_sheet_objects()
        assert sheet_names == ['SheEt2']
        assert sheet_objects['sheet2'] is not None
        with pytest.raises(KeyError):
            assert not sheet_objects['sheet1']
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert value == Decimal(2)
        value = wb1.get_cell_value('Sheet2','A2')
        assert value == Decimal(2)

        wb1.new_sheet('Sheet3')
        wb1.rename_sheet('Sheet2', 'Sheet4')
        new_sheet_names = wb1.list_sheets()
        new_sheet_objects = wb1.get_sheet_objects()
        assert new_sheet_names == ['Sheet4', 'Sheet3']
        assert new_sheet_objects['sheet4'] is not None
        assert new_sheet_objects['sheet3'] is not None
        with pytest.raises(KeyError):
            assert not new_sheet_objects['sheet2']
        value = wb1.get_cell_value('Sheet4', 'A1')
        assert value == Decimal(2)

    def test_rename_sheet_update(self) -> None:
        '''
        Test updating simple references on sheet rename

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Sheet2')
        wb1.set_cell_contents('Sheet1', 'A1', '=2')
        wb1.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1')
        wb1.rename_sheet('Sheet1', 'Sheet3')
        contents = wb1.get_cell_contents('Sheet2', 'A1')
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert contents == '=Sheet3!A1'
        assert value == Decimal(2)

        wb1.set_cell_contents('Sheet3', 'A2', '=3')
        wb1.set_cell_contents('Sheet3', 'A3', '=4')
        wb1.set_cell_contents('Sheet2', 'A2', '=Sheet3!A2')
        wb1.set_cell_contents('Sheet2', 'A3', '=Sheet3!A3')
        wb1.rename_sheet('Sheet3', 'Sheet4')
        contents = wb1.get_cell_contents('Sheet2', 'A1')
        value = wb1.get_cell_value('Sheet2', 'A1')
        assert contents == '=Sheet4!A1'
        assert value == Decimal(2)
        contents = wb1.get_cell_contents('Sheet2', 'A2')
        value = wb1.get_cell_value('Sheet2', 'A2')
        assert contents == '=Sheet4!A2'
        assert value == Decimal(3)
        contents = wb1.get_cell_contents('Sheet2', 'A3')
        value = wb1.get_cell_value('Sheet2', 'A3')
        assert contents == '=Sheet4!A3'
        assert value == Decimal(4)

        wb1.new_sheet('Sheet5')
        wb1.set_cell_contents('Sheet5', 'A1', '=Sheet2!A2 + Sheet2!A3')
        wb1.set_cell_contents('Sheet5', 'A2', '=Sheet2!A2 + Sheet4!A3')
        wb1.rename_sheet('Sheet2', 'Sheet6')
        wb1.rename_sheet('Sheet4', 'Sheet7')
        contents = wb1.get_cell_contents('Sheet5', 'A1')
        value = wb1.get_cell_value('Sheet5', 'A1')
        assert contents == '=Sheet6!A2 + Sheet6!A3'
        assert value == Decimal(7)
        contents = wb1.get_cell_contents('Sheet5', 'A2')
        value = wb1.get_cell_value('Sheet5', 'A2')
        assert contents == '=Sheet6!A2 + Sheet7!A3'
        assert value == Decimal(7)

        wb1.new_sheet('A Sheet')
        wb1.set_cell_contents('A Sheet', 'A1', '=0.1')
        wb1.set_cell_contents('Sheet5', 'A1', '=\'A Sheet\'!A1')
        wb1.rename_sheet('A Sheet', 'Darth Jar Jar')
        contents = wb1.get_cell_contents('Sheet5', 'A1')
        value = wb1.get_cell_value('Sheet5', 'A1')
        assert contents == '=\'Darth Jar Jar\'!A1'
        assert value == Decimal('0.1')

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

    def test_move_sheet(self) -> None:
        '''
        Test moving a sheet

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Sheet2')
        wb1.new_sheet('Sheet3')
        with pytest.raises(KeyError):
            wb1.move_sheet('Sheet4', 0)
        with pytest.raises(IndexError):
            wb1.move_sheet('Sheet3', -1)
        with pytest.raises(IndexError):
            wb1.move_sheet('Sheet3', 4)
        wb1.move_sheet('Sheet3', 0)
        assert wb1.list_sheets() == ['Sheet3', 'Sheet1', 'Sheet2']
        wb1.move_sheet('Sheet3', 2)
        assert wb1.list_sheets() == ['Sheet1', 'Sheet2', 'Sheet3']

    def test_copy_sheet(self) -> None:
        '''
        Test copying a sheet

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.set_cell_contents('Sheet1', 'A1', '=1')
        wb1.copy_sheet('Sheet1')
        wb1.copy_sheet('Sheet1')
        assert wb1.num_sheets() == 3
        assert wb1.list_sheets() == ['Sheet1', 'Sheet1_1', 'Sheet1_2']

        wb1.copy_sheet('Sheet1_2')
        assert wb1.list_sheets() == ['Sheet1', 'Sheet1_1', 'Sheet1_2',
            'Sheet1_2_1']

        assert wb1.get_cell_value('Sheet1', 'A1') == Decimal('1')
        assert wb1.get_cell_value('Sheet1_1', 'A1') == Decimal('1')
        assert wb1.get_cell_value('Sheet1_2', 'A1') == Decimal('1')
        assert wb1.get_cell_value('Sheet1_2_1', 'A1') == Decimal('1')

        wb1.set_cell_contents('Sheet1', 'A1', '=2')
        assert wb1.get_cell_value('Sheet1', 'A1') == Decimal('2')
        assert wb1.get_cell_value('Sheet1_1', 'A1') == Decimal('1')
        assert wb1.get_cell_value('Sheet1_2', 'A1') == Decimal('1')
        assert wb1.get_cell_value('Sheet1_2_1', 'A1') == Decimal('1')

        wb1.set_cell_contents('Sheet1_2', 'A1', '=3')
        assert wb1.get_cell_value('Sheet1', 'A1') == Decimal('2')
        assert wb1.get_cell_value('Sheet1_1', 'A1') == Decimal('1')
        assert wb1.get_cell_value('Sheet1_2', 'A1') == Decimal('3')
        assert wb1.get_cell_value('Sheet1_2_1', 'A1') == Decimal('1')

        wb1.new_sheet('Sheet2')
        wb1.new_sheet('Sheet3')
        wb1.set_cell_contents('Sheet2', 'B2', '=2')
        wb1.set_cell_contents('Sheet3', 'B2', '=Sheet2!B2+Sheet2!B2')
        wb1.copy_sheet('Sheet2')
        wb1.copy_sheet('Sheet3')
        wb1.copy_sheet('Sheet2')

        assert wb1.list_sheets() == ['Sheet1', 'Sheet1_1', 'Sheet1_2',
            'Sheet1_2_1', 'Sheet2', 'Sheet3', 'Sheet2_1', 'Sheet3_1',
            'Sheet2_2']
        assert wb1.get_cell_value('Sheet2_1', 'B2') == Decimal('2')
        assert wb1.get_cell_value('Sheet3_1', 'B2') == Decimal('4')

        wb1.new_sheet('Sheet4')
        wb1.set_cell_contents('Sheet4', 'D4', '=#CIRCREF!')
        wb1.copy_sheet('Sheet4')
        contents = wb1.get_cell_contents('Sheet4_1', 'D4')
        value = wb1.get_cell_value('Sheet4_1', 'D4')
        assert contents == '=#CIRCREF!'
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

        (_, name) = wb1.copy_sheet('Sheet4')
        assert name == 'Sheet4_2'
        (_, name) = wb1.copy_sheet('Sheet4')
        assert name == 'Sheet4_3'

    def test_mutate_returned_attributes(self) -> None:
        '''
        Test mutating returned attributes

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Sheet2')
        sheet_names = wb1.list_sheets()
        sheet_names.remove('Sheet1')
        sheet_names = wb1.list_sheets()
        assert sheet_names == ['Sheet1', 'Sheet2']

        sheet_objects = wb1.get_sheet_objects()
        del sheet_objects['sheet1']
        new_sheet_objects = wb1.get_sheet_objects()
        assert new_sheet_objects['sheet1'] is not None

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

        # move below test to one with cell references, just using to check work
        wb1.set_cell_contents('Sheet1', 'A1', 'test')
        wb1.set_cell_contents('Sheet1', 'B1', 'test2')
        wb1.set_cell_contents('Sheet1', 'A2', '=\'Sheet1\'!A$1 & "pass!A1"')
        wb1.move_cells('Sheet1', 'A1', 'A2', 'B2')
        contents = wb1.get_cell_contents('Sheet1', 'B3')
        value = wb1.get_cell_value('Sheet1', 'B3')
        assert contents ==  '=\'Sheet1\'!B$1 & "pass!A1"'
        assert value == 'test2pass!A1'

        # add some tests with cell references

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
