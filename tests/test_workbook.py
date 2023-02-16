'''
Test Workbook

Tests the Workbook module found at ../sheets/workbook.py.

Classes:
- TestWorkbook

    Methods:
    - test_empty_workbook(object) -> None
    - test_new_sheet_name(object) -> None
    - test_new_sheet_complex(object) -> None
    - test_del_sheet(object) -> None
    - test_set_and_extent(object) -> None
    - test_get_contents_and_value(object) -> None
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
    Workbook tests
    
    '''

    def test_empty_workbook(self) -> None:
        '''
        Test empty workbook

        '''

        wb1 = Workbook()
        assert wb1.num_sheets() == 0
        assert wb1.list_sheets() == []

    def test_new_sheet_name(self) -> None:
        '''
        Test new sheet names

        '''

        wb1 = Workbook()
        index, name = wb1.new_sheet()
        assert index == 0
        assert name == 'Sheet1'
        assert wb1.num_sheets() == 1
        assert wb1.list_sheets() == ['Sheet1']

        index, name = wb1.new_sheet()
        assert index == 1
        assert name == 'Sheet2'
        assert wb1.num_sheets() == 2
        assert wb1.list_sheets() == ['Sheet1', 'Sheet2']

        wb1 = Workbook()
        index, name = wb1.new_sheet('July Totals')
        assert index == 0
        assert name == 'July Totals'
        assert wb1.num_sheets() == 1
        assert wb1.list_sheets() == ['July Totals']

        index, name = wb1.new_sheet('June Totals')
        assert index == 1
        assert name == 'June Totals'
        assert wb1.num_sheets() == 2
        assert wb1.list_sheets() == ['July Totals', 'June Totals']

        wb1 = Workbook()
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('')
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('   July Totals')
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('July Totals  ')
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('July Totals >')
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('July\'s Totals')

        index, name = wb1.new_sheet()
        assert name == 'Sheet1'
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('Sheet1')

    def test_new_sheet_complex(self) -> None:
        '''
        Test new sheet with more complex checks to adding new fields

        '''

        wb1 = Workbook()
        index, name = wb1.new_sheet('Sheet2')
        assert index == 0
        assert name == 'Sheet2'
        assert wb1.num_sheets() == 1
        assert wb1.list_sheets() == ['Sheet2']

        index, name = wb1.new_sheet('July Totals')
        assert index == 1
        assert name == 'July Totals'
        assert wb1.num_sheets() == 2
        assert wb1.list_sheets() == ['Sheet2', 'July Totals']

        index, name = wb1.new_sheet()
        assert index == 2
        assert name == 'Sheet1'
        assert wb1.num_sheets() == 3
        assert wb1.list_sheets() == ['Sheet2', 'July Totals', 'Sheet1']

        index, name = wb1.new_sheet()
        assert index == 3
        assert name == 'Sheet3'
        assert wb1.num_sheets() == 4
        assert wb1.list_sheets() == ['Sheet2', 'July Totals', 'Sheet1', 'Sheet3']

    def test_del_sheet(self) -> None:
        '''
        Test deleting a sheet

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Sheet2')
        assert wb1.num_sheets() == 2
        assert wb1.list_sheets() == ['Sheet1', 'Sheet2']

        with pytest.raises(KeyError):
            wb1.del_sheet('July Totals')

        wb1.del_sheet('Sheet1')
        assert wb1.num_sheets() == 1
        assert wb1.list_sheets() == ['Sheet2']

        wb1.del_sheet('shEet2')
        assert wb1.num_sheets() == 0
        assert wb1.list_sheets() == []

    def test_set_and_extent(self) -> None:
        '''
        Test setting sheet and getting extent

        '''

        wb1 = Workbook()
        with pytest.raises(KeyError):
            wb1.get_sheet_extent('Sheet1')

        (_, name) = wb1.new_sheet()
        assert wb1.get_sheet_extent(name) == (0, 0)

        wb1.set_cell_contents(name, 'A1', 'red')
        assert wb1.get_sheet_extent(name) == (1, 1)

        wb1.set_cell_contents(name, 'A1', '')
        assert wb1.get_sheet_extent(name) == (0, 0)

        wb1.set_cell_contents(name, 'A1', 'red')
        assert wb1.get_sheet_extent(name) == (1, 1)

        wb1.set_cell_contents(name, 'D14', 'green')
        assert wb1.get_sheet_extent(name) == (4, 14)

        with pytest.raises(KeyError):
            wb1.set_cell_contents('Sheet2', 'B2', 'blue')

    def test_get_contents_and_value(self) -> None:
        '''
        Test getting contents and value of a cell

        '''

        wb1 = Workbook()
        (_, name) = wb1.new_sheet()
        assert name == 'Sheet1'

        with pytest.raises(KeyError):
            wb1.get_cell_value('July Totals', 'A1')

        with pytest.raises(KeyError):
            wb1.get_cell_contents('July Totals', 'A1')

        # empty cells
        contents = wb1.get_cell_contents(name, 'ABC123')
        assert contents is None
        value = wb1.get_cell_value(name, 'ABC123')
        assert value is None

        # setting contents to None
        wb1.set_cell_contents(name, 'A1', None)
        contents = wb1.get_cell_contents(name, 'A1')
        assert contents is None
        value = wb1.get_cell_value(name, 'A1')
        assert value is None

        # number values
        wb1.set_cell_contents(name, 'A1', '8')
        contents = wb1.get_cell_contents(name, 'A1')
        assert contents == '8'
        value = wb1.get_cell_value(name, 'A1')
        assert value == Decimal(8)

        # string values
        wb1.set_cell_contents(name, 'A1', 'eight')
        contents = wb1.get_cell_contents(name, 'A1')
        assert contents == 'eight'
        value = wb1.get_cell_value(name, 'A1')
        assert value == 'eight'

        # string values with whitespace
        wb1.set_cell_contents(name, 'A1', '\'   eight')
        contents = wb1.get_cell_contents(name, 'A1')
        assert contents == '\'   eight'
        value = wb1.get_cell_value(name, 'A1')
        assert value == '   eight'

        # simple formulas (no references to other cells)
        # formulas tested more extensively elsewhere
        wb1.set_cell_contents(name, 'A1', '=1+1')
        contents = wb1.get_cell_contents(name, 'A1')
        assert contents == '=1+1'
        value = wb1.get_cell_value(name, 'A1')
        assert value == Decimal(2)

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
        wb1.set_cell_contents('Sheet1', 'C1', None)
        assert test_changed[-1] == [('Sheet1', 'C1')]
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
        assert contents == None
        assert value == None