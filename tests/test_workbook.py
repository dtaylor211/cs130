import context

import pytest
import io
import json
from decimal import Decimal

from sheets.workbook import Workbook
from sheets.cell_error import CellError, CellErrorType


class TestWorkbook:
    ''' Workbook tests (Project 1 & 2) '''

    def test_empty_workbook(self):
        wb = Workbook()
        assert wb.num_sheets() == 0
        assert wb.list_sheets() == []

    def test_new_sheet_generate_name(self):
        wb = Workbook()
        (index, name) = wb.new_sheet()
        assert index == 0
        assert name == "Sheet1"
        assert wb.num_sheets() == 1
        assert wb.list_sheets() == ["Sheet1"]

        (index, name) = wb.new_sheet()
        assert index == 1
        assert name == "Sheet2"
        assert wb.num_sheets() == 2
        assert wb.list_sheets() == ["Sheet1", "Sheet2"]

    def test_new_sheet_valid_name(self):
        wb = Workbook()
        (index, name) = wb.new_sheet("July Totals")
        assert index == 0
        assert name == "July Totals"
        assert wb.num_sheets() == 1
        assert wb.list_sheets() == ["July Totals"]

        (index, name) = wb.new_sheet("June Totals")
        assert index == 1
        assert name == "June Totals"
        assert wb.num_sheets() == 2
        assert wb.list_sheets() == ["July Totals", "June Totals"]

    def test_new_sheet_invalid_name(self):
        wb = Workbook()
        with pytest.raises(ValueError):
            (index, name) = wb.new_sheet("")
        with pytest.raises(ValueError):
            (index, name) = wb.new_sheet("   July Totals")
        with pytest.raises(ValueError):
            (index, name) = wb.new_sheet("July Totals  ")
        with pytest.raises(ValueError):
            (index, name) = wb.new_sheet("July Totals >")
        with pytest.raises(ValueError):
            (index, name) = wb.new_sheet("July\'s Totals")

        (index, name) = wb.new_sheet()
        assert name == "Sheet1"
        with pytest.raises(ValueError):
            (index, name) = wb.new_sheet("Sheet1")

    def test_new_sheet_complex(self):
        wb = Workbook()
        (index, name) = wb.new_sheet("Sheet2")
        assert index == 0
        assert name == "Sheet2"
        assert wb.num_sheets() == 1
        assert wb.list_sheets() == ["Sheet2"]

        (index, name) = wb.new_sheet("July Totals")
        assert index == 1
        assert name == "July Totals"
        assert wb.num_sheets() == 2
        assert wb.list_sheets() == ["Sheet2", "July Totals"]

        (index, name) = wb.new_sheet()
        assert index == 2
        assert name == "Sheet1"
        assert wb.num_sheets() == 3
        assert wb.list_sheets() == ["Sheet2", "July Totals", "Sheet1"]

        (index, name) = wb.new_sheet()
        assert index == 3
        assert name == "Sheet3"
        assert wb.num_sheets() == 4
        assert wb.list_sheets() == ["Sheet2", "July Totals", "Sheet1", "Sheet3"]

    def test_del_sheet(self):
        wb = Workbook()
        (index, name) = wb.new_sheet("Sheet1")
        (index, name) = wb.new_sheet("Sheet2")
        assert wb.num_sheets() == 2
        assert wb.list_sheets() == ["Sheet1", "Sheet2"]

        with pytest.raises(KeyError):
            wb.del_sheet("July Totals")

        wb.del_sheet("Sheet1")
        assert wb.num_sheets() == 1
        assert wb.list_sheets() == ["Sheet2"]

        wb.del_sheet("shEet2")
        assert wb.num_sheets() == 0
        assert wb.list_sheets() == []

    def test_set_and_extent(self):
        wb = Workbook()
        with pytest.raises(KeyError):
            wb.get_sheet_extent("Sheet1")

        (index, name) = wb.new_sheet()
        assert wb.get_sheet_extent(name) == (0, 0)

        wb.set_cell_contents(name, "A1", "red")
        assert wb.get_sheet_extent(name) == (1, 1)

        wb.set_cell_contents(name, "A1", "")
        assert wb.get_sheet_extent(name) == (0, 0)

        wb.set_cell_contents(name, "A1", "red")
        assert wb.get_sheet_extent(name) == (1, 1)

        wb.set_cell_contents(name, "D14", "green")
        assert wb.get_sheet_extent(name) == (4, 14)

        with pytest.raises(KeyError):
            wb.set_cell_contents("Sheet2", "B2", "blue")

    def test_get_contents_and_value(self):
        wb = Workbook()
        (index, name) = wb.new_sheet()
        assert name == "Sheet1"

        with pytest.raises(KeyError):
            wb.get_cell_value("July Totals", "A1")

        with pytest.raises(KeyError):
            wb.get_cell_contents("July Totals", "A1")

        # empty cells
        contents = wb.get_cell_contents(name, "ABC123")
        assert contents is None
        value = wb.get_cell_value(name, "ABC123")
        assert value is None

        # setting contents to None 
        wb.set_cell_contents(name, "A1", None)
        contents = wb.get_cell_contents(name, "A1")
        assert contents is None
        value = wb.get_cell_value(name, "A1")
        assert value is None

        # number values
        wb.set_cell_contents(name, "A1", "8")
        contents = wb.get_cell_contents(name, "A1")
        assert contents == "8"
        value = wb.get_cell_value(name, "A1")
        assert value == Decimal(8)

        # string values
        wb.set_cell_contents(name, "A1", "eight")
        contents = wb.get_cell_contents(name, "A1")
        assert contents == "eight"
        value = wb.get_cell_value(name, "A1")
        assert value == "eight"

        # string values with whitespace
        wb.set_cell_contents(name, "A1", "'   eight")
        contents = wb.get_cell_contents(name, "A1")
        assert contents == "'   eight"
        value = wb.get_cell_value(name, "A1")
        assert value == "   eight"

        # simple formulas (no references to other cells)
        # formulas tested more extensively elsewhere
        wb.set_cell_contents(name, "A1", "=1+1")
        contents = wb.get_cell_contents(name, "A1")
        assert contents == "=1+1"
        value = wb.get_cell_value(name, "A1")
        assert value == Decimal(2)

    ########################################################################
    # Project 2
    ########################################################################

    def test_load_workbook_valid(self):
        with open("tests/json_data/wb_data_valid.json") as fp:
            wb = Workbook.load_workbook(fp)
            assert wb.num_sheets() == 2
            assert wb.list_sheets() == ["Sheet1", "Sheet2"]
            assert wb.get_cell_contents("Sheet1", "A1") == "'123"
            assert wb.get_cell_contents("Sheet1", "B1") == "5.3"
            assert wb.get_cell_contents("Sheet1", "C1") == "=A1*B1"
            assert wb.get_cell_value("Sheet1", "A1") == "123"
            assert wb.get_cell_value("Sheet1", "B1") == Decimal("5.3")
            assert wb.get_cell_value("Sheet1", "C1") == Decimal("651.9")

    def test_load_workbook_dup_sheet(self):
        with open("tests/json_data/wb_data_invalid_dup.json") as fp:
            with pytest.raises(ValueError):
                wb = Workbook.load_workbook(fp)

    def test_load_workbook_missing_value(self):
        with open("tests/json_data/wb_data_missing_sheets.json") as fp:
            with pytest.raises(KeyError):
                wb = Workbook.load_workbook(fp)
        with open("tests/json_data/wb_data_missing_name.json") as fp:
            with pytest.raises(KeyError):
                wb = Workbook.load_workbook(fp)
        with open("tests/json_data/wb_data_missing_contents.json") as fp:
            with pytest.raises(KeyError):
                wb = Workbook.load_workbook(fp)

    def test_load_workbook_improper_type(self):
        with open("tests/json_data/wb_data_bad_type_sheets.json") as fp:
            with pytest.raises(TypeError):
                wb = Workbook.load_workbook(fp)
        with open("tests/json_data/wb_data_bad_type_sheet.json") as fp:
            with pytest.raises(TypeError):
                wb = Workbook.load_workbook(fp)
        with open("tests/json_data/wb_data_bad_type_name.json") as fp:
            with pytest.raises(TypeError):
                wb = Workbook.load_workbook(fp)
        with open("tests/json_data/wb_data_bad_type_contents.json") as fp:
            with pytest.raises(TypeError):
                wb = Workbook.load_workbook(fp)
        with open("tests/json_data/wb_data_bad_type_cell_contents.json") as fp:
            with pytest.raises(TypeError):
                wb = Workbook.load_workbook(fp)
        
    def test_load_workbook_prop_json_error(self):
        with open("tests/json_data/wb_data_bad_type_location.json") as fp:
            with pytest.raises(json.JSONDecodeError):
                wb = Workbook.load_workbook(fp)

    def test_save_workbook(self):
        with io.StringIO("") as fp:
            wb = Workbook()
            wb.new_sheet("Sheet1")
            wb.new_sheet("Sheet2")
            wb.set_cell_contents("Sheet1", "A1", "1")
            wb.set_cell_contents("Sheet2", "B2", "2")
            wb.save_workbook(fp)
            fp.seek(0)
            json_act = json.load(fp)
            json_exp = {
                "sheets":[
                    {
                        "name":"Sheet1",
                        "cell-contents":{
                            "A1":"1"
                        }
                    },
                    {
                        "name":"Sheet2",
                        "cell-contents":{
                            "B2":"2"
                        }
                    }
                ]
            }
            assert json_act == json_exp

        with io.StringIO("") as fp:
            wb = Workbook()
            wb.new_sheet("Sheet1")
            wb.set_cell_contents("Sheet1", "A1", "1")
            wb.new_sheet("Sheet2")
            wb.set_cell_contents("Sheet2", "B2", "=Sheet1!A1")
            wb.save_workbook(fp)
            fp.seek(0)
            json_act = json.load(fp)
            json_exp = {
                "sheets":[
                    {
                        "name":"Sheet1",
                        "cell-contents":{
                            "A1":"1"
                        }
                    },
                    {
                        "name":"Sheet2",
                        "cell-contents":{
                            "B2":"=Sheet1!A1"
                        }
                    }
                ]
            }
            assert json_act == json_exp

    def test_notify_cell(self):
        pass # TODO @Kyle?

    def test_rename_sheet(self):
        # test if Sheet1Sheet1 and Sheet1, and we rename Sheet1
        wb = Workbook()
        wb.new_sheet('Sheet1')
        wb.set_cell_contents('Sheet1', 'A1', '2')
        wb.set_cell_contents('Sheet1', 'A2', '=A1')
        wb.rename_sheet('Sheet1', 'SheEt2')
        sheet_names = wb.list_sheets()
        sheet_objects = wb.sheet_objects
        assert sheet_names == ['SheEt2']
        assert sheet_objects['sheet2'] is not None
        with pytest.raises(KeyError):
            sheet_objects['sheet1']
        value = wb.get_cell_value('Sheet2', 'A1')
        assert value == Decimal(2)
        value = wb.get_cell_value('Sheet2','A2')
        assert value == Decimal(2)

        wb.new_sheet('Sheet3')
        wb.rename_sheet('Sheet2', 'Sheet4')
        new_sheet_names = wb.list_sheets()
        new_sheet_objects = wb.sheet_objects
        assert new_sheet_names == ['Sheet4', 'Sheet3']
        assert new_sheet_objects['sheet4'] is not None
        assert new_sheet_objects['sheet3'] is not None
        with pytest.raises(KeyError):
            sheet_objects['sheet2']
        value = wb.get_cell_value('Sheet4', 'A1')
        assert value == Decimal(2)

    def test_rename_sheet_update_refs(self):
        wb = Workbook()
        wb.new_sheet('Sheet1')
        wb.new_sheet('Sheet2')
        wb.set_cell_contents('Sheet1', 'A1', '=2')
        wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1')
        wb.rename_sheet('Sheet1', 'Sheet3')
        contents = wb.get_cell_contents('Sheet2', 'A1')
        value = wb.get_cell_value('Sheet2', 'A1')
        assert contents == '=Sheet3!A1'
        assert value == Decimal(2)

        wb.set_cell_contents('Sheet3', 'A2', '=3')
        wb.set_cell_contents('Sheet3', 'A3', '=4')
        wb.set_cell_contents('Sheet2', 'A2', '=Sheet3!A2')
        wb.set_cell_contents('Sheet2', 'A3', '=Sheet3!A3')
        wb.rename_sheet('Sheet3', 'Sheet4')
        contents = wb.get_cell_contents('Sheet2', 'A1')
        value = wb.get_cell_value('Sheet2', 'A1')
        assert contents == '=Sheet4!A1'
        assert value == Decimal(2)
        contents = wb.get_cell_contents('Sheet2', 'A2')
        value = wb.get_cell_value('Sheet2', 'A2')
        assert contents == '=Sheet4!A2'
        assert value == Decimal(3)
        contents = wb.get_cell_contents('Sheet2', 'A3')
        value = wb.get_cell_value('Sheet2', 'A3')
        assert contents == '=Sheet4!A3'
        assert value == Decimal(4)

        wb.new_sheet('Sheet5')
        wb.set_cell_contents('Sheet5', 'A1', '=Sheet2!A2 + Sheet2!A3')
        wb.set_cell_contents('Sheet5', 'A2', '=Sheet2!A2 + Sheet4!A3')
        wb.rename_sheet('Sheet2', 'Sheet6')
        wb.rename_sheet('Sheet4', 'Sheet7')
        contents = wb.get_cell_contents('Sheet5', 'A1')
        value = wb.get_cell_value('Sheet5', 'A1')
        assert contents == '=Sheet6!A2 + Sheet6!A3'
        assert value == Decimal(7)
        contents = wb.get_cell_contents('Sheet5', 'A2')
        value = wb.get_cell_value('Sheet5', 'A2')
        assert contents == '=Sheet6!A2 + Sheet7!A3'
        assert value == Decimal(7)

        wb.new_sheet('A Sheet')
        wb.set_cell_contents('A Sheet', 'A1', '=0.1')
        wb.set_cell_contents('Sheet5', 'A1', "='A Sheet'!A1")
        wb.rename_sheet('A Sheet', 'Darth Jar Jar')
        contents = wb.get_cell_contents('Sheet5', 'A1')
        value = wb.get_cell_value('Sheet5', 'A1')
        assert contents == "='Darth Jar Jar'!A1"
        assert value == Decimal('0.1')

        wb.new_sheet('Sheet1')
        wb.new_sheet('Sheet1Sheet1')
        wb.set_cell_contents('Sheet1', 'A1', '2')
        wb.set_cell_contents('Sheet1Sheet1', 'A1', '1.5')
        wb.set_cell_contents('Sheet5', 'A1', '=Sheet1Sheet1!A1 * Sheet1!A1')
        wb.rename_sheet('Sheet1', 'Sheet99')
        contents = wb.get_cell_contents('Sheet5', 'A1')
        value = wb.get_cell_value('Sheet5', 'A1')
        assert contents == '=Sheet1Sheet1!A1 * Sheet99!A1'
        assert value == Decimal('3')

    def test_rename_sheet_apply_quotes(self):
        wb = Workbook()
        wb.new_sheet('Sheet1')
        wb.new_sheet('Sheet2')
        wb.set_cell_contents('Sheet1', 'A1', 'do or do not, there is no try')
        wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1')
        wb.rename_sheet('Sheet1', 'Sheet 1')
        contents = wb.get_cell_contents('Sheet2', 'A1')
        value = wb.get_cell_value('Sheet2', 'A1')
        assert contents == "='Sheet 1'!A1"
        assert value == 'do or do not, there is no try'

        wb.new_sheet('Sheet3')
        wb.rename_sheet('Sheet 1', 'Sheet1')
        wb.set_cell_contents('Sheet1', 'A2', '\' roger')
        wb.set_cell_contents('Sheet3', 'A1', '- Yoda (master shiesty)')
        wb.set_cell_contents('Sheet2', 'A2', '=Sheet1!A1 & Sheet3!A1')
        wb.set_cell_contents('Sheet2', 'A3', '=Sheet1!A2 & Sheet1!A2')
        wb.rename_sheet('Sheet1', 'Sheet 1')
        wb.rename_sheet('Sheet3', 'Sheet 3')
        contents = wb.get_cell_contents('Sheet2', 'A2')
        value = wb.get_cell_value('Sheet2', 'A2')
        assert contents == "='Sheet 1'!A1 & 'Sheet 3'!A1"
        assert value == 'do or do not, there is no try- Yoda (master shiesty)'
        contents = wb.get_cell_contents('Sheet2', 'A3')
        value = wb.get_cell_value('Sheet2', 'A3')
        assert contents == "='Sheet 1'!A2 & 'Sheet 1'!A2"
        assert value == ' roger roger'

        wb.new_sheet('Sheet4')
        wb.new_sheet('ShEet5')
        wb.set_cell_contents('Sheet4', 'A1', 'good relations')
        wb.set_cell_contents('Sheet4', 'A2', '\' with the wookies,')
        wb.set_cell_contents('Sheet5', 'A1', '\' I have')
        wb.set_cell_contents('Sheet2', 'A1', '=Sheet4!A1 & Sheet4!A2 & Sheet5!A1')
        wb.rename_sheet('Sheet4', 'Sheet4?')
        contents = wb.get_cell_contents('Sheet2', 'A1')
        value = wb.get_cell_value('Sheet2', 'A1')
        assert contents == "='Sheet4?'!A1 & 'Sheet4?'!A2 & Sheet5!A1"
        assert value == 'good relations with the wookies, I have'

    def test_rename_sheet_remove_quotes(self):
        wb = Workbook()
        wb.new_sheet('Sheet1')
        wb.new_sheet('Benjamin Juarez')
        wb.set_cell_contents('Benjamin Juarez', 'A1', 'i heart jar jar binks')
        wb.set_cell_contents('Sheet1', 'A1', '=\'Benjamin Juarez\'!A1')
        wb.rename_sheet('Benjamin Juarez', 'BJ')
        contents = wb.get_cell_contents('Sheet1', 'A1')
        value = wb.get_cell_value('Sheet1', 'A1')
        assert contents == '=BJ!A1'
        assert value == 'i heart jar jar binks'

        wb.new_sheet('Kyle McGraw')
        wb.set_cell_contents('Kyle McGraw', 'A1', 'anakin skywalker')
        wb.set_cell_contents('Sheet1', 
                             'A2', '=BJ!A1&" and "&\'Kyle McGraw\'!A1')
        wb.rename_sheet('Kyle McGraw', 'KM')
        contents = wb.get_cell_contents('Sheet1', 'A2')
        value = wb.get_cell_value('Sheet1', 'A2')
        assert contents == '=BJ!A1&" and "&KM!A1'
        assert value == 'i heart jar jar binks and anakin skywalker'

        wb.new_sheet('Sheet2')
        wb.set_cell_contents('Sheet2', 'A1', '=\'KM\'!A1 & KM!A1')
        wb.rename_sheet('KM', 'DT')
        contents = wb.get_cell_contents('Sheet2', 'A1')
        value = wb.get_cell_value('Sheet2', 'A1')
        assert contents == '=DT!A1 & DT!A1'
        assert value == 'anakin skywalkeranakin skywalker'

    def test_rename_sheet_parse_error(self):
        wb = Workbook()
        wb.new_sheet('Sheet1')
        wb.new_sheet('Sheet2')
        wb.set_cell_contents('Sheet1', 'A1', 'Dallas Taylor')
        wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1 & Sheet1!!A1')
        wb.rename_sheet('Sheet1', 'Sheet11')
        contents = wb.get_cell_contents('Sheet2', 'A1')
        value = wb.get_cell_value('Sheet2', 'A1')
        assert contents == '=Sheet1!A1 & Sheet1!!A1'
        assert isinstance(value, CellError)
        assert(value.get_type() == CellErrorType.PARSE_ERROR)

        wb.new_sheet('Sheet1')
        wb.set_cell_contents('Sheet1', 'A1', 'Dallas Taylor')
        wb.set_cell_contents('Sheet2', 'A1', '=Sheet1!A1 &&')
        wb.rename_sheet('Sheet1', 'Sheet12')
        contents = wb.get_cell_contents('Sheet2', 'A1')
        value = wb.get_cell_value('Sheet2', 'A1')
        assert contents == '=Sheet1!A1 &&'
        assert isinstance(value, CellError)
        assert(value.get_type() == CellErrorType.PARSE_ERROR)

    def test_move_sheet(self):
        wb = Workbook()
        wb.new_sheet("Sheet1")
        wb.new_sheet("Sheet2")
        wb.new_sheet("Sheet3")
        with pytest.raises(KeyError):
            wb.move_sheet("Sheet4", 0)
        with pytest.raises(IndexError):
            wb.move_sheet("Sheet3", -1)
        with pytest.raises(IndexError):
            wb.move_sheet("Sheet3", 4)
        wb.move_sheet("Sheet3", 0)
        assert wb.list_sheets() == ["Sheet3", "Sheet1", "Sheet2"]
        wb.move_sheet("Sheet3", 2)
        assert wb.list_sheets() == ["Sheet1", "Sheet2", "Sheet3"]

    def test_copy_sheet(self):
        wb = Workbook()
        wb.new_sheet("Sheet1")
        wb.set_cell_contents("Sheet1", "A1", "=1")
        wb.copy_sheet("Sheet1")
        wb.copy_sheet("Sheet1")
        assert wb.num_sheets() == 3
        assert wb.list_sheets() == ["Sheet1", "Sheet1_1", "Sheet1_2"]

        wb.copy_sheet("Sheet1_2")
        assert wb.list_sheets() == ["Sheet1", "Sheet1_1", "Sheet1_2", 
            "Sheet1_2_1"]
        
        assert wb.get_cell_value("Sheet1", "A1") == Decimal("1")
        assert wb.get_cell_value("Sheet1_1", "A1") == Decimal("1")
        assert wb.get_cell_value("Sheet1_2", "A1") == Decimal("1")
        assert wb.get_cell_value("Sheet1_2_1", "A1") == Decimal("1")

        wb.set_cell_contents("Sheet1", "A1", "=2")
        assert wb.get_cell_value("Sheet1", "A1") == Decimal("2")
        assert wb.get_cell_value("Sheet1_1", "A1") == Decimal("1")
        assert wb.get_cell_value("Sheet1_2", "A1") == Decimal("1")
        assert wb.get_cell_value("Sheet1_2_1", "A1") == Decimal("1")

        wb.set_cell_contents("Sheet1_2", "A1", "=3")
        assert wb.get_cell_value("Sheet1", "A1") == Decimal("2")
        assert wb.get_cell_value("Sheet1_1", "A1") == Decimal("1")
        assert wb.get_cell_value("Sheet1_2", "A1") == Decimal("3")
        assert wb.get_cell_value("Sheet1_2_1", "A1") == Decimal("1")

        wb.new_sheet("Sheet2")
        wb.new_sheet("Sheet3")
        wb.set_cell_contents("Sheet2", "B2", "=2")
        wb.set_cell_contents("Sheet3", "B2", "=Sheet2!B2+Sheet2!B2")
        wb.copy_sheet("Sheet2")
        wb.copy_sheet("Sheet3")
        wb.copy_sheet("Sheet2")

        assert wb.list_sheets() == ["Sheet1", "Sheet1_1", "Sheet1_2", 
            "Sheet1_2_1", "Sheet2", "Sheet3", "Sheet2_1", "Sheet3_1", 
            "Sheet2_2"]
        assert wb.get_cell_value("Sheet2_1", "B2") == Decimal("2")
        assert wb.get_cell_value("Sheet3_1", "B2") == Decimal("4")

        wb.new_sheet("Sheet4")
        wb.set_cell_contents("Sheet4", "D4", "=#CIRCREF!")
        wb.copy_sheet("Sheet4")
        contents = wb.get_cell_contents("Sheet4_1", "D4")
        value = wb.get_cell_value("Sheet4_1", "D4")
        assert contents == "=#CIRCREF!"
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

        (idx, name) = wb.copy_sheet("Sheet4")
        assert name == "Sheet4_2"
        (idx, name) = wb.copy_sheet("Sheet4")
        assert name == "Sheet4_3"