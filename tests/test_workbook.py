import context

import pytest
from decimal import Decimal

from sheets.workbook import Workbook


class TestWorkbook:
    ''' Workbook tests (Project 1) '''

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

        # simple formulas (no references to other cells)
        # formulas tested more extensively elsewhere
        wb.set_cell_contents(name, "A1", "=1+1")
        contents = wb.get_cell_contents(name, "A1")
        assert contents == "=1+1"
        value = wb.get_cell_value(name, "A1")
        assert value == Decimal(2)