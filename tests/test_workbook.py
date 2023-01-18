import pytest
from sheets.workbook import Workbook
from decimal import Decimal

class TestWorkbook:
    '''
    Workbook tests (project 1)
    '''

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
        # with pytest.raises(KeyError) -> test deleting non-existent sheet
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

        # finish?

    def test_get_contents(self):
        pass

    def test_get_value(self):
        wb = Workbook()
        (index, name) = wb.new_sheet()
        with pytest.raises(KeyError):
            wb.get_cell_value("July Totals", "A1")

        # TODO: test all types of cell values

    def test_reference_sheet(self):
        wb = Workbook()
        (index, name) = wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "1")
        (index, name) = wb.new_sheet()
        wb.set_cell_contents("Sheet2", "B2", "=Sheet1!A1")
        cell_value = wb.get_cell_value("Sheet2", "B2")
        assert cell_value == decimal.Decimal(1)

        # finish/more to do

    def test_dependency(self):
        pass

    # def test_set_contents(self):
    #     pass
