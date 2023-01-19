import pytest
import context
from sheets.sheet import Sheet
from decimal import Decimal
from sheets.formula_evaluator import Evaluator
from sheets.workbook import Workbook # marked for removal

workbook = Workbook()
evaluator = Evaluator(workbook,'')


class TestSheet:
    '''
    (Spread)Sheet tests (project 1)
    '''

    def test_extent(self):
        sheet = Sheet("July Totals", None)
        assert sheet.get_extent() == (0, 0)

        sheet.set_cell_contents("A1", "1")
        assert sheet.get_extent() == (1, 1)

        sheet.set_cell_contents("D5", "50")
        assert sheet.get_extent() == (4, 5)

        sheet.set_cell_contents("D5", "")
        assert sheet.get_extent() == (1, 1)

        sheet.set_cell_contents("D5", "   ")
        assert sheet.get_extent() == (1, 1)

    def test_get_coords_from_loc(self):
        sheet = Sheet("Sheet1", None)
        with pytest.raises(ValueError):
            sheet.get_coords_from_loc("A0") # check more cases?

        (col, row) = sheet.get_coords_from_loc("A1")
        assert (col, row) == (1, 1)

        (col, row) = sheet.get_coords_from_loc("A5")
        assert (col, row) == (1, 5)

        (col, row) = sheet.get_coords_from_loc("AA15")
        assert (col, row) == (27, 15)

        (col, row) = sheet.get_coords_from_loc("AAC750")
        assert (col, row) == (705, 750)

    # REMOVE THIS FROM THIS TESTING FILE, TRANSFER RELEVANT STUFF TO TEST_WORKBOOK

    # not applicable?
    # def test_get_cell(self):
    #     pass

    # def test_get_contents(self):
    #     sheet = Sheet("Sheet1")
    #     assert sheet.set_cell_contents("A1") is None

    #     sheet.set_cell_contents("A1", "1")
    #     contents = sheet.get_cell_contents("A1")
    #     assert contents == "1"

    # def test_get_value(self):
    #     sheet = Sheet("Sheet1")
    #     assert sheet.set_cell_contents("A1") is None

    #     sheet.set_cell_contents("A1", "1")
    #     value = sheet.get_cell_value("A1")
    #     assert value == Decimal(1)
