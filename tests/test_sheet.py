import pytest
from sheets.sheet import Sheet
from decimal import Decimal

class TestSheet:
    '''
    (Spread)Sheet tests (project 1)
    '''

    def test_extent(self):
        sheet = Sheet("July Totals")
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
        sheet = Sheet("Sheet1")
        (col, row) = sheet.get_coords_from_loc("A1")
        assert (col, row) == (1, 1)

        (col, row) = sheet.get_coords_from_loc("A5")
        assert (col, row) == (1, 5)

        (col, row) = sheet.get_coords_from_loc("AA15")
        assert (col, row) == (27, 15)

        (col, row) = sheet.get_coords_from_loc("AAC750")
        assert (col, row) == (705, 750)

        # check more

    # not applicable?
    # def test_get_cell(self):
    #     pass

    def test_get_contents(self):
        # TODO
        pass

    def test_set_contents(self):
        # TODO
        pass

    def test_get_value(self):
        # TODO
        pass
