import context

import pytest
from decimal import Decimal

from sheets.sheet import Sheet


class TestSheet:
    '''
    (Spread)Sheet tests (Project 1)

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

        sheet.set_cell_contents("A1", "")
        assert sheet.get_extent() == (0, 0)

    def test_get_coords_from_loc(self):
        sheet = Sheet("Sheet1", None)
        with pytest.raises(ValueError):
            sheet.__get_coords_from_loc("A0")
        with pytest.raises(ValueError):
            sheet.__get_coords_from_loc("A-1")
        with pytest.raises(ValueError):
            sheet.__get_coords_from_loc("A 1")
        with pytest.raises(ValueError):
            sheet.__get_coords_from_loc(" A1")
        with pytest.raises(ValueError):
            sheet.__get_coords_from_loc("A1 ")
        with pytest.raises(ValueError):
            sheet.__get_coords_from_loc("AAAAA1")
        with pytest.raises(ValueError):
            sheet.__get_coords_from_loc("A11111")
        with pytest.raises(ValueError):
            sheet.__get_coords_from_loc("A0001")

        (col, row) = sheet.__get_coords_from_loc("A1")
        assert (col, row) == (1, 1)

        (col, row) = sheet.__get_coords_from_loc("a1")
        assert (col, row) == (1, 1)

        (col, row) = sheet.__get_coords_from_loc("A5")
        assert (col, row) == (1, 5)

        (col, row) = sheet.__get_coords_from_loc("a5")
        assert (col, row) == (1, 5)

        (col, row) = sheet.__get_coords_from_loc("AA15")
        assert (col, row) == (27, 15)

        (col, row) = sheet.__get_coords_from_loc("Aa15")
        assert (col, row) == (27, 15)

        (col, row) = sheet.__get_coords_from_loc("AAC750")
        assert (col, row) == (705, 750)

        (col, row) = sheet.__get_coords_from_loc("AAc750")
        assert (col, row) == (705, 750)
