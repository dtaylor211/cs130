import context

import pytest
from decimal import Decimal

from sheets.sheet import Sheet


class TestSheet:
    '''
    (Spread)Sheet tests (Project 1 & 2)

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
            sheet.set_cell_contents("A0", None)
        with pytest.raises(ValueError):
            sheet.set_cell_contents("A-1", None)
        with pytest.raises(ValueError):
            sheet.set_cell_contents("A 1", None)
        with pytest.raises(ValueError):
            sheet.set_cell_contents(" A1", None)
        with pytest.raises(ValueError):
            sheet.set_cell_contents("A1 ", None)
        with pytest.raises(ValueError):
            sheet.set_cell_contents("AAAAA1", None)
        with pytest.raises(ValueError):
            sheet.set_cell_contents("A11111", None)
        with pytest.raises(ValueError):
            sheet.set_cell_contents("A0001", None)

        # (col, row) = sheet.__get_coords_from_loc("a1")
        # assert (col, row) == (1, 1)
        sheet.set_cell_contents("a1", '1')
        assert sheet.get_extent() == (1, 1)

        # (col, row) = sheet.__get_coords_from_loc("a5")
        # assert (col, row) == (1, 5)
        sheet.set_cell_contents("a5", '1')
        assert sheet.get_extent() == (1, 5)

        # (col, row) = sheet.__get_coords_from_loc("AA15")
        # assert (col, row) == (27, 15)
        sheet.set_cell_contents("AA15", '1')
        assert sheet.get_extent() == (27, 15)

        # (col, row) = sheet.__get_coords_from_loc("Aa16")
        # assert (col, row) == (27, 16)
        sheet.set_cell_contents("Aa16", '1')
        assert sheet.get_extent() == (27, 16)

        # (col, row) = sheet.__get_coords_from_loc("AAC750")
        # assert (col, row) == (705, 750)
        sheet.set_cell_contents("AAC750", '1')
        assert sheet.get_extent() == (705, 750)

        # (col, row) = sheet.__get_coords_from_loc("AAc751")
        # assert (col, row) == (705, 751)
        sheet.set_cell_contents("AAc751", '1')
        assert sheet.get_extent() == (705, 751)
