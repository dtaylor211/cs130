'''
Test Sheet

Tests the Sheet module found at ../sheets/sheet.py with
valid inputs.

Classes:
- TestSheet

    Methods:
    - test_extent_simple(object) -> None
    - test_extent_complex(object) -> None

'''

# pylint: disable=unused-import, import-error
import context
import pytest
from sheets.sheet import Sheet


class TestSheet:
    '''
    (Spread)Sheet tests (Project 1 & 2)

    '''

    def test_extent_simple(self):
        '''
        Test simple extents of sheet

        '''

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

    def test_extent_complex(self):
        '''
        Test complex extents of sheet

        '''

        sheet = Sheet("July Totals", None)
        assert sheet.get_extent() == (0, 0)

        sheet.set_cell_contents("A1", "1")
        assert sheet.get_extent() == (1, 1)

        sheet.set_cell_contents("AAc751", "50")
        assert sheet.get_extent() == (705, 751)

        sheet.set_cell_contents("D5", "3")
        assert sheet.get_extent() == (705, 751)

        sheet.set_cell_contents("AAc751", "  ")
        assert sheet.get_extent() == (4, 5)

        sheet.set_cell_contents("A1", "")
        assert sheet.get_extent() == (4, 5)

        sheet.set_cell_contents("E1", "7")
        assert sheet.get_extent() == (5, 5)

        sheet.set_cell_contents("C10", "7")
        assert sheet.get_extent() == (5, 10)

        sheet.set_cell_contents("E1", "")
        assert sheet.get_extent() == (4, 10)
        