'''
Test Sheet

Tests the Sheet module found at ../sheets/sheet.py.

Classes:
- TestSheet

    Methods:
    - test_extent_simple(object) -> None
    - test_extent_complex(object) -> None
    - test_get_source_cells(object) -> None
    - test_get_target_cells(object) -> None

'''

import pytest

# pylint: disable=unused-import, import-error
import context
from sheets.sheet import Sheet


class TestSheet:
    '''
    Sheet tests

    '''

    def test_extent_simple(self) -> None:
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

    def test_extent_complex(self) -> None:
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

    def test_get_source_cells(self) -> None:
        '''
        Test getting a group of source cells

        '''

        sheet = Sheet('Aloha', None)
        sheet.set_cell_contents('A1', '1')
        source_cells = sheet.get_source_cells('A1', 'A1')
        assert source_cells == ['A1']

        sheet.set_cell_contents('B3', '2')
        sheet.set_cell_contents('A2', '3')
        sheet.set_cell_contents('A4', '1')
        source_cells = sheet.get_source_cells('A1', 'B3')
        result_list = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
        assert source_cells == result_list
        
        source_cells = sheet.get_source_cells('B3', 'A1')
        assert source_cells == result_list

        source_cells = sheet.get_source_cells('A3', 'B1')
        assert source_cells == result_list

        source_cells = sheet.get_source_cells('B1', 'A3')
        assert source_cells == result_list

        with pytest.raises(ValueError):
            sheet.get_source_cells('AAAAA1', 'B2')

        with pytest.raises(ValueError):
            sheet.get_source_cells('A1', 'BB12345')

    def test_get_target_cells(self) -> None:
        '''
        Test getting a group of target cells

        '''

        sheet = Sheet('Source', None)
        sheet.set_cell_contents('A1', '1')
        sheet.set_cell_contents('A3', '2')
        sheet.set_cell_contents('B3', '3')
        source_cells = sheet.get_source_cells('A1', 'B3')
        target_cells = sheet.get_target_cells('A1', 'B3', 'B2', source_cells)
        result_dict = {
            'B2': '1', 'B3': None, 'B4': '2',
            'C2': None, 'C3': None, 'C4': '3'
        }
        assert target_cells == result_dict

        target_cells = sheet.get_target_cells('A3', 'B1', 'B2', source_cells)
        assert target_cells == result_dict

        target_cells = sheet.get_target_cells('B1', 'A3', 'B2', source_cells)
        assert target_cells == result_dict

        target_cells = sheet.get_target_cells('B3', 'A1', 'B2', source_cells)
        assert target_cells == result_dict

        with pytest.raises(ValueError):
            sheet.get_target_cells('AAAAA1', 'B2', 'A1', None)

        with pytest.raises(ValueError):
            sheet.get_target_cells('A1', 'BB12345', 'B2', None)
