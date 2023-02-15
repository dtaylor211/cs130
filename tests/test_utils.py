'''
Test Utils

Tests the Utils module found at ../sheets/utils.py with both
valid and invalid inputs.

Classes:
- TestUtils

    Methods:
    - test_get_coords_from_loc(object) -> None
    - test_get_loc_from_coords(object) -> None

'''

# pylint: disable=unused-import, import-error
import context
import pytest
from sheets.utils import get_loc_from_coords, get_coords_from_loc


class TestUtils:
    '''
    Utils tests

    '''

    def test_get_coords_from_loc(self) -> None:
        '''
        Test getting coordinates from location

        '''

        with pytest.raises(ValueError):
            get_coords_from_loc('A0')
        with pytest.raises(ValueError):
            get_coords_from_loc('A-1')
        with pytest.raises(ValueError):
            get_coords_from_loc('A 1')
        with pytest.raises(ValueError):
            get_coords_from_loc(' A1')
        with pytest.raises(ValueError):
            get_coords_from_loc('A1 ')
        with pytest.raises(ValueError):
            get_coords_from_loc('AAAAA1')
        with pytest.raises(ValueError):
            get_coords_from_loc('A11111')
        with pytest.raises(ValueError):
            get_coords_from_loc('A0001')

        col, row = get_coords_from_loc('a1')
        assert col, row == (1, 1)

        col, row = get_coords_from_loc('a5')
        assert col, row == (1, 5)

        col, row = get_coords_from_loc('AA15')
        assert col, row == (27, 15)

        col, row = get_coords_from_loc('Aa16')
        assert col, row == (27, 16)

        col, row = get_coords_from_loc('AAC750')
        assert col, row == (705, 750)

        col, row = get_coords_from_loc('AAc751')
        assert col, row == (705, 751)

    def test_get_loc_from_coords(self) -> None:
        '''
        Test getting location from coordinates

        '''

        loc = get_loc_from_coords((1, 1))
        assert loc == 'A1'
