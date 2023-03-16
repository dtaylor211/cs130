'''
Test Utils

Tests the Utils module found at ../sheets/utils.py with both
valid and invalid inputs.

Classes:
- TestUtils

    Methods:
    - test_get_coords_from_loc(object) -> None
    - test_get_loc_from_coords(object) -> None
    - test_convert_to_bool(object) -> None
    - test_compare_values(object) -> None

'''


# pylint: disable=unused-import, import-error
from decimal import Decimal

import pytest

import context
from sheets import CellError, CellErrorType
from sheets.utils import get_loc_from_coords, get_coords_from_loc,\
    convert_to_bool, compare_values


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

        with pytest.raises(ValueError):
            get_loc_from_coords((0, 0))
        with pytest.raises(ValueError):
            get_loc_from_coords((0, 1))
        with pytest.raises(ValueError):
            get_loc_from_coords((1, 0))
        with pytest.raises(ValueError):
            get_loc_from_coords((10000, 0))
        with pytest.raises(ValueError):
            get_loc_from_coords((0, 10000))
        with pytest.raises(ValueError):
            get_loc_from_coords((10000, 10000))

        loc = get_loc_from_coords((1, 1))
        assert loc == 'A1'

        loc = get_loc_from_coords((1, 5))
        assert loc == 'A5'

        loc = get_loc_from_coords((27, 15))
        assert loc == 'AA15'

        loc = get_loc_from_coords((27, 16))
        assert loc == 'AA16'

        loc = get_loc_from_coords((705, 750))
        assert loc == 'AAC750'

        loc = get_loc_from_coords((705, 751))
        assert loc == 'AAC751'

    def test_covert_to_bool(self) -> None:
        '''
        Test converting strings and Decimals to bools

        '''

        with pytest.raises(TypeError):
            convert_to_bool('anystr')
        with pytest.raises(TypeError):
            convert_to_bool('1')
        with pytest.raises(TypeError):
            convert_to_bool(int(1))
        with pytest.raises(TypeError):
            convert_to_bool(int(1))
        with pytest.raises(TypeError):
            convert_to_bool(['True'])

        booly = convert_to_bool('true', str)
        assert booly

        booly = convert_to_bool('trUe', str)
        assert booly

        booly = convert_to_bool('FALSE', str)
        assert not booly

        booly = convert_to_bool(Decimal(0), Decimal)
        assert not booly

        booly = convert_to_bool(Decimal(1), Decimal)
        assert booly

        booly = convert_to_bool(Decimal('-0.231'), Decimal)
        assert booly

    def test_compare_values(self) -> None:
        '''
        Test comparing two values

        '''

        booly = compare_values(None, None, (type(None), type(None)), '=')
        assert booly

        booly = compare_values(Decimal('1'), True, (Decimal, bool), '<')
        assert booly

        booly = compare_values(Decimal('1'), "true", (Decimal, str), '<')
        assert booly

        booly = compare_values(None, "true", (type(None), str), '<')
        assert booly

        booly = compare_values(CellError(CellErrorType.BAD_REFERENCE, ""),
                               "true", (CellError, str), '<')
        assert booly

        booly = compare_values(CellError(CellErrorType.BAD_REFERENCE, ""),
                               None, (CellError, type(None)), '>')
        assert booly

