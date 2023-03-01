'''
Utils

This module holds utility functions for the Sheets package

See the Workbook and Sheet modules for implementation.

Methods:
- get_loc_from_coords(Tuple[int, int]) -> str
- get_coords_from_loc(str) -> Tuple[int, int]

'''


import re
from typing import Tuple
from decimal import Decimal


def get_loc_from_coords(coords: Tuple[int, int]) -> str:
    '''
    Get a cell location from its coordinates

    Throw a ValueError if the coordinates are invalid or out of bounds

    Arguments:
    - coords: Tuple[int, int] - tuple of col, row coordinates

    Returns:
    - str of cell location

    '''

    col, row = coords
    if col < 1 or row < 1 or col > 9999 or row > 9999:
        raise ValueError("Invalid coordinates")

    col_name = ""
    while col > 0:
        col_name = chr((col - 1) % 26 + ord('A')) + col_name
        col = (col - 1) // 26

    return col_name.upper() + str(row)

def get_coords_from_loc(location: str) -> Tuple[int, int]:
    '''
    Get the coordinate tuple from a location

    raise ValueError is cell location isn't available
    need to check A-Z (max 4) then 1-9999 for valid lcoation

    Arguments:
    - location: str - location formatted as "B12"

    Returns:
    - tuple containing the coordinates (col, row)

    '''
    if not re.match(r"^[A-Z]{1,4}[1-9][0-9]{0,3}$", location.upper()):
        raise ValueError("Cell location is invalid")

    # example: "D14" -> (4, 14)
    # splits into [characters, numbers, ""]
    split_loc = re.split(r'(\d+)', location.upper())
    (col, row) = split_loc[0], split_loc[1]
    row_num = int(row)
    col_num = 0
    for letter in col:
        col_num = col_num * 26 + ord(letter) - ord('A') + 1

    return (col_num, row_num)

def convert_to_bool(input, input_type: type) -> bool:
        '''
        '''
        
        result = None
        if input_type == bool:
            return input
        elif input_type == str:
            if input.lower() == 'true':
                result = True
            elif input.lower() == 'false':
                result = False
            else:
                raise TypeError('Cannot convert given string to boolean')
        else:
            if input != Decimal(0):
                result = True
            else:
                result = False
        return result
