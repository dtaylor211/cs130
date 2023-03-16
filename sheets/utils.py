'''
Utils

This module holds utility functions for the Sheets package

See the Workbook and Sheet modules for implementation.

Global Variables:
- COMP_OPERATORS (Dict[str, built_in_function]) - converts string of operator
    to the operator function
- EMPTY_SUBS (Dict[type, Any]) - converts type of not None expression to the
    correct empty value

Methods:
- get_loc_from_coords(Tuple[int, int]) -> str
- get_coords_from_loc(str) -> Tuple[int, int]
- convert_to_bool(Any, type) -> bool
- get_tl_br_corners()
- compare_values(Any, Any, Tuple[type, type], str) -> bool

'''


import re
import operator
from typing import Tuple, Any, List
from decimal import Decimal

from .cell_error import CellError


COMP_OPERATORS = {
    ">": operator.gt,
    "<": operator.lt,
    "<=": operator.le,
    ">=": operator.ge,
    "=": operator.eq,
    "==": operator.eq,
    "!=": operator.ne,
    "<>": operator.ne
}

EMPTY_SUBS = {
    str: '',
    Decimal: Decimal(0),
    bool: False
}

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
    if col < 1 or row < 1 or col > 475254 or row > 9999:
        raise ValueError("Invalid coordinates")

    col_name = ""
    while col > 0:
        col_name = chr((col - 1) % 26 + ord('A')) + col_name
        col = (col - 1) // 26

    return col_name.upper() + str(row)

def get_coords_from_loc(location: str) -> Tuple[int, int]:
    '''
    Get the coordinate tuple from a location

    Throw a ValueError is cell location isn't available
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

def convert_to_bool(inp: Any, inp_type: type) -> bool:
    '''
    Convert given input of type str, Decimal, or boolean to boolean

    Throw a TypeError if given input cannot be converted

    Arguments:
    - inp: Any - input to convert
    - inp_type: type - type of input

    Returns:
    - boolean representation of input

    '''

    result = None
    if inp_type == bool:
        result = inp
    elif inp_type == str:
        if inp.lower() == 'true':
            result = True
        elif inp.lower() == 'false':
            result = False
        else:
            raise TypeError('Cannot convert given string to boolean')
    elif inp_type == Decimal:
        result = bool(inp)
    else: raise TypeError('Cannot convert given type to boolean')
    return result

def get_tl_br_corners(start_location: str, end_location: str
                          ) -> List[Tuple[int, int]]:
    '''
    Get the top left and bottom right corner coordinates from given
    corners that define the desired cell range

    Arguments:
    - start_location: str - corner cell location of source area
    - end_location: str - corner cell location of source area

    Returns:
    - List of top left, bottom right coordinates where each tuple 
        contains the col, row integer

    '''

    # get_coords_from_loc raises ValueError for invalid location
    start_col, start_row = get_coords_from_loc(start_location)
    end_col, end_row = get_coords_from_loc(end_location)

    top_left_col = min(start_col, end_col)
    top_left_row = min(start_row, end_row)
    bottom_right_col = max(start_col, end_col)
    bottom_right_row = max(start_row, end_row)

    return [
        (top_left_col, top_left_row),
        (bottom_right_col, bottom_right_row)
    ]

def get_source_cells(start_location: str,
        end_location: str) -> List[str]:
    '''
    Gets the list of source cell locations using start/end locations.

    Arguments:
    - start_location: str - corner cell location of source area
    - end_location: str - corner cell location of source area

    Returns:
    - List of string locations in the source area

    '''

    corners = get_tl_br_corners(start_location, end_location)
    top_left_col, top_left_row = corners[0]
    bottom_right_col, bottom_right_row = corners[-1]

    # List[str] = List[cell location]
    # get_loc_from_coords raises ValueError for invalid coords
    source_cells: List[str] = []
    for col in range(top_left_col, bottom_right_col + 1):
        for row in range(top_left_row, bottom_right_row + 1):
            loc = get_loc_from_coords((col, row))
            source_cells.append(loc)

    return source_cells

def compare_values(left: Any, right: Any, types: Tuple[type, type],
                        oper: str) -> bool:
    '''
    Get the boolean value for a comparison between types of bool, str,
    and/or Decimal

    Arguments:
    - left: Any - left side of comparison
    - right: Any - right side of comparison
    - types: Tuple[type, type] - types of the left and right sides of the
        comparison operator
    - oper: str - comparison operator

    Returns:
    - boolean result of comparison

    '''

    result = False
    if (left, right) == (None, None):
        result = COMP_OPERATORS[oper]('', '')

    elif types[0] == types[-1]:
        if types == (str, str):
            left = left.lower()
            right = right.lower()
        elif types == (CellError, CellError):
            left = left.get_type().value
            right = right.get_type().value
        result = COMP_OPERATORS[oper](left, right)

    elif (types in [(bool, str), (str, Decimal), (bool, Decimal)]) or \
        (types[0] == CellError and right is None) or \
            (types[-1] == CellError and left is not None):
        if oper in ['>', '>=', '!=', '<>']:
            result = True

    elif (types in [(str, bool), (Decimal, str), (Decimal, bool)]) or \
        (types[0] == CellError and right is not None) or \
            (types[-1] == CellError and left is None):
        if oper in ['<', '<=', '!=', '<>']:
            result = True

    else:
        if left is not None:
            result = COMP_OPERATORS[oper](left, EMPTY_SUBS[types[0]])
        else:
            result = COMP_OPERATORS[oper](EMPTY_SUBS[types[-1]], right)

    return result
