'''
Performance - Test (M x N) Mesh Cycle

Performance tests for when M large circular cell reference chains (of N cells) 
are created in the workbook (meant to simulate benchmark test 
MxNMeshCycleBenchmark with M-row mesh where each row has an N-cell-long 
circular chain of cells).

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_mesh_cycle() -> None
- get_coords_from_loc(str) -> Tuple[int, int]

'''

import cProfile
from pstats import Stats
from typing import Tuple

# pylint: disable=unused-import, import-error
from sheets import Workbook, CellError, CellErrorType
from tests import context

def get_loc_from_coords(coords: Tuple[int, int]) -> str:
    '''
    Get a cell location from its coordinates (FROM utils.py)

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

def test_mesh_cycle() -> None:
    '''
    Stress tests for a long chain of cell references

    '''
    wb1 = Workbook()
    _, name = wb1.new_sheet('Sheet1')

    wb1.set_cell_contents(name, 'A1', '=1')

    for row in range(1, 21): # M = 20
        for col in range(2, 51): # N = 50
            loc = get_loc_from_coords((row, col))
            next_loc = str(get_loc_from_coords((row, col - 1)))
            wb1.set_cell_contents(name, loc, '=' + next_loc + ' + 1')
    for row in range(1, 21):
        loc = get_loc_from_coords((row, 1))
        last_loc = get_loc_from_coords((col, 50))
        wb1.set_cell_contents(name, loc, '=' + last_loc)

    value = wb1.get_cell_value(name, 'A1')
    assert isinstance(value, CellError)
    assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE
    value = wb1.get_cell_value(name, 'A50')
    assert isinstance(value, CellError)
    assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()

    test_mesh_cycle()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(20)
