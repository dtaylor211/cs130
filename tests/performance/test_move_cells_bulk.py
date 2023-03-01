'''
Performance - Test Move Cells Bulk

Performance tests for computing a large volume of move cell operations.

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_move_cells_bulk() -> None

'''

import cProfile
from pstats import Stats

# pylint: disable=unused-import, import-error
from sheets import Workbook, CellError, CellErrorType
from tests import context 


def test_move_cells_bulk() -> None:
    '''
    Stress test for moving regions of cells

    '''

    wb1 = Workbook()
    _, name = wb1.new_sheet('Sheet1')
    for i in range(2, 51):
        wb1.set_cell_contents(name, f'A{i}', f'=A{i - 1} + 1')
    wb1.set_cell_contents(name, 'A1', "=1")

    for i in range(10):
        wb1.move_cells("Sheet1", f'A{i + 1}', f'A{i + 50}', f'A{i + 2}')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()

    test_move_cells_bulk()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)
