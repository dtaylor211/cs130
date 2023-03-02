'''
Performance - Test Rename Sheet Bulk

Performance tests for computing a large volume of rename sheet operations.

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_rename_bulk() -> None

'''

import cProfile
from pstats import Stats

# pylint: disable=unused-import, import-error
from sheets import Workbook, CellError, CellErrorType
from tests import context


def test_rename_bulk() -> None:
    '''
    Stress test for renaming sheets

    '''

    wb1 = Workbook()
    _, name = wb1.new_sheet('Sheet1')
    for i in range(2, 501):
        wb1.set_cell_contents(name, f'A{i}', f'=A{i - 1} + 1')
    wb1.set_cell_contents(name, 'A1', "=1")
    # for i in range(2, 501):
    #     wb1.set_cell_contents(name, f'B{i}', f'=B{i - 1} + 2')
    # wb1.set_cell_contents(name, 'B1', "=1")
    # for i in range(2, 501):
    #     wb1.set_cell_contents(name, f'C{i}', f'=C{i - 1} + 3')
    # wb1.set_cell_contents(name, 'C1', "=1")

    for i in range(10):
        if i % 2 == 0:
            wb1.rename_sheet("Sheet1", "Sheet2")
        else:
            wb1.rename_sheet("Sheet2", "Sheet1")

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()

    test_rename_bulk()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)
