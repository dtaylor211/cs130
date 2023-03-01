'''
Performance - Test Load Workbook

Performance tests for loading a workbook.

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_load_wb() -> None

'''

import cProfile
import io
from pstats import Stats

# pylint: disable=unused-import, import-error
from sheets import Workbook, CellError, CellErrorType
from tests import context

if __name__ == '__main__':

    wb1 = Workbook()
    _, name = wb1.new_sheet('Sheet1')
    for i in range(2, 501):
        wb1.set_cell_contents(name, f'A{i}', f'=A{i - 1} + 1')
    wb1.set_cell_contents(name, 'A1', "=1")
    for i in range(2, 501):
        wb1.set_cell_contents(name, f'B{i}', f'=B{i - 1} + 2')
    wb1.set_cell_contents(name, 'B1', "=1")
    for i in range(2, 501):
        wb1.set_cell_contents(name, f'C{i}', f'=C{i - 1} + 3')
    wb1.set_cell_contents(name, 'C1', "=1")

    with io.StringIO() as fp:
        wb1.save_workbook(fp)
        fp.seek(0)

        profiler = cProfile.Profile()
        profiler.enable()

        wb2 = Workbook.load_workbook(fp)

        profiler.disable()

    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)
