'''
Performance - Test Fibonacci

Performance tests for computing the Fibonacci sequence of cell references.

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_fib() -> None

'''

import cProfile
from pstats import Stats

# pylint: disable=unused-import, import-error
from sheets import Workbook, CellError, CellErrorType
from tests import context


def test_fib() -> None:
    '''
    Stress test for a Fibonacci sequence of cell references

    '''

    wb1 = Workbook()
    _, name = wb1.new_sheet('Sheet1')

    wb1.set_cell_contents(name, 'A2', '=1')

    for i in range(3, 1001):
        wb1.set_cell_contents(name, f'A{i}', f'=A{i - 2} + A{i - 1}')

    wb1.set_cell_contents(name, 'A1', '=1')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()

    test_fib()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)
