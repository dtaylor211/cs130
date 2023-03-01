'''
Performance - Test Long Chain Update

Performance tests for when a large cell reference chain is created
in the workbook (meant to simulate benchmark test LongChainUpdateBenchmark)

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_long_chain_update() -> None

'''

import cProfile
from pstats import Stats
from decimal import Decimal

# pylint: disable=unused-import, import-error
from sheets import Workbook
from tests import context


def test_long_chain_update() -> None:
    '''
    Stress tests for a long chain of cell references

    '''

    wb1 = Workbook()
    _, name = wb1.new_sheet('Sheet1')

    wb1.set_cell_contents(name, 'A1', '=1')

    for i in range(2, 1001):
        wb1.set_cell_contents(name, f'A{i}', f'=A{i - 1} + 1')

    wb1.set_cell_contents(name, 'A1', '=1')
    assert wb1.get_cell_value(name, 'A1000') == Decimal('1000')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()

    test_long_chain_update()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)
