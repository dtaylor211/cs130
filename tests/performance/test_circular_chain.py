'''
Performance - Test Circular Chain

Performance tests for when a large circular cell reference chain is created
in the workbook.

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_circular_chain() -> None

'''

import cProfile
from pstats import Stats

# pylint: disable=unused-import, import-error
from sheets import Workbook, CellError, CellErrorType
from tests import context


def test_circular_chain() -> None:
    '''
    Stress test for a long chain of references that produces a cycle

    '''

    wb1 = Workbook()
    _, name = wb1.new_sheet('Sheet1')

    wb1.set_cell_contents(name, 'A1', '=1')

    for i in range(2, 801):
        wb1.set_cell_contents(name, f'A{i}', f'=A{i - 1}')

    wb1.set_cell_contents(name, 'A1', '=A800')
    value = wb1.get_cell_value(name, 'A1')
    assert isinstance(value, CellError)
    assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE
    value = wb1.get_cell_value(name, 'A800')
    assert isinstance(value, CellError)
    assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()

    test_circular_chain()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)
