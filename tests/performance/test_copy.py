'''
Performance - Test Copy

Performance tests for when a large group of source cells is copied to another
sheet in the workbook.

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_copy() -> None

'''

import cProfile
from pstats import Stats
from decimal import Decimal

# pylint: disable=unused-import, import-error
from tests import context
from sheets import Workbook


def test_copy() -> None:
    '''
    Stress test for copying a large group of source cells

    '''

    wb1 = Workbook()
    _, s_sheet = wb1.new_sheet('Source')
    _, t_sheet = wb1.new_sheet('Target')

    for i in range(1, 501):
        wb1.set_cell_contents(s_sheet, f'A{i}', f'={i}')
        wb1.set_cell_contents(s_sheet, f'B{i}', f'={s_sheet}!A{i}')

    wb1.copy_cells(s_sheet, 'A1', 'B500', 'B1', t_sheet)
    value = wb1.get_cell_value(t_sheet, 'B1')
    assert value == Decimal('1')

    contents = wb1.get_cell_contents(t_sheet, 'C500')
    value = wb1.get_cell_value(t_sheet, 'C500')
    assert contents == f'={s_sheet}!B500'
    assert value == Decimal('500')

    contents = wb1.get_cell_contents(s_sheet, 'A1')
    assert contents == '=1'

    contents = wb1.get_cell_contents(s_sheet, 'B250')
    value = wb1.get_cell_value(s_sheet, 'B250')
    assert contents == f'={s_sheet}!A250'
    assert value == Decimal('250')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()

    test_copy()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)
