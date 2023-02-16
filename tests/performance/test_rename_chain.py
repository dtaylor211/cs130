'''
Performance - Test Rename Chain

Performance tests for when a large reference chain is created in the
workbook and then the sheet referenced in that chain is renamed.

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_rename_chain() -> None

'''


import cProfile
from pstats import Stats
from decimal import Decimal

# pylint: disable=unused-import, import-error
from sheets import Workbook
from tests import context


def test_rename_chain() -> None:
    '''
    Stress tests for renaming sheet when it will cause a long chain of updates 
    in another sheet

    '''

    wb1 = Workbook()
    _, name = wb1.new_sheet('Sheet1')
    _, name2 = wb1.new_sheet('Sheet2')

    wb1.set_cell_contents(name, 'A1', '=1')
    wb1.set_cell_contents(name2, 'A1', f'={name}!A1')

    for i in range(2, 201):
        wb1.set_cell_contents(name2, f'A{i}', f'=A{i-1}')

    wb1.rename_sheet(name, name+'1')
    contents = wb1.get_cell_contents(name2, 'A1')
    value = wb1.get_cell_value(name2, 'A1')
    assert contents == f'={name}1!A1'
    assert value == Decimal('1')

    contents = wb1.get_cell_contents(name2, 'A200')
    value = wb1.get_cell_value(name2, 'A200')
    assert contents == '=A199'
    assert value == Decimal('1')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()

    test_rename_chain()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)
