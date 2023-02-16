'''
Performance - Test Rename Volume 1

Performance tests for when a large volume of cells in another sheet
references a sheet that is renamed.  

For a similar test that involves a reference chain in another sheet,
see test_rename_chain.py.
For a similar test that involves a reference chain in its own sheet,
see test_rename_volume2.py

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_rename_volume1() -> None

'''

import cProfile
from pstats import Stats
from decimal import Decimal

# pylint: disable=unused-import, import-error
from sheets import Workbook
from tests import context


def test_rename_volume1() -> None:
    '''
    Stress tests for renaming a sheet when a large number of cells in another 
    sheet reference a cell in the sheet to be renamed

    '''

    wb1 = Workbook()
    _, name = wb1.new_sheet('Sheet1')
    _, name2 = wb1.new_sheet('Sheet2')

    wb1.set_cell_contents(name, 'A1', '=1')

    for i in range(1, 201):
        wb1.set_cell_contents(name2, f'A{i}', f'={name}!A1')

    wb1.rename_sheet(name, name+'1')
    contents = wb1.get_cell_contents(name2, 'A1')
    value = wb1.get_cell_value(name2, 'A1')
    assert contents == f'={name}1!A1'
    assert value == Decimal('1')

    contents = wb1.get_cell_contents(name2, 'A200')
    value = wb1.get_cell_value(name2, 'A200')
    assert contents == f'={name}1!A1'
    assert value == Decimal('1')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()

    test_rename_volume1()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)
