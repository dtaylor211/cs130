'''
Performance - Test Rename Volume 2

Performance tests for when a large volume of cells in a sheet that 
references itself and that sheet is renamed.  

For a similar test that involves a reference chain in another sheet, 
see test_rename_chain.py.
For a similar test that involves a large volume of references in another
sheet, see test_rename_volume1.py

When running this file, a profile is output to the terminal with the 10
most time consuming parts of the computation listed.

Methods:
- test_rename_volume2() -> None

'''


import cProfile
from pstats import Stats
from decimal import Decimal

# pylint: disable=unused-import, import-error
from sheets import Workbook
from tests import context


def test_rename_volume2():
    '''
    Stress tests for renaming a sheet when it causes a reference chain
    of cells in its own sheet to be updated

    '''

    wb1 = Workbook()
    _, name = wb1.new_sheet('Sheet1')

    wb1.set_cell_contents(name, 'A1', '1')
    value = wb1.get_cell_value(name, 'A1')

    for i in range(2, 801):
        wb1.set_cell_contents(name, f'A{i}', f'={name}!A{i-1}')

    wb1.rename_sheet(name, name+'1')
    value = wb1.get_cell_value(name+'1', 'A1')
    assert value == Decimal('1')

    contents = wb1.get_cell_contents(name+'1', 'A800')
    value = wb1.get_cell_value(name+'1', 'A800')
    assert contents == f'={name}1!A799'
    assert value == Decimal('1')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()

    test_rename_volume2()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)
