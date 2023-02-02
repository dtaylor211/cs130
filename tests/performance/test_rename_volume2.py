import cProfile
from tests import context

from pstats import Stats
from decimal import Decimal
from sheets import *

def test_rename_volume2():
    wb = Workbook()
    _, name = wb.new_sheet("Sheet1")

    for i in range(1, 801):
        wb.set_cell_contents(name, f"A{i}", f'{i}')

    wb.set_cell_contents(name, 'A801', f'={name}!A1')
    wb.rename_sheet(name, name+'1')
    value = wb.get_cell_value(name+'1', 'A1')
    assert value == Decimal('1')
    
    contents = wb.get_cell_contents(name+'1', 'A801')
    value = wb.get_cell_value(name+'1', 'A801')
    assert contents == f'={name}1!A1'
    assert value == Decimal('1')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()
    
    test_rename_volume2()

    profiler.disable()
    stats = Stats(profiler).sort_stats("cumtime")
    stats.print_stats(10)