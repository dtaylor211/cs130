import cProfile
from tests import context

from pstats import Stats
from decimal import Decimal
from sheets import *

def test_rename_volume1():
    wb = Workbook()
    _, name = wb.new_sheet("Sheet1")
    _, name2 = wb.new_sheet("Sheet2")

    wb.set_cell_contents(name, "A1", "=1")

    for i in range(1, 201):
        wb.set_cell_contents(name2, f"A{i}", f"={name}!A1")

    wb.rename_sheet(name, name+'1')
    contents = wb.get_cell_contents(name2, 'A1')
    value = wb.get_cell_value(name2, 'A1')
    assert contents == f'={name}1!A1'
    assert value == Decimal('1')
    
    contents = wb.get_cell_contents(name2, 'A200')
    value = wb.get_cell_value(name2, 'A200')
    assert contents == f'={name}1!A1'
    assert value == Decimal('1')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()
    
    test_rename_volume1()

    profiler.disable()
    stats = Stats(profiler).sort_stats("cumtime")
    stats.print_stats(10)