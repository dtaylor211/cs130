import cProfile
from tests import context

from pstats import Stats
from decimal import Decimal
from sheets import *

def test_rename_chain():
    '''
    Stress tests for renaming sheet when it will cause a long chain of updates 
    in another sheet

    '''
    
    wb = Workbook()
    _, name = wb.new_sheet('Sheet1')
    _, name2 = wb.new_sheet('Sheet2')

    wb.set_cell_contents(name, 'A1', '=1')
    wb.set_cell_contents(name2, 'A1', f'={name}!A1')

    for i in range(2, 201):
        wb.set_cell_contents(name2, f'A{i}', f'=A{i-1}')

    wb.rename_sheet(name, name+'1')
    contents = wb.get_cell_contents(name2, 'A1')
    value = wb.get_cell_value(name2, 'A1')
    assert contents == f'={name}1!A1'
    assert value == Decimal('1')
    
    contents = wb.get_cell_contents(name2, 'A200')
    value = wb.get_cell_value(name2, 'A200')
    assert contents == '=A199'
    assert value == Decimal('1')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()
    
    test_rename_chain()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)