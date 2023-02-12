import cProfile
from tests import context

from pstats import Stats
from decimal import Decimal
from sheets import *

def test_reference_chain():
    '''
    Stress tests for a long chain of cell references

    '''
    
    wb = Workbook()
    (index, name) = wb.new_sheet('Sheet1')

    wb.set_cell_contents(name, 'A1', '=1')

    for i in range(2, 101):
        wb.set_cell_contents(name, f'A{i}', f'=A{i - 1}')

    wb.set_cell_contents(name, 'A1', '=2')
    assert wb.get_cell_value(name, 'A100') == Decimal('2')

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()
    
    test_reference_chain()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)