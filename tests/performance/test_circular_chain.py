import cProfile
from tests import context

from pstats import Stats
from decimal import Decimal
from sheets import *

def test_circular_chain():
    '''
    Stress test for a long chain of references that produces a cycle

    '''

    wb = Workbook()
    (index, name) = wb.new_sheet('Sheet1')

    wb.set_cell_contents(name, 'A1', '=1')

    for i in range(2, 801):
        wb.set_cell_contents(name, f'A{i}', f'=A{i - 1}')

    wb.set_cell_contents(name, 'A1', '=A800')
    value = wb.get_cell_value(name, 'A1')
    assert isinstance(value, CellError)
    assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE
    value = wb.get_cell_value(name, 'A800')
    assert isinstance(value, CellError)
    assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

if __name__ == '__main__':
    profiler = cProfile.Profile()
    profiler.enable()
    
    test_circular_chain()

    profiler.disable()
    stats = Stats(profiler).sort_stats('cumtime')
    stats.print_stats(10)