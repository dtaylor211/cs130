'''
Test Workbook

Tests the Workbook module found at ../sheets/workbook.py with more complex
operations.  Tests for basic operations can be found at test_workbook.py

Classes:
- TestWorkbook

    Methods:
    - test_empty_workbook(object) -> None
    - test_new_sheet_name(object) -> None
    - test_new_sheet_complex(object) -> None
    - test_del_sheet(object) -> None
    - test_set_and_extent(object) -> None
    - test_get_contents_and_value(object) -> None

'''

from decimal import Decimal

# pylint: disable=unused-import, import-error
import context
import pytest
from sheets.workbook import Workbook
from sheets.cell_error import CellError, CellErrorType

class TestWorkbook:
    ''' 
    Workbook tests
    
    '''

    def test_empty_workbook(self) -> None:
        '''
        Test empty workbook

        '''

        wb1 = Workbook()
        assert wb1.num_sheets() == 0
        assert wb1.list_sheets() == []

    def test_new_sheet_name(self) -> None:
        '''
        Test new sheet names

        '''

        wb1 = Workbook()
        index, name = wb1.new_sheet()
        assert index == 0
        assert name == 'Sheet1'
        assert wb1.num_sheets() == 1
        assert wb1.list_sheets() == ['Sheet1']

        index, name = wb1.new_sheet()
        assert index == 1
        assert name == 'Sheet2'
        assert wb1.num_sheets() == 2
        assert wb1.list_sheets() == ['Sheet1', 'Sheet2']

        wb1 = Workbook()
        index, name = wb1.new_sheet('July Totals')
        assert index == 0
        assert name == 'July Totals'
        assert wb1.num_sheets() == 1
        assert wb1.list_sheets() == ['July Totals']

        index, name = wb1.new_sheet('June Totals')
        assert index == 1
        assert name == 'June Totals'
        assert wb1.num_sheets() == 2
        assert wb1.list_sheets() == ['July Totals', 'June Totals']

        wb1 = Workbook()
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('')
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('   July Totals')
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('July Totals  ')
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('July Totals >')
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('July\'s Totals')

        index, name = wb1.new_sheet()
        assert name == 'Sheet1'
        with pytest.raises(ValueError):
            index, name = wb1.new_sheet('Sheet1')

    def test_new_sheet_complex(self) -> None:
        '''
        Test new sheet with more complex checks to adding new fields

        '''

        wb1 = Workbook()
        index, name = wb1.new_sheet('Sheet2')
        assert index == 0
        assert name == 'Sheet2'
        assert wb1.num_sheets() == 1
        assert wb1.list_sheets() == ['Sheet2']

        index, name = wb1.new_sheet('July Totals')
        assert index == 1
        assert name == 'July Totals'
        assert wb1.num_sheets() == 2
        assert wb1.list_sheets() == ['Sheet2', 'July Totals']

        index, name = wb1.new_sheet()
        assert index == 2
        assert name == 'Sheet1'
        assert wb1.num_sheets() == 3
        assert wb1.list_sheets() == ['Sheet2', 'July Totals', 'Sheet1']

        index, name = wb1.new_sheet()
        assert index == 3
        assert name == 'Sheet3'
        assert wb1.num_sheets() == 4
        assert wb1.list_sheets() == ['Sheet2', 'July Totals', 'Sheet1', 'Sheet3']

    def test_del_sheet(self) -> None:
        '''
        Test deleting a sheet

        '''

        wb1 = Workbook()
        wb1.new_sheet('Sheet1')
        wb1.new_sheet('Sheet2')
        assert wb1.num_sheets() == 2
        assert wb1.list_sheets() == ['Sheet1', 'Sheet2']

        with pytest.raises(KeyError):
            wb1.del_sheet('July Totals')

        wb1.del_sheet('Sheet1')
        assert wb1.num_sheets() == 1
        assert wb1.list_sheets() == ['Sheet2']

        wb1.del_sheet('shEet2')
        assert wb1.num_sheets() == 0
        assert wb1.list_sheets() == []

    def test_set_and_extent(self) -> None:
        '''
        Test setting sheet and getting extent

        '''

        wb1 = Workbook()
        with pytest.raises(KeyError):
            wb1.get_sheet_extent('Sheet1')

        (_, name) = wb1.new_sheet()
        assert wb1.get_sheet_extent(name) == (0, 0)

        wb1.set_cell_contents(name, 'A1', 'red')
        assert wb1.get_sheet_extent(name) == (1, 1)

        wb1.set_cell_contents(name, 'A1', '')
        assert wb1.get_sheet_extent(name) == (0, 0)

        wb1.set_cell_contents(name, 'A1', 'red')
        assert wb1.get_sheet_extent(name) == (1, 1)

        wb1.set_cell_contents(name, 'D14', 'green')
        assert wb1.get_sheet_extent(name) == (4, 14)

        with pytest.raises(KeyError):
            wb1.set_cell_contents('Sheet2', 'B2', 'blue')

    def test_get_contents_and_value(self) -> None:
        '''
        Test getting contents and value of a cell

        '''

        wb1 = Workbook()
        (_, name) = wb1.new_sheet()
        assert name == 'Sheet1'

        with pytest.raises(KeyError):
            wb1.get_cell_value('July Totals', 'A1')

        with pytest.raises(KeyError):
            wb1.get_cell_contents('July Totals', 'A1')

        # empty cells
        contents = wb1.get_cell_contents(name, 'ABC123')
        assert contents is None
        value = wb1.get_cell_value(name, 'ABC123')
        assert value is None

        # setting contents to None
        wb1.set_cell_contents(name, 'A1', None)
        contents = wb1.get_cell_contents(name, 'A1')
        assert contents is None
        value = wb1.get_cell_value(name, 'A1')
        assert value is None

        # number values
        wb1.set_cell_contents(name, 'A1', '8')
        contents = wb1.get_cell_contents(name, 'A1')
        assert contents == '8'
        value = wb1.get_cell_value(name, 'A1')
        assert value == Decimal(8)

        # string values
        wb1.set_cell_contents(name, 'A1', 'eight')
        contents = wb1.get_cell_contents(name, 'A1')
        assert contents == 'eight'
        value = wb1.get_cell_value(name, 'A1')
        assert value == 'eight'

        # string values with whitespace
        wb1.set_cell_contents(name, 'A1', '\'   eight')
        contents = wb1.get_cell_contents(name, 'A1')
        assert contents == '\'   eight'
        value = wb1.get_cell_value(name, 'A1')
        assert value == '   eight'

        # simple formulas (no references to other cells)
        # formulas tested more extensively elsewhere
        wb1.set_cell_contents(name, 'A1', '=1+1')
        contents = wb1.get_cell_contents(name, 'A1')
        assert contents == '=1+1'
        value = wb1.get_cell_value(name, 'A1')
        assert value == Decimal(2)
