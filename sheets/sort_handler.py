'''
Sort Handler

This module holds the extended functionality for sorting rows of cells
in a given range

See the Workbook module for implementation.

Classes:
- Row

    Attributes:
    - row_order (int) - original order of the row in sorting area
    - row_length (int) - length of the row (including None values)

    Methods:
    - get_column_value(object) -> Any
    - set_current_column(object, int) -> None

'''

from typing import Dict, Any

from .utils import compare_values


class Row:
    '''
    A single instance of a row to be sorted.

    '''

    def __init__(self, row_order: int, cells: Dict[int, Any], row_length: int):
        '''
        Initialize a new row

        Arguments:
        - row_order: int - original order of the row in all rows to be sorted
        - cells: Dict[int, Any] - dictionary mapping the column index (1-based)
            to the cell's value
        - row_length: int - number of cells in the row

        '''

        self.row_order = row_order
        self.row_length = row_length
        self._cells = cells
        self._current_column = None

    def get_column_value(self) -> Any:
        '''
        Get the value of a cell at the current column's index

        Returns:
        - value of the specified cell

        '''
        print(self._current_column, self._cells[self._current_column])
        return self._cells[self._current_column]

    def set_current_column(self, column: int) -> None:
        '''
        Set the current column index to compare across

        Arguments:
        - column: int - index of current column to set

        '''

        if column not in range(1, self.row_length + 1):
            raise ValueError(f'incorrect column entry: {column}')

        self._current_column = column

    def __repr__(self) -> str:
        '''
        Provide the string representation of a row, to aid with debugging

        Returns:
        - string of the row order followed by the cell dictionary

        '''

        return str(self.row_order) + ': ' + str(self._cells)

    def __lt__(self, obj: 'Row') -> bool:
        '''
        Compare if self is less than another instance of Row

        Arguments:
        - obj: Row - instance to compare against

        Returns
        - bool result

        '''

        curr_value = self.get_column_value()
        obj_value = obj.get_column_value()
        types = (type(curr_value), type(obj_value))
        return compare_values(curr_value, obj_value, types, '<')

    def __gt__(self, obj: 'Row') -> bool:
        '''
        Compare if self is greater than another instance of Row

        Arguments:
        - obj: Row - instance to compare against

        Returns
        - bool result

        '''

        curr_value = self.get_column_value()
        obj_value = obj.get_column_value()
        types = (type(curr_value), type(obj_value))
        return compare_values(curr_value, obj_value, types, '>')

    def __le__(self, obj: 'Row') -> bool:
        '''
        Compare if self is less than or equal toanother instance of Row

        Arguments:
        - obj: Row - instance to compare against

        Returns
        - bool result

        '''

        curr_value = self.get_column_value()
        obj_value = obj.get_column_value()
        types = (type(curr_value), type(obj_value))
        return compare_values(curr_value, obj_value, types, '<=')

    def __ge__(self, obj: 'Row') -> bool:
        '''
        Compare if self is greater than or equal to another instance of Row

        Arguments:
        - obj: Row - instance to compare against

        Returns
        - bool result

        '''

        curr_value = self.get_column_value()
        obj_value = obj.get_column_value()
        types = (type(curr_value), type(obj_value))
        return compare_values(curr_value, obj_value, types, '>=')

    def __eq__(self, obj: 'Row') -> bool:
        '''
        Compare if self is equal to another instance of Row

        Arguments:
        - obj: Row - instance to compare against

        Returns
        - bool result

        '''

        curr_value = self.get_column_value()
        obj_value = obj.get_column_value()
        types = (type(curr_value), type(obj_value))
        return compare_values(curr_value, obj_value, types, '=')
