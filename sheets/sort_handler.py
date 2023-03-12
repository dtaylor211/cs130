'''
Sort Handler

This module holds the extended functionality for sorting rows of cells
in a given range

See the Workbook module for implementation.

Classes:
- Row

    Methods:
    - 

'''

from typing import Dict, Any

class Row:
    '''
    
    '''

    def __init__(self, row_order: int, cells: Dict[str, Any]):
        self.row_order = row_order
        self.cells = cells

    def get_column_value(self, column: str):
        # print(self.cells)
        # print(11, column, self.cells[column])
        print(1, self.cells[column])
        return self.cells[column]
    
    def __repr__(self): 
        # this is only to help w debugging - formats in print statements
        return str(self.row_order) + str(self.cells)
