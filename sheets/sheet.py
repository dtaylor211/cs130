import re
from typing import Dict, List, Tuple, Optional, Any

from .cell import Cell

class Sheet:
    '''
     A spreadsheet containing zero or more cells

    '''

    def __init__(self, sheet_name, evaluator):
        ''' Initialize a new spreadsheet '''
        
        self.name = sheet_name 

        # dictonary that maps (row, col) tuple -> Cell
        # need to make sure inputted location "D14" is converted to (4, 14)
        ## location = D14, coords = (4,14)
        self.cells: Dict[Tuple[int, int], Cell] = {}
        self.evaluator = evaluator

    def get_extent(self) -> Tuple[int, int]:
        '''
        Get the extent of spreadsheet (# rows, # cols).
        New empty sheet has extent of (0, 0)

        Returns:
        - tuple of ints with extent

        '''

        if len(self.cells.keys()) == 0:
            return (0, 0) # empty sheet

        coords = list(self.cells.keys())
        rows = [coord[0] for coord in coords]
        cols = [coord[1] for coord in coords]
        return max(rows), max(cols)

    def get_cell(self, location: str) -> Optional[Cell]:
        '''
        Get the cell object from a given location

        Arguments:
        - location: str - cell's location

        Returns:
        - either the Cell or None

        '''

        return self.cells[self.get_coords_from_loc(location)]

    def get_coords_from_loc(self, location: str) -> Tuple[int, int]:
        '''
        Get the coordinate tuple from a location

        raise ValueError is cell location isn't available
        need to check A-Z (max 4) then 1-9999 for valid lcoation

        Arguments:
        - location: str - location formatted as "B12"

        Returns:
        - tuple containing the coordinates (col, row) 

        '''
        
        if not re.match(r"^[A-Z]{1,4}[1-9][0-9]{0,3}$", location.upper()):
            raise ValueError("Cell location is invalid")

        # example: "D14" -> (4, 14)
        # splits into [characters, numbers, ""]
        split_loc = re.split(r'(\d+)', location.upper())
        (row, col) = split_loc[0], split_loc[1]
        col_num = int(col)
        row_num = 0
        for letter in row:
            row_num = row_num * 26 + ord(letter) - ord('A') + 1

        return (row_num, col_num)

    def get_cell_contents(self, location: str) -> Optional[str]:
        '''
        Get the contents of a cell

        Arguments:
        - location: str - cell's location

        Returns:
        - either the string contents or None

        '''

        coords = self.get_coords_from_loc(location)
        if coords not in self.cells.keys():
            return None
        
        return self.cells[coords].contents

    def set_cell_contents(self, location: str, contents: Optional[str]) -> None:
        '''
        Set the contents of a cell

        Arguments:
        - location: str - cell's location
        - contents: Optional[str] - the instructions on contents to set or None

        '''

        coords = self.get_coords_from_loc(location)
        if coords not in self.cells.keys():
            cell = Cell(location, self.evaluator)
            self.cells[coords] = cell

        if contents is None or contents.strip() == "":
            self.cells[coords].empty()
            del self.cells[coords]
            return

        self.cells[coords].set_contents(contents)

    def get_cell_value(self, location: str) -> Any:
        '''
        Get the value of a cell

        Arguments:
        - location: str - cell's location

        Returns:
        - The value of a cell

        '''

        coords = self.get_coords_from_loc(location)
        if coords not in self.cells.keys():
            return None

        return self.cells[coords].value # I think we need a get_value()

    def get_cell_adjacency_list(self) -> Dict[
        Tuple[str, str], List[Tuple[str, str]]]:
        '''
        Gets the adjacency list of cells in the sheet.

        Returns:
        - dictionary of coordinate tuples as keys and a list of 
            coordinate tuples as values
            
        '''

        adjacency_list = {}
        for cell in self.cells.values():
            adjacency_list[(self.name.lower(), cell.loc.lower())] = cell.get_children()
        return adjacency_list
