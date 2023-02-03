import re
from typing import Dict, List, Tuple, Optional, Any

from .cell import Cell
from .evaluator import Evaluator
from .utils import get_loc_from_coords

class Sheet:
    '''
     A spreadsheet containing zero or more cells

    '''

    def __init__(self, sheet_name, evaluator):
        ''' Initialize a new spreadsheet '''
        
        self._name = sheet_name 

        # dictonary that maps (row, col) tuple -> Cell
        # need to make sure inputted location "D14" is converted to (4, 14)
        ## location = D14, coords = (4,14)
        self._cells: Dict[Tuple[int, int], Cell] = {}
        self._evaluator = evaluator

    ########################################################################
    # Getters and Setters
    ########################################################################

    def get_name(self) -> str:
        '''
        Get the name of the sheet

        Returns:
        - string of sheet name

        '''

        return self._name
    
    def set_name(self, sheet_name: str):
        '''
        Set the name of the sheet

        Arguments:
        - sheet_name: str 

        '''

        self._name = sheet_name

    def get_all_cells(self) -> Dict[Tuple[int, int], Cell]:
        '''
        Get the name of the sheet

        Returns:
        - string of sheet name

        '''
        
        return self._cells
    
    def get_evaluator(self) -> Evaluator:
        '''
        Get the Evaluator for a sheet

        Returns:
        - the sheet's Evaluator

        '''

        return self._evaluator
    
    ########################################################################
    # Base Functionality
    ########################################################################

    def get_extent(self) -> Tuple[int, int]:
        '''
        Get the extent of spreadsheet (# rows, # cols).
        New empty sheet has extent of (0, 0)

        Returns:
        - tuple of ints with extent

        '''

        cells = self.get_all_cells()
        if len(cells.keys()) == 0:
            return (0, 0) # empty sheet

        coords = list(cells.keys())
        cols = [coord[0] for coord in coords]
        rows = [coord[1] for coord in coords]
        return max(cols), max(rows)

    def get_cell(self, location: str) -> Optional[Cell]:
        '''
        Get the cell object from a given location

        Arguments:
        - location: str - cell's location

        Returns:
        - either the Cell or None

        '''

        cells = self.get_all_cells()
        return cells[self.__get_coords_from_loc(location)]

    def __get_coords_from_loc(self, location: str) -> Tuple[int, int]:
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
        (col, row) = split_loc[0], split_loc[1]
        row_num = int(row)
        col_num = 0
        for letter in col:
            col_num = col_num * 26 + ord(letter) - ord('A') + 1

        return (col_num, row_num)

    def get_cell_contents(self, location: str) -> Optional[str]:
        '''
        Get the contents of a cell

        Arguments:
        - location: str - cell's location

        Returns:
        - either the string contents or None

        '''

        cells = self.get_all_cells()
        coords = self.__get_coords_from_loc(location)
        if coords not in cells.keys():
            return None
        
        return cells[coords].get_contents()

    def set_cell_contents(self, location: str, contents: Optional[str]) -> None:
        '''
        Set the contents of a cell

        Arguments:
        - location: str - cell's location
        - contents: Optional[str] - the instructions on contents to set or None

        '''

        cells = self.get_all_cells()
        coords = self.__get_coords_from_loc(location)
        if coords not in cells.keys():
            cell = Cell(location, self.get_evaluator())
            cells[coords] = cell

        if contents is None or contents.strip() == "":
            cells[coords].empty()
            del cells[coords]
            return

        cells[coords].set_contents(contents)

    def get_cell_value(self, location: str) -> Any:
        '''
        Get the value of a cell

        Arguments:
        - location: str - cell's location

        Returns:
        - The value of a cell

        '''

        cells = self.get_all_cells()
        coords = self.__get_coords_from_loc(location)
        if coords not in cells.keys():
            return None

        return cells[coords].get_value()

    def get_cell_adjacency_list(self) -> Dict[
        Tuple[str, str], List[Tuple[str, str]]]:
        '''
        Gets the adjacency list of cells in the sheet.

        Returns:
        - dictionary of coordinate tuples as keys and a list of 
            coordinate tuples as values

        '''

        adj_list = {}
        cells = self.get_all_cells()
        for cell in cells.values():
            name = self.get_name()
            adj_list[(name.lower(), cell.get_loc().lower())] = cell.get_children()
        return adj_list

    def save_sheet(self) -> Dict[str, str]:
        '''
        Saves Sheet in proper dictionary formatting for JSON export

        Returns:
        - dictionary with sheet name and nested dictionary for cell contents
        '''
        cell_contents = {}

        for coords, cell in self.get_all_cells().items():
            cell_loc = get_loc_from_coords(coords) # uppercase cell location
            cell_contents[cell_loc] = cell.get_contents()

        return {
            "name": self.get_name(),
            "cell-contents": cell_contents
        }
