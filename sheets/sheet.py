'''
Sheet

This module is where all operations on an individual sheet takes place.

See the Cell module for more detailed commentary.
See the Workbook module for implementation.

Classes:
- Sheet

    Methods:
    - get_name(object) -> str
    - set_name(object, str) -> None
    - get_all_cells(object) -> Dict[Tuple[int, int], Cell]
    - get_evaluator(object) -> Evaluator
    - get_extent(object) -> Tuple[int, int]
    - get_cell(object, str) -> Optional[Cell]
    - get_cell_contents(object, str) -> Optional[str]
    - set_cell_contents(object, str, Optional[str]) -> None
    - get_cell_value(object, str) -> Any
    - get_cell_adjacency_list(object) -> Dict[Tuple[str, str],
        List[Tuple[str, str]]]
    - save_sheet(object) -> Dict[str, str]

'''


from typing import Dict, List, Tuple, Optional, Any

from .cell import Cell
from .evaluator import Evaluator
from .utils import get_loc_from_coords, get_coords_from_loc


class Sheet:
    '''
     A spreadsheet containing zero or more cells

    '''

    def __init__(self, sheet_name, evaluator):
        '''
        Initialize a new spreadsheet

        '''

        self._name = sheet_name

        # dictonary that maps (row, col) tuple -> Cell
        # need to make sure inputted location "D14" is converted to (4, 14)
        # location = D14, coords = (4,14)
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

    def set_name(self, sheet_name: str) -> None:
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
        return cells[get_coords_from_loc(location)]

    def get_cell_contents(self, location: str) -> Optional[str]:
        '''
        Get the contents of a cell

        Arguments:
        - location: str - cell's location

        Returns:
        - either the string contents or None

        '''

        cells = self.get_all_cells()
        coords = get_coords_from_loc(location)
        if coords not in cells:
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
        coords = get_coords_from_loc(location)
        if coords not in cells:
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
        coords = get_coords_from_loc(location)
        if coords not in cells:
            return None

        return cells[coords].get_value()

    def get_cell_adjacency_list(self) -> Dict[Tuple[str, str],
                                              List[Tuple[str, str]]]:
        '''
        Gets the adjacency list of cells in the sheet.

        Returns:
        - Dict of coordinate tuples as keys and a list of coordinate
            tuples as values

        '''

        adj_list = {}
        cells = self.get_all_cells()
        for cell in cells.values():
            name = self.get_name()
            adj_list[(name, cell.get_loc())] = cell.get_children()
        return adj_list

    def save_sheet(self) -> Dict[str, str]:
        '''
        Saves Sheet in proper dictionary formatting for JSON export

        Returns:
        - Dict with sheet name and nested dictionary for cell contents

        '''

        cell_contents = {}

        for coords, cell in self.get_all_cells().items():
            cell_loc = get_loc_from_coords(coords) # uppercase cell location
            cell_contents[cell_loc] = cell.get_contents()

        return {
            "name": self.get_name(),
            "cell-contents": cell_contents
        }

    def get_source_cells(self, start_location: str, 
        end_location: str) -> List[str]:
        '''
        Gets the list of source cell locations using start/end locations.

        Arguments:
        - start_location: str - corner cell location of source area
        - end_location: str - corner cell location of source area

        Returns:
        - Dict mapping str cell locations to str source contents

        '''

        # get_coords_from_loc raises ValueError for invalid location
        start_col, start_row = get_coords_from_loc(start_location)
        end_col, end_row = get_coords_from_loc(end_location)

        top_left_col = min(start_col, end_col)
        top_left_row = min(start_row, end_row)
        bottom_right_col = max(start_col, end_col)
        bottom_right_row = max(start_row, end_row)

        # List[str] = List[cell location]
        # get_loc_from_coords raises ValueError for invalid coords
        source_cells: Dict[str, str] = {}
        for col in range(top_left_col, bottom_right_col + 1):
            for row in range(top_left_row, bottom_right_row + 1):
                coords = (col, row)
                loc = get_loc_from_coords(coords)
                source_cells[loc] = self.get_cell_contents(loc)

        return source_cells

    def get_target_cells(self, start_location: str, end_location: str, 
            to_location: str, source_cells: List[str]) -> Dict[str, str]:
        '''
        Gets list of target cell location and contents (considering shift)

        Arguments:
        - start_location: str - corner cell location of source area
        - end_location: str - corner cell location of source area
        - source_cells: Dict[str, str] - maps source cell locs to contents

        Returns:
        - Dict mapping str cell locations to str shifted contents

        '''
        
        target_top_left = get_coords_from_loc(to_location)

        start_col, start_row = get_coords_from_loc(start_location)
        end_col, end_row = get_coords_from_loc(end_location)
        top_left_coords = (min(start_col, end_col), min(start_row, end_row))

        col_diff = target_top_left[0] - top_left_coords[0]
        row_diff = target_top_left[1] - top_left_coords[1]
        coord_shift = (col_diff, row_diff)

        target_cells: Dict[str, str] = {}
        for source_loc, source_contents in source_cells.items():
            source_coords = get_coords_from_loc(source_loc)
            target_col = source_coords[0] + col_diff
            target_row = source_coords[1] + row_diff
            target_coords = (target_col, target_row)
            target_loc = get_loc_from_coords(target_coords) # checks boundaries
            
            target_cell = Cell(target_loc, self._evaluator)
            target_contents = target_cell.get_shifted_contents(source_contents,
                coord_shift)
            target_cells[target_loc] = target_contents

        return target_cells
