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
    - get_target_cells(object, str, str, str, List[str]) -> Dict[str, str]

'''


from typing import Dict, List, Tuple, Optional, Any

from .cell import Cell
from .evaluator import Evaluator
from .utils import get_loc_from_coords, get_coords_from_loc, get_tl_br_corners


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

        evaluator = self.get_evaluator()
        if evaluator is not None:
            evaluator.set_working_sheet(self.get_name())

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
            adj_list[(name, cell.get_loc().upper())] = cell.get_children()
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

    def get_target_cells(self, start_location: str, end_location: str,
            to_location: str, source_cells: List[str]) -> Dict[str, str]:
        '''
        Gets list of target cell location and contents (considering shift)

        Arguments:
        - start_location: str - corner cell location of source area
        - end_location: str - corner cell location of source area
        - source_cells: List[str] - maps source cell locs to contents

        Returns:
        - Dict mapping str cell locations to str shifted contents

        '''

        target_top_left = get_coords_from_loc(to_location)

        src_top_left = get_tl_br_corners(start_location, end_location)[0]

        diff_coords = (
            target_top_left[0] - src_top_left[0],
            target_top_left[1] - src_top_left[1]
        )

        target_cells: Dict[str, str] = {}
        for source_loc in source_cells:
            source_coords = get_coords_from_loc(source_loc)
            target_coords = (
                source_coords[0] + diff_coords[0],
                source_coords[1] + diff_coords[1]
            )
            try:
                target_contents = self.get_all_cells()[source_coords]\
                    .get_shifted_contents(diff_coords)
            except KeyError:
                target_contents = None
            target_cells[get_loc_from_coords(target_coords)] = target_contents

        return target_cells
