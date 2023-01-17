from typing import Dict, Tuple
from .cell import Cell

class Sheet:
    # A spreadsheet containing zero or more cells.

    def __init__(self, sheet_name):
        # Initialize a new spreadsheet.
        
        self.name = sheet_name # could we add get_name() and use in workbook.py? 

        # dictonary that maps (row, col) tuple -> Cell
        # need to make sure inputted location "D14" is converted to (4, 14)
        ## location = D14, coords = (4,14)
        self.cells = Dict[Tuple[int, int], Cell] = {}

        self.extent = (0, 0) # do we need to initialize (0, 0) here? I think we can handle in get_extent.
    
    def get_extent(self) -> Tuple[int, int]:
        # Get the extent of spreadsheet (# rows, # cols).
        # New empty sheet has extent of (0, 0)

        if len(self.cells.keys()) == 0:
            return (0, 0) # empty sheet

        # Need to find max values for row/col values
        # Assume locations are not in order?  Better way to maintain order?
        coords = list(self.cells.keys())
        max_row = 0
        max_col = 0
        for tup in coords:
            row, col = tup
            if row > max_row:
                max_row = row
            if col > max_col:
                max_col 
        return (max_row, max_col)

    def get_cell(self, location: str) -> Optional[Cell]:
        pass

    def get_coords_from_loc(self, location: str) -> Tuple[int, int]:
        # raise ValueError is cell location isn't available
        pass

    def get_cell_value(self, location: str) -> Any:
        pass

    def get_cell_contents(self, location: str) -> Optional[str]:
        pass

    def set_cell_contents(self, location: str, contents: Optional[str]) -> None:
        pass

