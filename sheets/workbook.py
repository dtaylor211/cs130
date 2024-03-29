'''
Workbook

This module is where all operations on a workbook takes place, including
updating and altering cells and sheet names.

See the Sheet and Cell modules for more detailed commentary.

Classes:
- Workbook

    Attributes:
    - notify_functions (Callable[['Workbook', Iterable[Tuple[str, str]]], None])
    - evaluator (Evaluator)

    Methods:
    - list_sheets(object) -> List[str]
    - get_sheet_objects(object) -> Dict[str, Sheet]
    - num_sheets(object) -> int
    - new_sheet(object, Optional[str]) -> Tuple[int, str]
    - del_sheet(object, str) -> None
    - get_sheet_extent(object, str) -> Tuple[int, int]
    - set_cell_contents(object, str, str, Optional[str]) -> None
    - get_cell_contents(object, str, str) -> Optional[str]
    - get_cell_value(object, str, str) -> Any
    - update_cell_values(object, str, Optional[str], Optional[str],
        Optional[bool]) -> None
    - save_workbook(object, TextIO) -> None
    - notify_cells_changed(object, Callable[[Workbook,
        Iterable[Tuple[str, str]]], None]) -> None
    - rename_sheet(object, str, str) -> None
    - move_sheet(object, str, int) -> None
    - copy_sheet(object, str) -> Tuple[int, str]
    - move_cells(object, str, str, str, str, Optional[str]) -> None
    - copy_cells(object, str, str, str, str, Optional[str]) -> None
    - sort_region(object, str, str, str, List) -> None

    Static Methods:
    - load_workbook(TextIO) -> Workbook

'''

# pylint: disable=too-many-lines

# due to the complexity of workbook, we require just over the specified limit
# of 1000, but want to keep this limit across all other files

import re
import json
from typing import Optional, List, Tuple, Any, Dict, Callable, Iterable, TextIO

from .sheet import Sheet
from .evaluator import Evaluator
from .graph import Graph
from .utils import get_loc_from_coords, get_source_cells, get_tl_br_corners
from .sort_handler import Row


class Workbook:
    '''
    A workbook containing zero or more named spreadsheets.

    Any and all operations on a workbook that may affect calculated cell
    values should cause the workbook's contents to be updated properly.

    '''

    def __init__(self):
        '''
        Initialize a new empty workbook.
        Contains (lowercase) sheet names, in a preserved order

        '''

        self.evaluator = Evaluator(self, '')
        self._notify_cells = set()
        self._notify_functions = []
        self._update_cells = set()
        self._sheet_names = []
        self._sheet_objects: Dict[str, Sheet] = {}

    ########################################################################
    # Getters and Setters
    ########################################################################

    def list_sheets(self) -> List[str]:
        '''
        Get list of the spreadsheet names in the workbook, with the
        capitalization specified at creation, and in the order that the sheets
        appear within the workbook.

        In this project, the sheet names appear in the order that the user
        created them.

        A user should be able to mutate the return-value without affecting the
        workbook's internal state.

        Returns:
        - List of sheet names

        '''

        return list(self._sheet_names)

    def __set_sheet_names(self, sheet_names: List[str]) -> None:
        '''
        Set the sheet names list

        Private method, as the user should not be able to manually set the
        list of sheet names

        Arguments:
        - sheet_names: List[str] - new list of sheet names

        '''

        self._sheet_names = sheet_names

    def get_sheet_objects(self) -> Dict[str, Sheet]:
        '''
        Get the current dictionary of sheet objects

        Returns:
        - Dict of sheet names to the corresponding sheet object

        '''

        return dict(self._sheet_objects)

    def __set_sheet_objects(self, sheet_objects: Dict[str, Sheet]) -> None:
        '''
        Set the sheet objects dictionary to given dictionary

        Private method, as the user should not be able to manually set the
        dict of sheet objects

        Arguments:
        - sheet_objects: Dict[str, Sheet] - new dictionary of sheet objects

        '''

        self._sheet_objects = sheet_objects

    ########################################################################
    # Base Functionality
    ########################################################################

    def num_sheets(self) -> int:
        '''
        Get the number of spreadsheets in the workbook

        Returns:
        - int number of sheets

        '''

        return len(self.list_sheets())

    def new_sheet(self, sheet_name: Optional[str] = None) -> Tuple[int, str]:
        '''
        Add a new sheet to the workbook.  Sheet names must be unique
        (case-insensitive). If the sheet name is None, a unique sheet name
        is generated.

        The function returns a tuple with two elements:
        (0-based index of sheet in workbook, sheet name).

        If the spreadsheet name is an empty string (not None), or it is
        otherwise invalid, a ValueError is raised.

        Arguments:
        - sheet_name: Optional[str] (default None) new sheet's name

        Returns:
        - tuple with the zero-index of the sheet and the name

        '''

        sheet_names = self.list_sheets()
        sheet_objects = self.get_sheet_objects()

        if sheet_name is not None:

            if sheet_name == "":
                raise ValueError("Invalid Sheet name: cannot be empty string")
            # check whitespace
            if sheet_name != sheet_name.strip():
                raise ValueError(
                    "Invalid Sheet name: cannot start/end with whitespace")
            if not re.match(R'^[a-zA-Z0-9 .?!,:;!@#$%^&*\(\)\-\_]+$',
                              sheet_name):
                raise ValueError("Invalid Sheet name: improper characters used")

            self.__validate_sheet_uniqueness(sheet_name)

        # sheet name not specified -> generate ununused "Sheet" + "#" name
        else:
            curr_sheet_num = 1
            sheet_name = "Sheet1"
            while sheet_name.lower() in sheet_objects:
                curr_sheet_num += 1
                sheet_name = "Sheet" + str(curr_sheet_num)

        sheet_names.append(sheet_name)
        sheet_objects[sheet_name.lower()] = Sheet(
            sheet_name, self.evaluator)

        self.__set_sheet_names(sheet_names) # preserves case
        self.__set_sheet_objects(sheet_objects)

        self.update_cell_values(sheet_name)
        self.__notify()
        return self.num_sheets() - 1, sheet_name

    def del_sheet(self, sheet_name: str) -> None:
        '''
        Delete the spreadsheet with the specified name (case-insensitive).

        If the specified sheet name is not found, a KeyError is raised.

        Arguments:
        - sheet_name: str - sheet's name

        '''

        sheet_names = self.list_sheets()
        sheet_objects = self.get_sheet_objects()
        self.__validate_sheet_existence(sheet_name)

        original_sheet_name = sheet_objects[sheet_name.lower()].get_name()
        sheet_names.remove(original_sheet_name)
        del sheet_objects[sheet_name.lower()]

        self.__set_sheet_names(sheet_names)
        self.__set_sheet_objects(sheet_objects)
        # update all cells dependent on deleted sheet
        self.update_cell_values(sheet_name)
        self.__notify()

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        '''
        Get the current extent of the specified spreadsheet (case-insensitive).

        If the specified sheet name is not found, a KeyError is raised.

        Arguments:
        - sheet_name: str - sheet's name

        Returns:
        - tuple (num-cols, num-rows) with extent

        '''

        sheet_objects = self.get_sheet_objects()
        self.__validate_sheet_existence(sheet_name)

        return sheet_objects[sheet_name.lower()].get_extent()

    def set_cell_contents(self, sheet_name: str, location: str,
                          contents: Optional[str], notify: Optional[bool] = True
                          ) -> None:
        '''
        Set the contents of the specified cell on the specified sheet
        (case-insensitive).

        If the specified sheet name is not found, a KeyError is raised.
        If the cell location is invalid, a ValueError is raised.

        A cell may be set to "empty" by specifying a contents of None or
        storing a zero-length string "".

        If the cell contents appear to be a formula, and the formula is
        invalid for some reason, the cell's value will be a CellError object
        indicating the naure of the issue.

        Arguments:
        - sheet_name: str - sheet's name
        - location: str - cell location
        - contents: Optional[str] - either string of instructions to set
            contents or None

        '''

        sheet_objects = self.get_sheet_objects()
        sheet_name_lower = sheet_name.lower()
        self.__validate_sheet_existence(sheet_name_lower)

        prev_contents = sheet_objects[sheet_name_lower].get_cell_contents(location)
        prev_value = sheet_objects[sheet_name_lower].get_cell_value(location)

        sheet_objects[sheet_name_lower].set_cell_contents(
            location, contents)
        new_value = sheet_objects[sheet_name_lower].get_cell_value(location)

        if notify:
            # update other cells
            if new_value == prev_value and prev_contents is not None:
                self.update_cell_values(sheet_name, [(sheet_name, location.upper())], notify=False)
            else:
                self.update_cell_values(sheet_name, [(sheet_name, location.upper())])
            self.__notify()
        else:
            if new_value != prev_value:
                self._update_cells.add((sheet_name, location.upper()))

    def get_cell_contents(self, sheet_name: str, location: str)-> Optional[str]:
        '''
        Return the contents of the specified cell on the specified sheet
        (case-insensitive).

        If the specified sheet name is not found, a KeyError is raised.
        If the cell location is invalid, a ValueError is raised.

        Arguments:
        - sheet_name: str - sheet's name
        - location: str - cell's location

        Returns:
        - either string contents or None

        '''

        sheet_objects = self.get_sheet_objects()
        self.evaluator.set_working_sheet(sheet_name)
        sheet_name = sheet_name.lower()
        self.__validate_sheet_existence(sheet_name)

        return sheet_objects[sheet_name].get_cell_contents(location)

    def get_cell_value(self, sheet_name: str, location: str) -> Any:
        '''
        Return the evaluated value of the specified cell on the specified
        sheet (case-insensitive).

        If the specified sheet name is not found, a KeyError is raised.
        If the cell location is invalid, a ValueError is raised.

        Cells may contain a value of str, decimal.Decimal, CellError, or None.

        Arguments:
        - sheet_name: str - sheet's name
        - location: str - cell's location

        Returns:
        - the cell value

        '''

        sheet_objects = self.get_sheet_objects()
        sheet_name = sheet_name.lower()
        self.__validate_sheet_existence(sheet_name)

        return sheet_objects[sheet_name].get_cell_value(location)

    def update_cell_values(self, updated_sheet: str, updated_cell: Optional[str]
        = None, renamed_sheet: Optional[str] = None, notify: Optional[bool] =
        True) -> None:
        '''
        Updates the contents of all cells. If given a sheet and/or cell,
        only updates cells effected.

        Arguments:
        - updated_sheet - sheet that has been updated
        - updated_cell - cell that has been updated
        - renamed_sheet - new name of renamed sheet with preserved case

        '''

        sheet_objects = self.get_sheet_objects()

        adj = {}
        for sheet in sheet_objects.values():
            adj.update(sheet.get_cell_adjacency_list())
        # make a graph of cell children, transpose to get graph of cell parents
        cell_graph = Graph(adj)
        cell_graph.transpose()

        # get cells to update if only given a sheet
        if updated_cell is None:
            # get the cells in the sheet
            updated_cells = [(child_sheet, child_cell)
            for children in adj.values()
            for (child_sheet, child_cell) in children
            if (renamed_sheet is None and child_sheet == updated_sheet)
            or (renamed_sheet is not None and child_sheet == renamed_sheet)]
            # rename references if we have a renamed sheet
            if renamed_sheet is not None:
                # fix new sheet name
                if re.search(R'[ .?!,:;!@#$%^&*\(\)\-]', renamed_sheet):
                    renamed_sheet = "'"+renamed_sheet+"'"
                # get the adjacency list of the cell parents graph
                parent_adj = cell_graph.get_adjacency_list()
                # get the cells that references to cells on sheet
                refer_cells = {(child_sheet, child_cell)
                for children in adj.values()
                for (child_sheet, child_cell) in children
                if child_sheet == updated_sheet}
                ref_cells = set()
                for ref in refer_cells:
                    for cell in parent_adj[ref]:
                        ref_cells.add(cell)
                # go through cells that reference the cells on sheet
                for (sheet, cell) in ref_cells:
                    # get cell contents
                    contents = self.get_cell_contents(sheet, cell)
                    # replace sheet name with new name
                    contents=re.sub(R"([=\+\-*/& ])"+updated_sheet+"!",
                                    R"\1"+renamed_sheet+"!",
                                    contents, flags=re.IGNORECASE)
                    contents=re.sub("'"+updated_sheet+"'"+"!",
                    renamed_sheet+"!", contents, flags=re.IGNORECASE)
                    # set the new contents with new sheet name
                    sheet_objects[sheet.lower()].set_cell_contents(
                        cell, contents)
                    # call helper function to update sheet names in contents
                    self.__format_sheet_names(sheet, cell,
                                                adj[(sheet, cell)])
                self.__set_sheet_objects(sheet_objects)
        else:
            updated_cells = updated_cell
        # call helper to update and notify cells that need updating
        self.__update_notify_cells(updated_cells,
            self.__get_topological(cell_graph, updated_cells, adj),
            notify)

    @staticmethod
    def load_workbook(fp: TextIO) -> 'Workbook':
        '''
        Load a workbook from a text file or file-like object in JSON format,
        and return the new Workbook instance.

        If the contents of the input cannot be parsed by the Python json
        module then a json.JSONDecodeError should be raised by the method.
        Let any raised exception propagate through.

        If any expected value in the input JSON is missing raise a KeyError.
        If any expected value in the input JSON is not of the proper type
        raise a TypeError.

        Arguments:
        - fp: TextIO - json file to load

        Returns:
        - a Workbook object with contents specified by fp

        '''

        loaded_wb = json.load(fp)
        new_wb = Workbook()

        if "sheets" not in loaded_wb:
            raise KeyError("Missing: 'sheets'")
        sheets = loaded_wb["sheets"]
        if not isinstance(sheets, list):
            raise TypeError("'sheets' is not proper type (list)")

        for sheet in sheets:
            if not isinstance(sheet, dict):
                raise TypeError(
                    "Sheet representation is not proper type (dict)")

            if "name" not in sheet:
                raise KeyError("Missing: 'name'")
            if not isinstance(sheet["name"], str):
                raise TypeError("Sheet name is not proper type (string)")
            sheet_name = sheet["name"]

            if "cell-contents" not in sheet:
                raise KeyError("Missing: 'cell-contents'")
            if not isinstance(sheet["cell-contents"], dict):
                raise TypeError("Cell-contents is not proper type (dictionary))")
            cell_contents = sheet["cell-contents"]

            new_wb.new_sheet(sheet_name)

            for location, contents in cell_contents.items():
                if not isinstance(location, str):
                    raise TypeError("Location is not proper type (string)")

                if not isinstance(contents, str):
                    raise TypeError("Contents is not proper type (string)")

                new_wb.set_cell_contents(sheet_name, location, contents)

        return new_wb

    def save_workbook(self, fp: TextIO) -> None:
        '''
        Save a workbook to a text file or file-like object in JSON format.

        Let any raised exception propagate through.

        Arguments:
        - fp: TextIO - write supporting file like object to save to

        '''

        obj = {}
        json_sheets = []
        sheet_names = self.list_sheets()
        sheet_objects = self.get_sheet_objects()

        for sheet_name in sheet_names: # preserves ordering
            sheet  = sheet_objects[sheet_name.lower()]
            json_sheets.append(sheet.save_sheet())

        obj = {"sheets": json_sheets}
        json.dump(obj=obj, fp=fp)

    def notify_cells_changed(self, notify_function:
        Callable[['Workbook', Iterable[Tuple[str, str]]], None]) -> None:
        '''
        Request that all changes to cell values in the workbook are reported
        to the specified notify_function.  The values passed to the notify
        function are the workbook, and an iterable of 2-tuples of strings,
        ([sheet name], [cell location]).

        Functions will be called in the order that they are registered.

        A given notification function may be registered more than once; it
        will receive each notification as many times as it was registered.

        Arguments:
        - notify_function: Callable[['Workbook', Iterable[Tuple], None]] -
            callable function to notify that a cell value is changed

        '''

        self._notify_functions.append(notify_function)

    def rename_sheet(self, sheet_name: str, new_sheet_name: str) -> None:
        '''
        Rename the specified sheet to the new sheet name.  All relevant
        cell formulas are updated with the new sheet name (case-insensitive).

        Case of the new_sheet_name is preserved by the workbook.

        If the sheet_name is not found, a KeyError is raised.
        If the new_sheet_name is invalid, a ValueError is raised.

        Arguments:
        - sheet_name: str - name of sheet to be renamed
        - new_sheet_name: str - new name to be used

        '''

        sheet_names = self.list_sheets()
        sheet_objects = self.get_sheet_objects()
        self.__validate_sheet_existence(sheet_name)

        # checking empty string
        if new_sheet_name == "":
            raise ValueError("Invalid Sheet name: cannot be empty string")

        if new_sheet_name != new_sheet_name.strip():
            raise ValueError(
                "Invalid Sheet name: cannot start/end with whitespace")
        if not re.match(R'^[a-zA-Z0-9 .?!,:;!@#$%^&*\(\)\-\_]+$',
                            new_sheet_name):
            raise ValueError("Invalid Sheet name: improper characters used")

        self.__validate_sheet_uniqueness(new_sheet_name)

        # Update sheet_names (list preserving order & case of sheet names)
        # old_sheet_name used to retrieve proper casing
        old_sheet_name = sheet_objects[sheet_name.lower()].get_name()
        old_sheet_idx = sheet_names.index(old_sheet_name)
        sheet_names[old_sheet_idx] = new_sheet_name
        self.__set_sheet_names(sheet_names)

        # Update sheet_objects dict (delete old key, add key with new name)
        sheet = sheet_objects[sheet_name.lower()]
        sheet.set_name(new_sheet_name)
        sheet_objects[new_sheet_name.lower()] = sheet
        del sheet_objects[sheet_name.lower()]
        self.__set_sheet_objects(sheet_objects)

        self.update_cell_values(sheet_name, renamed_sheet = new_sheet_name)
        self.__notify()


    def move_sheet(self, sheet_name: str, index: int) -> None:
        '''
        Move the specified sheet to the specified index in the workbook's
        ordered sequence of sheets (case-insensitive).

        If the specified sheet name is not found, a KeyError is raised.
        If the index is outside the valid range, an IndexError is raised.

        Arguments:
        - sheet_name: str - name of the sheet to be moved
        - index: int - index to move the desired sheet to

        '''

        sheet_names = self.list_sheets()
        sheet_objects = self.get_sheet_objects()
        self.__validate_sheet_existence(sheet_name)

        if index < 0 or index >= self.num_sheets():
            raise IndexError("Provided index is outside valid range")

        sheet_name = sheet_objects[sheet_name.lower()].get_name()
        sheet_names.remove(sheet_name)
        sheet_names.insert(index, sheet_name)
        self.__set_sheet_names(sheet_names)

    def copy_sheet(self, sheet_name: str) -> Tuple[int, str]:
        '''
        Make a copy of the specified sheet, storing the copy at the end of the
        workbook's sequence of sheets.  The copy's name is unique, and is
        generated by appending "_1", "_2", ... to the original sheet's name
        (case-insensitive).

        If the specified sheet name is not found, a KeyError is raised.

        Arguments:
        - sheet_name: str - name of the sheet to copy

        Returns:
        - Tuple of int, str holding the index of the sheet in the workbook as
            well as the copied sheet's name

        '''

        sheet_objects = self.get_sheet_objects()
        self.__validate_sheet_existence(sheet_name)

        og_sheet_name = sheet_objects[sheet_name.lower()].get_name()
        copy_num = 1
        sheet_copy_name = og_sheet_name + "_" + str(copy_num)
        while sheet_copy_name.lower() in sheet_objects:
            copy_num += 1
            sheet_copy_name = og_sheet_name + "_" + str(copy_num)

        # explicitly set each cell in (new) copy sheet using locations and
        # contents from copied sheet
        sheet_copy_idx, sheet_copy_name = self.new_sheet(sheet_copy_name)
        cells_dict = sheet_objects[sheet_name.lower()].get_all_cells()
        for coords, cell in cells_dict.items():
            loc = get_loc_from_coords(coords)
            self.set_cell_contents(sheet_copy_name, loc, cell.get_contents(),
                                    notify=False)

        self.update_cell_values(sheet_copy_name, list(self._update_cells))
        self._update_cells = set()
        self.__notify()
        return sheet_copy_idx, sheet_copy_name

    def move_cells(self, sheet_name: str, start_location: str,
            end_location: str, to_location: str, to_sheet: Optional[str] = None
            ) -> None:
        '''
        Move cells from one location to another, possibly to another sheet.
        All formulas being moved will have all relative/mixed cell-references
        updated by relative distance.

        Cells in the source area (and not target area) are set to empty.

        If any specified sheet name is not found, a KeyError is raised.
        If any cell location is invalid, a ValueError is raised.
        If the target area would extend outside the valid area of the
        spreadsheet , a ValueError is raised, and no changes are made.

        If a formula being moved contains a relative or mixed cell-reference
        that will become invalid after updating the cell-reference, then the
        cell-reference is replaced with a #REF! error-literal in the formula.

        Arguments:
        - sheet_name: str - name of sheet containing cells to be moved
        - start_location: str - location of corner of area of cells to be moved
        - end_location: str - location of opposite corner of area of cells to
            be moved
        - to_location: str - location of top left corner of area where cells
            are to be moved to
        - to_sheet: Optional[str] - either string of new sheet to move cells
            to or None (moving cells to source sheet)

        '''

        self.__validate_sheet_existence(sheet_name)
        sheet_objects = self.get_sheet_objects()

        source_sheet = sheet_objects[sheet_name.lower()]
        source_cells = get_source_cells(start_location, end_location)

        if to_sheet is None:
            to_sheet = sheet_name
        else:
            self.__validate_sheet_existence(to_sheet)

        target_cells = source_sheet.get_target_cells(start_location,
            end_location, to_location, source_cells) # Dict[locs, contents]

        # Set contents of source cells (not in target area) to None
        # could set all values in source area to None **
        if to_sheet is None or to_sheet == sheet_name:
            to_sheet = sheet_name
            source_set = set(source_cells)
            target_set = set(target_cells.keys())
            source_target_set_diff = source_set.difference(target_set)
        else:
            self.__validate_sheet_existence(to_sheet)
            source_target_set_diff = source_cells
        for loc in list(source_target_set_diff):
            self.set_cell_contents(sheet_name, loc, None, notify=False)

        # Set contents of target cells (within same sheet if to_sheet is None)
        for loc, contents in target_cells.items():
            self.set_cell_contents(to_sheet, loc, contents, notify=False)

        self.update_cell_values(to_sheet, list(self._update_cells))
        self._update_cells = set()
        self.__notify()


    def copy_cells(self, sheet_name: str, start_location: str,
            end_location: str, to_location: str, to_sheet: Optional[str] = None
            ) -> None:
        '''
        Copy cells from one location to another, possibly to another sheet.
        All formulas being copied will have relative/mixed cell-references
        updated by the relative distance.

        Cells in the source area (that are not also in the target area) are
        left unchanged by the copy operation.

        If any specified sheet name is not found, a KeyError is raised.
        If any cell location is invalid, a ValueError is raised.
        If the target area would extend outside the valid area of the
        spreadsheet, a ValueError is raised, and no changes are made.

        If a formula being copied contains a relative or mixed cell-reference
        that will become invalid after updating the cell-reference, then the
        cell-reference is replaced with a #REF! error-literal in the formula.

        Arguments:
        - sheet_name: str - name of the sheet containing cells to be copied
        - start_location: str - location of corner of area of cells to be
            copied
        - end_location: str - location of opposite corner of area of cells to
            be copied
        - to_location: str - location of top left corner of area where cells
            are to be copied to
        - to_sheet: Optional[str] - either string of new sheet to copy cells
            to or None

        '''

        self.__validate_sheet_existence(sheet_name)
        sheet_objects = self.get_sheet_objects()

        source_sheet = sheet_objects[sheet_name.lower()]
        source_cells = get_source_cells(start_location, end_location)

        if to_sheet is None:
            to_sheet = sheet_name
        else:
            self.__validate_sheet_existence(to_sheet)

        target_cells = source_sheet.get_target_cells(start_location,
            end_location, to_location, source_cells) # Dict[locs, contents]

        # Set contents of target cells (within same sheet if to_sheet is None)
        for loc, contents in target_cells.items():
            self.set_cell_contents(to_sheet, loc, contents, notify=False)

        self.update_cell_values(to_sheet, list(self._update_cells))
        self._update_cells = set()
        self.__notify()

    def sort_region(self, sheet_name: str, start_location: str,
                    end_location: str, sort_cols: List) -> None:
        '''
        Sort the specified region of a spreadsheet with a stable sort, using
        the specified columns for the comparison (case-insensitive).

        The sort_cols argument specifies one or more columns to sort on.
        Column indexes are one-based, and are positive for ascending order, or
        negative for descending order. Ordering occurs for first column index,
        then the second, etc. Sort_cols cannot be empty.

        The sorting implementation is a stable sort.

        If the specified sheet name is not found, a KeyError is raised.
        If any cell location is invalid, a ValueError is raised.
        If the sort_cols list is invalid in any way, a ValueError is raised.

        '''

        self.__validate_sheet_existence(sheet_name)
        sheet =  self.get_sheet_objects()[sheet_name.lower()]
        tl_br_corners = get_tl_br_corners(start_location, end_location)
        source_cells = get_source_cells(start_location, end_location)

        all_rows = self.__populate_row_sorter(tl_br_corners, source_cells, sheet_name)

        prev_cols = []
        for sort_col in sort_cols[::-1]:
            if abs(sort_col) in prev_cols:
                raise ValueError
            prev_cols.append(abs(sort_col))
            for row in all_rows:
                row.set_current_column(abs(sort_col))
            all_rows = sorted(all_rows, reverse=sort_col < 0)

        all_target_cells = self.__get_sorted_row_contents(all_rows, tl_br_corners, sheet)

        # Set contents of target cells
        for cell in all_target_cells.items():
            self.set_cell_contents(sheet_name, cell[0], cell[-1], notify=False)

        self.update_cell_values(sheet_name, list(self._update_cells))
        self._update_cells = set()
        self.__notify()

    ########################################################################
    # Private Helpers
    ########################################################################

    def __validate_sheet_existence(self, sheet_name: str) -> None:
        '''
        Validate whether the given sheet name already exists within the workbook

        Throw a KeyError if the name cannot be found

        Arguments:
        - sheet_name: str - name of the sheet to validate

        '''

        sheet_objects = self.get_sheet_objects()

        if sheet_name.lower() not in sheet_objects:
            raise KeyError(f"Specified sheet name '{sheet_name}' is not found")

    def __validate_sheet_uniqueness(self, sheet_name: str) -> None:
        '''
        Validate that the given sheet name does not already exist within the
        workbook

        Throw a ValueError if the name already exists

        Arguments:
        - sheet_name: str - name of the sheet to validate

        '''

        sheet_objects = self.get_sheet_objects()

        if sheet_name.lower() in sheet_objects:
            raise ValueError(f"Sheet name '{sheet_name}' already exists")

    def __format_sheet_names(self, sheet_name: str, location: str,
        sheets_in_contents: List[Tuple]) -> None:
        '''
        Set the cell's contents to be properly formatted so that sheet names
        with quotations that do not need them, will have them removed

        Arguments:
        - sheet_name: str - name of the current cell's sheet
        - location: str - location of the current cell
        - sheets_in_contents: List[Tuple] - list of the sheets referenced in
            the current cell's contents

        '''

        sheet_objects = self.get_sheet_objects()

        for sheet, _ in sheets_in_contents:
            if not re.search(R'[ .?!,:;!@#$%^&*\(\)\-]', sheet):
                curr_contents = self.get_cell_contents(sheet_name, location)
                contents = re.sub("'"+sheet+"'", sheet, curr_contents)
                sheet_objects[sheet_name.lower()]\
                    .set_cell_contents(location, contents)
        self.__set_sheet_objects(sheet_objects)

    def __get_topological(self, cell_graph: Graph, updated_cells: List[Tuple],
        adj: Dict[Tuple, List[Tuple]]) -> None:
        '''
        Updates and notifies cells using the topological sort provided.

        Arguments:
        - updated_cells - list of cells that were changed
        - cell_topological - list of cells affected by updated_cells in order

        '''

        # get the graph of only cells needing to be updated
        reachable = cell_graph.get_reachable_nodes(updated_cells)
        cell_graph.subgraph_from_nodes(reachable)

        # get the acyclic components from the scc
        components = cell_graph.get_strongly_connected_components()
        dag_nodes = set()
        sheet_objects = self.get_sheet_objects()

        # if nodes are part of cycle make them a circlular reference
        # else add them to dag graph
        for component in components:
            if len(component) == 1 and component[0] not in adj[component[0]]:
                dag_nodes.add(component[0])
            else:
                for sheet, cell in component:
                    sheet_objects[
                        sheet.lower()].get_cell(cell).set_circular_error()
                self.__set_sheet_objects(sheet_objects)

        cell_graph.subgraph_from_nodes(dag_nodes)

        # get the topological sort of non-circular nodes needing to be updated
        cell_topological = cell_graph.topological_sort()
        return cell_topological

    def __populate_row_sorter(self, tl_br_corners: List[Tuple[int, int]],
                        source_cells: List[str], sheet_name: str) -> List[Row]:
        '''
        Populate a list of Rows to be sorted

        Arguments:
        - tl_br_corners: List[Tuple[int, int]] - coordinates for the top left
            and bottom right corners
        - source_cells: List[str] - list of source cell locations
        - sheet_name: str - name of the sheet where cells are being sorted

        Returns:
        - list of Rows to be sorted

        '''

        row_orders = range(1, tl_br_corners[-1][1] - tl_br_corners[0][1] + 2)
        row_len = tl_br_corners[1][0] - tl_br_corners[0][0] + 1

        all_rows = []
        for row_idx in row_orders:
            cells = {}
            if row_len == 1:
                temp_cells = [source_cells[row_idx - 1]]
            else:
                temp_cells = source_cells[row_idx - 1 :: row_len]

            for col_idx, cell in enumerate(temp_cells):
                cells[col_idx + 1] = self.get_cell_value(sheet_name, cell)

            all_rows.append(Row(row_idx, cells, row_len))
        return all_rows

    def __get_sorted_row_contents(self, all_rows: List[Row],
                    tl_br_corners: List[Tuple], sheet: Sheet) -> Dict[str, str]:
        '''
        Get the shifted contents of a sorted list of rows

        Arguments:
        - all_rows: List[Row] - sorted rows
        - tl_br_corners: List[Tuple] - coordinates for the top left
            and bottom right corners
        - sheet: Sheet - sheet where sorting is occur

        Returns:
        - dictionary mapping cell locations to their sorted contents

        '''

        all_target_cells = {}
        for i, row in enumerate(all_rows):
            if row.row_order != i + 1:
                start_loc = get_loc_from_coords(
                    (tl_br_corners[0][0], tl_br_corners[0][-1]+row.row_order-1))
                end_loc = get_loc_from_coords(
                    (tl_br_corners[1][0], tl_br_corners[0][-1]+row.row_order-1))
                to_loc = get_loc_from_coords(
                    (tl_br_corners[0][0], tl_br_corners[0][-1]+i))
                source_cells = get_source_cells(start_loc, end_loc)
                all_target_cells.update(sheet.get_target_cells(start_loc,
                                                end_loc, to_loc, source_cells))

        return all_target_cells

    # pylint: disable=broad-exception-caught

    # We disable the checks for a general exception here, since we are
    # specified to not propogate an error that occurs in a given notify
    # function, and we are unsure of exception types

    def __update_notify_cells(self, updated_cells: List[Tuple],
        cell_topological: List[Tuple], notify: bool) -> None:
        '''
        Updates and notifies cells using the topological sort provided.

        Arguments:
        - updated_cells - list of cells that were changed
        - cell_topological - list of cells affected by updated_cells in order

        '''

        sheet_objects = self.get_sheet_objects()

        # get cells to notify
        notify_cells = []
        if notify:
            notify_cells = updated_cells

        # update cells
        for sheet, cell in cell_topological:
            if len(updated_cells) > 1 or (sheet, cell) not in updated_cells:
                name = sheet.lower()
                prev_value = sheet_objects[name].get_cell_value(cell)
                sheet_objects[name].set_cell_contents(cell,
                    sheet_objects[name].get_cell_contents(cell))
                new_value = sheet_objects[name].get_cell_value(cell)
                if new_value != prev_value:
                    notify_cells.append((sheet, cell))
        self.__set_sheet_objects(sheet_objects)
        self._notify_cells.update(notify_cells)

    def __notify(self):
        for notify_function in self._notify_functions:
            try:
                notify_function(self, list(self._notify_cells))
            except Exception:
                pass
        self._notify_cells = set()
