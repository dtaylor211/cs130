import re
from typing import Optional, List, Tuple, Any, Dict, Callable, Iterable

from .sheet import Sheet
from .evaluator import Evaluator
from .graph import Graph
from .helpers import get_loc_from_coords

class Workbook:
    '''
    A workbook containing zero or more named spreadsheets.

    Any and all operations on a workbook that may affect calculated cell
    values should cause the workbook's contents to be updated properly.

    '''

    # DEFAULT_SHEET_NUM = 1 - Plan to use in future projects

    def __init__(self):
        '''
        Initialize a new empty workbook.

        Contains (lowercase) sheet names (preserves order?)

        '''
        self.sheet_names = []

        # dictionary that maps lowercase sheet name to Sheet object
        self.sheet_objects: Dict[str, Sheet] = {}
        self.evaluator = Evaluator(self, '')
        self.notify_functions = []

    def num_sheets(self) -> int:
        '''
        Get the number of spreadsheets in the workbook

        Returns:
        - int number of sheets
        
        '''
        
        return len(self.sheet_names)

    def list_sheets(self) -> List[str]:
        '''
        Get list of the spreadsheet names in the workbook, with the
        capitalization specified at creation, and in the order that the sheets
        appear within the workbook.

        In this project, the sheet names appear in the order that the user
        created them; later, when the user is able to move and copy sheets,
        the ordering of the sheets in this function's result will also reflect
        such operations.

        A user should be able to mutate the return-value without affecting the
        workbook's internal state.

        Returns:
        - List of sheet names

        '''
        # Fixes "mutating the list comes back from list_sheets() shouldn't 
        # mangle the workbook's internal details"
        return [name for name in self.sheet_names]

    def new_sheet(self, sheet_name: Optional[str] = None) -> Tuple[int, str]:
        '''
        Add a new sheet to the workbook.  If the sheet name is specified, it
        must be unique.  If the sheet name is None, a unique sheet name is
        generated.  "Uniqueness" is determined in a case-insensitive manner,
        but the case specified for the sheet name is preserved.
        
        The function returns a tuple with two elements:
        (0-based index of sheet in workbook, sheet name).  This allows the
        function to report the sheet's name when it is auto-generated.
        
        If the spreadsheet name is an empty string (not None), or it is
        otherwise invalid, a ValueError is raised.

        Arguments:
        - sheet_name: Optional[str] (default None) new sheet's name

        Returns:
        - tuple with the zero-index of the sheet and the name

        '''

        # check if sheet name is specified
        if sheet_name is not None:
            # checking empty string
            if sheet_name == "":
                raise ValueError("Invalid Sheet name: cannot be empty string")
            # check whitespace
            elif sheet_name != sheet_name.strip():
                raise ValueError(
                    "Invalid Sheet name: cannot start/end with whitespace")
            # check valid (letters, numbers, spaces, .?!,:;!@#$%^&*()-_)
            elif not re.match(R'^[a-zA-Z0-9 .?!,:;!@#$%^&*\(\)\-\_]+$', 
                              sheet_name):
                raise ValueError("Invalid Sheet name: improper characters used")

            # check uniqueness
            if sheet_name.lower() in self.sheet_objects.keys():
                raise ValueError("Sheet name already exists")
            
        # sheet name not specified -> generate ununused "Sheet" + "#" name
        else:
            curr_sheet_num = 1
            sheet_name = "Sheet1"
            while sheet_name.lower() in self.sheet_objects.keys():
                curr_sheet_num += 1
                sheet_name = "Sheet" + str(curr_sheet_num)

        # update list of sheet names and sheet dictionary
        self.sheet_names.append(sheet_name) # preserves case
        self.sheet_objects[sheet_name.lower()] = Sheet(sheet_name, 
                                                       self.evaluator)

        self.update_cell_values(sheet_name.lower(), None)
        return self.num_sheets() - 1, sheet_name

    def del_sheet(self, sheet_name: str) -> None:
        '''
        Delete the spreadsheet with the specified name.
        
        The sheet name match is case-insensitive; the text must match but the
        case does not have to.
        
        If the specified sheet name is not found, a KeyError is raised.

        Arguments:
        - sheet_name: str - sheet's name

        '''
        
        if sheet_name.lower() not in self.sheet_objects.keys():
            raise KeyError("Specified sheet name is not found")
        
        original_sheet_name = self.sheet_objects[sheet_name.lower()].get_name()

        del self.sheet_objects[sheet_name.lower()]
        self.sheet_names.remove(original_sheet_name)
        # update all cells dependent on deleted sheet
        self.update_cell_values(sheet_name.lower(), None)

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        '''
        Get the current extent of the specified spreadsheet.
        
        The sheet name match is case-insensitive; the text must match but the
        case does not have to.
        
        If the specified sheet name is not found, a KeyError is raised.

        Arguments:
        - sheet_name: str - sheet's name

        Returns:
        - tuple (num-cols, num-rows) with extent

        '''
        
        if sheet_name.lower() not in self.sheet_objects.keys():
            raise KeyError("Specified sheet name is not found")

        # get_extent() should be a function of Sheet object (implemented in spreadsheet.py)
        return self.sheet_objects[sheet_name.lower()].get_extent() 

    def set_cell_contents(self, sheet_name: str, location: str,
                          contents: Optional[str]) -> None:
        '''
        Set the contents of the specified cell on the specified sheet.
        
        The sheet name match is case-insensitive; the text must match but the
        case does not have to.  Additionally, the cell location can be
        specified in any case.
        
        If the specified sheet name is not found, a KeyError is raised.
        If the cell location is invalid, a ValueError is raised.
        
        A cell may be set to "empty" by specifying a contents of None.
        
        Leading and trailing whitespace are removed from the contents before
        storing them in the cell.  Storing a zero-length string "" (or a
        string composed entirely of whitespace) is equivalent to setting the
        cell contents to None.
        
        If the cell contents appear to be a formula, and the formula is
        invalid for some reason, this method does not raise an exception;
        rather, the cell's value will be a CellError object indicating the
        naure of the issue.

        Arguments:
        - sheet_name: str - sheet's name
        - location: str - cell location
        - contents: Optional[str] - either string of instructions to set
            contents or None

        '''

        self.evaluator.set_working_sheet(sheet_name)

        sheet_name = sheet_name.lower()

        if sheet_name not in self.sheet_objects.keys():
            raise KeyError("Specified sheet name is not found")
        
        # set cell contents
        cell = self.sheet_objects[sheet_name].set_cell_contents(
            location, contents)

        # update other cells
        self.update_cell_values(sheet_name, location.lower())

    def get_cell_contents(self, sheet_name: str, location: str)-> Optional[str]:
        '''
        Return the contents of the specified cell on the specified sheet.
        
        The sheet name match is case-insensitive; the text must match but the
        case does not have to.  Additionally, the cell location can be
        specified in any case.
        
        If the specified sheet name is not found, a KeyError is raised.
        If the cell location is invalid, a ValueError is raised.
        
        Any string returned by this function will not have leading or trailing
        whitespace, as this whitespace will have been stripped off by the
        set_cell_contents() function.
        
        This method will never return a zero-length string; instead, empty
        cells are indicated by a value of None.

        Arguments:
        - sheet_name: str - sheet's name
        - location: str - cell's location

        Returns: 
        - either string contents or None

        '''

        self.evaluator.set_working_sheet(sheet_name)

        sheet_name = sheet_name.lower()

        if sheet_name not in self.sheet_objects.keys():
            raise KeyError("Specified sheet name is not found")
        
        return self.sheet_objects[sheet_name].get_cell_contents(location)

    def get_cell_value(self, sheet_name: str, location: str) -> Any:
        '''
        Return the evaluated value of the specified cell on the specified
        sheet.
        
        The sheet name match is case-insensitive; the text must match but the
        case does not have to.  Additionally, the cell location can be
        specified in any case.
        
        If the specified sheet name is not found, a KeyError is raised.
        If the cell location is invalid, a ValueError is raised.
        
        The value of empty cells is None.  Non-empty cells may contain a
        value of str, decimal.Decimal, or CellError.
        
        Decimal values will not have trailing zeros to the right of any
        decimal place, and will not include a decimal place if the value is a
        whole number.  For example, this function would not return
        Decimal('1.000'); rather it would return Decimal('1').

        Arguments:
        - sheet_name: str - sheet's name
        - location: str - cell's location

        Returns: 
        - the cell value

        '''

        sheet_name = sheet_name.lower()

        if sheet_name not in self.sheet_objects.keys():
            raise KeyError("Specified sheet name is not found")
        
        # calls get_cell_value from Sheet
        return self.sheet_objects[sheet_name].get_cell_value(location)

    def update_cell_values(self, updated_sheet: str, 
                                updated_cell: Optional[str]) -> None:
        '''
        Updates the contents of all cells. If given a sheet and/or cell,
        only updates cells effected.

        Arguments:
        - updated_sheet - sheet that has been updated
        - updated_cell - cell that has been updated

        '''
        # get all the cell children
        adjacency_list = {}
        for sheet in self.sheet_objects.values():
            adjacency_list.update(sheet.get_cell_adjacency_list())
        # make a graph of cell children, transpose to get graph of cell parents
        cell_graph = Graph(adjacency_list)
        cell_graph.transpose()
        # get cells to update if only given a sheet
        if updated_cell is None:
            updated_cells = [(child_sheet, child_cell) 
            for children in adjacency_list.values() 
            for (child_sheet, child_cell) in children 
            if child_sheet == updated_sheet]
        else:
            updated_cells = [(updated_sheet, updated_cell)]
        # get the graph of only cells needing to be updated
        reachable = cell_graph.get_reachable_nodes(updated_cells)
        cell_graph.subgraph_from_nodes(reachable)
        # get the acyclic components from the scc
        components = cell_graph.get_strongly_connected_components()
        dag_nodes = set()
        # if nodes are part of cycle make them a circlular reference
        # else add them to dag graph
        for component in components:
            if len(component) == 1 and component[0] not in adjacency_list[component[0]]:
                dag_nodes.add(component[0])
            else:
                for sheet, cell in component:
                    self.sheet_objects[sheet].get_cell(cell).set_circular_error()
        cell_graph.subgraph_from_nodes(dag_nodes)
        # get the topological sort of non-circular nodes needing to be updated
        cell_topological = cell_graph.topological_sort()
        # get cells to notify
        notify_cells = []
        # update cells
        for sheet, cell in cell_topological:
            if (sheet, cell) not in updated_cells:
                notify_cells.append((sheet, cell))
                self.sheet_objects[sheet].set_cell_contents(cell, 
                    self.sheet_objects[sheet].get_cell_contents(cell))
        for notify_function in self.notify_functions:
            try:
                notify_function(self, notify_cells)
            except:
                pass

    def notify_cells_changed(self, notify_function: 
                Callable[['Workbook', Iterable[Tuple[str, str]]], None]) -> None:
        '''
        Request that all changes to cell values in the workbook are reported
        to the specified notify_function.  The values passed to the notify
        function are the workbook, and an iterable of 2-tuples of strings,
        of the form ([sheet name], [cell location]).  The notify_function is
        expected not to return any value; any return-value will be ignored.
        
        Multiple notification functions may be registered on the workbook;
        functions will be called in the order that they are registered.
        
        A given notification function may be registered more than once; it
        will receive each notification as many times as it was registered.
        
        If the notify_function raises an exception while handling a
        notification, this will not affect workbook calculation updates or
        calls to other notification functions.
        
        A notification function is expected to not mutate the workbook or
        iterable that it is passed to it.  If a notification function violates
        this requirement, the behavior is undefined.

        Arguments:
        - notify_function: Callable - a function to call on cell updates

        Returns: 
        - None

        '''
        self.notify_functions.append(notify_function)
