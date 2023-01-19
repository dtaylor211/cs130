import re
from typing import Optional, List, Tuple, Any, Dict
from .sheet import Sheet
from .formula_evaluator import Evaluator

class Workbook:
    # A workbook containing zero or more named spreadsheets.
    #
    # Any and all operations on a workbook that may affect calculated cell
    # values should cause the workbook's contents to be updated properly.

    DEFAULT_SHEET_NUM = 1

    def __init__(self):
        # Initialize a new empty workbook.

        # contains (lowercase) sheet names (preserves order?)
        self.sheet_names = [] # needs to be all lowercase?

        # dictionary that maps lowercase sheet name to Sheet object
        self.sheet_objects: Dict[str, Sheet] = {}
        self.evaluator = Evaluator(self, '')

    def num_sheets(self) -> int:
        # Return the number of spreadsheets in the workbook.
        
        return len(self.sheet_names)

    def list_sheets(self) -> List[str]:
        # Return a list of the spreadsheet names in the workbook, with the
        # capitalization specified at creation, and in the order that the sheets
        # appear within the workbook.
        #
        # In this project, the sheet names appear in the order that the user
        # created them; later, when the user is able to move and copy sheets,
        # the ordering of the sheets in this function's result will also reflect
        # such operations.
        #
        # A user should be able to mutate the return-value without affecting the
        # workbook's internal state.
        
        # assume that each Sheet object has a name attribute
        return self.sheet_names

    def new_sheet(self, sheet_name: Optional[str] = None) -> Tuple[int, str]:
        # Add a new sheet to the workbook.  If the sheet name is specified, it
        # must be unique.  If the sheet name is None, a unique sheet name is
        # generated.  "Uniqueness" is determined in a case-insensitive manner,
        # but the case specified for the sheet name is preserved.
        #
        # The function returns a tuple with two elements:
        # (0-based index of sheet in workbook, sheet name).  This allows the
        # function to report the sheet's name when it is auto-generated.
        #
        # If the spreadsheet name is an empty string (not None), or it is
        # otherwise invalid, a ValueError is raised.

        # check if sheet name is specified
        if sheet_name is not None:
            # checking empty string
            if sheet_name == "":
                raise ValueError("Invalid Sheet name: cannot be empty string")
            # check whitespace
            elif sheet_name != sheet_name.strip():
                raise ValueError("Invalid Sheet name: cannot start/end with whitespace")
            # check valid characters (letters, numbers, spaces, .?!,:;!@#$%^&*()-_)
            elif not re.match(R'^[a-zA-Z0-9 .?!,:;!@#$%^&*\(\)\-\_]+$', sheet_name):
                raise ValueError("Invalid Sheet name: improper characters used")

            # check uniqueness
            if sheet_name.lower() in self.sheet_objects.keys():
                raise ValueError("Sheet name already exists")
            
        # sheet name not specified -> generate ununused "Sheet" + "#" name
        # EXPLICITLY TEST THIS
        else:
            curr_sheet_num = 1
            sheet_name = "Sheet1"
            while sheet_name.lower() in self.sheet_objects.keys():
                curr_sheet_num += 1
                sheet_name = "Sheet" + str(curr_sheet_num)
            # curr_sheet_num += 1

        # update list of sheet names and sheet dictionary
        self.sheet_names.append(sheet_name) # preserves case
        self.sheet_objects[sheet_name.lower()] = Sheet(sheet_name, 
                                                       self.evaluator)

        # update evaluator to now have current sheet as the new sheet
        self.evaluator.working_sheet = sheet_name

        return self.num_sheets() - 1, sheet_name


    def del_sheet(self, sheet_name: str) -> None:
        # Delete the spreadsheet with the specified name.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        
        if sheet_name.lower() not in self.sheet_objects.keys():
            raise KeyError("Specified sheet name is not found")
        
        original_sheet_name = self.sheet_objects[sheet_name.lower()].name
        del self.sheet_objects[sheet_name.lower()]
        self.sheet_names.remove(original_sheet_name)

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        # Return a tuple (num-cols, num-rows) indicating the current extent of
        # the specified spreadsheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        
        if sheet_name.lower() not in self.sheet_objects.keys():
            raise KeyError("Specified sheet name is not found")

        # get_extent() should be a function of Sheet object (implemented in spreadsheet.py)
        return self.sheet_objects[sheet_name.lower()].get_extent() 

    def set_cell_contents(self, sheet_name: str, location: str,
                          contents: Optional[str]) -> None:
        # Set the contents of the specified cell on the specified sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # A cell may be set to "empty" by specifying a contents of None.
        #
        # Leading and trailing whitespace are removed from the contents before
        # storing them in the cell.  Storing a zero-length string "" (or a
        # string composed entirely of whitespace) is equivalent to setting the
        # cell contents to None.
        #
        # If the cell contents appear to be a formula, and the formula is
        # invalid for some reason, this method does not raise an exception;
        # rather, the cell's value will be a CellError object indicating the
        # naure of the issue.

        sheet_name = sheet_name.lower()

        if sheet_name not in self.sheet_objects.keys():
            raise KeyError("Specified sheet name is not found")
        
        return self.sheet_objects[sheet_name].set_cell_contents(location, contents)

    def get_cell_contents(self, sheet_name: str, location: str) -> Optional[str]:
        # Return the contents of the specified cell on the specified sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # Any string returned by this function will not have leading or trailing
        # whitespace, as this whitespace will have been stripped off by the
        # set_cell_contents() function.
        #
        # This method will never return a zero-length string; instead, empty
        # cells are indicated by a value of None.

        sheet_name = sheet_name.lower()

        if sheet_name not in self.sheet_objects.keys():
            raise KeyError("Specified sheet name is not found")
        
        return self.sheet_objects[sheet_name].get_cell_contents(location)

    def get_cell_value(self, sheet_name: str, location: str) -> Any:
        # Return the evaluated value of the specified cell on the specified
        # sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # The value of empty cells is None.  Non-empty cells may contain a
        # value of str, decimal.Decimal, or CellError.
        #
        # Decimal values will not have trailing zeros to the right of any
        # decimal place, and will not include a decimal place if the value is a
        # whole number.  For example, this function would not return
        # Decimal('1.000'); rather it would return Decimal('1').
        
        sheet_name = sheet_name.lower()

        if sheet_name not in self.sheet_objects.keys():
            raise KeyError("Specified sheet name is not found")

        # calls get_cell_value from Sheet
        return self.sheet_objects[sheet_name].get_cell_value(location)