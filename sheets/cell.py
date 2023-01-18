<<<<<<< HEAD
class CellType:
    LITERAL: int = 1
=======
import enum

class _CellType(enum.Enum):
    '''
    This enum specifies the kinds of values that spreadsheet cells can hold.
    '''

    NUMBER: int = 1
>>>>>>> 40d1757... DJJ-8: Fixed formatting
    STRING: int = 2
    FORMULA: int = 3 # May want to remove
    EMPTY: int = 4

    # ERROR: int = 5 ?


class _Cell:
    '''
    A cell containing values of CellType and their string contents.
      
    This class represents an individual cell in a spreadsheet.
    Stores the string contents as well as the value
    Stores the type of the value as a CellType

    '''
    
    def __init__(self, loc: str):
        '''
        Initialize a new Cell object

        Arguments:
        - loc: str - alphanumeric code representing location on a sheet (B2)

        '''

        self.loc = loc

        # new Cell is treated as an empty cell, contents and values are None
        self.contents = None
        self.value = None
        self.type: int = _CellType.EMPTY

    def set_value(self, input_str: str):
        '''
        Set the value of the cell.

        Arguments:
        - input_str: str - specifications to set new cell value

        '''

        # Remove leading and trailing whitespace
        inp = input_str.strip()
    
        # Check if empty string
        if inp != "":

            self.contents = inp

            # Check if there is a leading single quote, set to STRING type
            if inp[0] == "'":
                self.type = _CellType.STRING
                self.value = inp[1:]

            # Check if there is a leading equal sign, set to FORMULA type
            # and evaluate
            elif inp[0] == "=":
                self.type = _CellType.FORMULA
                self.value = Formula(inp) #todo
                pass # evaluate formula here, maybe formula class

            # Otherwise set to LITERAL type
            else:
                self.type = _CellType.NUMBER # want to do a better job figure out expression to match input  Lark?
                self.value = inp # prob not what I want to do - maybe need string and int literals
        else: 
            self.contents = None
            self.value = None
            self.type = _CellType.EMPTY

    def empty(self):
        '''
        Empty the contents of a cell

        '''

        self.contents = None
        self.value = None
        self.type: int = CellType.EMPTY


g