import enum
import lark
from lark import Visitor, Tree
from decimal import Decimal, DecimalException
from typing import Optional, List

from .evaluator import Evaluator
from .cell_error import CellError, CellErrorType, CELL_ERRORS

RESTRICTED_VALUES = [
    Decimal('Infinity'),
    Decimal('-Infinity'), 
    Decimal('NaN'),
    Decimal('-NaN')
]

class _CellTreeVisitor(Visitor):
    '''
    This visitor gets all children cells from the tree of a cell.

    '''

    def __init__(self, sheet: str):
        self.children = set()
        self.sheet = sheet

    def cell(self, tree: Tree):
        if len(tree.children) == 2:
            cell_sheet = str(tree.children[0])
            cell = str(tree.children[1])
        else:
            cell_sheet = self.sheet
            cell = str(tree.children[0])
        self.children.add((cell_sheet.lower(), cell.lower()))


class Cell:
    '''
    A cell containing values of CellType and their string contents.
      
    This class represents an individual cell in a spreadsheet.
    Stores the string contents as well as the value
    Stores the type of the value as a CellType

    '''
    
    def __init__(self, loc: str, evaluator: Evaluator):
        '''
        Initialize a new Cell object

        Arguments:
        - loc: str - alphanumeric code representing location on a sheet (B2)
        - evaluator: Evaluator - lark formula evaluator

        '''

        self.loc = loc

        # new Cell is treated as an empty cell, contents and values are None
        self.contents = None
        self.value = None
        self.children = []
        self.evaluator = evaluator
        self.parser = lark.Lark.open(
            'formulas.lark', start='formula', rel_to=__file__)


    def set_contents(self, input_str: Optional[str]):
        '''
        Set the contents of the cell.

        Arguments:
        - input_str: str - specifications to set new cell value

        '''

        # Remove leading and trailing whitespace
        inp = input_str.strip()
        self.contents = inp

        try:

            # Check if there is a leading single quote, set to STRING type
            if inp[0] == "'":
                self.value = inp[1:].strip() # Remove whitespace again

            # Check if there is a leading equal sign, set to FORMULA type
            # and evaluate
            elif inp[0] == "=":
                tree = self.parser.parse(inp)
                visitor = _CellTreeVisitor(str(self.evaluator.working_sheet))
                visitor.visit(tree)
                self.children = list(visitor.children)
                eval = self.evaluator.transform(tree).children[0]
                self.value = eval

            elif inp.upper() in list(CELL_ERRORS.values()):
                e_type = [i for i in CELL_ERRORS if CELL_ERRORS[i]==inp.upper()]
                c = CellErrorType(e_type[0])
                self.value = CellError(c, '', None)

            # Otherwise set to NUMBER type - works for now, will need to change
            # if we can have other cell types
            else:
                temp_value = Decimal(inp) 
                if temp_value in RESTRICTED_VALUES:
                    self.value = inp
                else: self.value = temp_value

        except DecimalException as d:
            self.value = inp
        
        except lark.exceptions.LarkError as l:
            self.value = CellError(CellErrorType.PARSE_ERROR, 
                            detail='Unable to parse entry', exception = l)


    def get_children(self) -> List:
        '''
        Gets the children of the cell

        Returns:
        - List of the children

        '''

        return self.children


    def empty(self):
        '''
        Empty the contents of a cell

        '''

        self.contents = None
        self.value = None


    def get_string_from_error(self, cell_error_type: CellErrorType) -> str:
        '''
        Get the string representation of an error from its CellErrorType

        Arguments:
        - cell_error_type: CellErrorType - type of the error to get string of

        Returns:
        - string of error type

        '''
        
        return CELL_ERRORS[cell_error_type]


    def set_circular_error(self):
        '''
        Set the value of a cell to be a circular reference

        '''

        self.value = CellError(CellErrorType.CIRCULAR_REFERENCE, 
                                "Cell is in a circular reference.")
