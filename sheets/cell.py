import enum
import lark
from lark import Visitor, Tree
from decimal import Decimal, DecimalException
from typing import Optional, List, Tuple, Any

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

        self._loc = loc

        # new Cell is treated as an empty cell, contents and values are None
        self._contents = None
        self._value = None
        self._children = []
        self._evaluator = evaluator
        self._parser = lark.Lark.open(
            'formulas.lark', start='formula', rel_to=__file__)

    def get_loc(self) -> str:
        '''
        Get the location of the cell

        Returns:
        - string location of the cell

        '''

        return self._loc
    
    def get_contents(self) -> str:
        '''
        Get the contents of the cell

        Returns:
        - string contents instructions of the cell

        '''

        return self._contents
    
    def get_value(self) -> str:
        '''
        Get the value of the cell

        Returns:
        - value of the cell with various type

        '''

        return self._value
    
    def get_parser_and_evaluator(self) -> Tuple:
        '''
        Get the parser and evaluator

        Returns:
        - tuple with [parser, evaluator]

        '''

        return self._parser, self._evaluator
    
    def set_contents_and_value(self, contents: Optional[str], 
                                                        value: Optional[Any]):
        '''
        Set the contents and value fields of the cell

        Arguments:
        - contents: Optional[str] - contents to set, can be None
        - value: Optional[str] - value to set, can be None

        '''

        self._contents = contents
        self._value = value

    def set_contents(self, input_str: Optional[str]):
        '''
        Set the contents of the cell.

        Arguments:
        - input_str: str - specifications to set new cell value

        '''

        # Remove leading and trailing whitespace
        inp = input_str.strip()
        contents = inp

        try:

            # Check if there is a leading single quote, set to STRING type
            if inp[0] == "'":
                value = inp[1:].strip() # Remove whitespace again
                self.set_contents_and_value(contents, value)

            # Check if there is a leading equal sign, set to FORMULA type
            # and evaluate
            elif inp[0] == "=":
                parser, evaluator = self.get_parser_and_evaluator()
                tree = parser.parse(inp)
                visitor = _CellTreeVisitor(str(evaluator.working_sheet))
                visitor.visit(tree)
                self.children = list(visitor.children)
                eval = evaluator.transform(tree).children[0]
                self.set_contents_and_value(contents, eval)

            elif inp.upper() in list(CELL_ERRORS.values()):
                e_type = [i for i in CELL_ERRORS if CELL_ERRORS[i]==inp.upper()]
                c = CellErrorType(e_type[0])
                self.set_contents_and_value(contents, CellError(c, '', None))

            # Otherwise set to NUMBER type - works for now, will need to change
            # if we can have other cell types
            else:
                temp_value = Decimal(inp) 
                if temp_value in RESTRICTED_VALUES:
                    temp_value = inp
                self.set_contents_and_value(contents, temp_value)

        except DecimalException as d:
            self.set_contents_and_value(contents, inp)
        
        except lark.exceptions.LarkError as l:
            value = CellError(CellErrorType.PARSE_ERROR, 
                            detail='Unable to parse entry', exception = l)
            self.set_contents_and_value(contents, value)

    def get_children(self) -> List:
        '''
        Gets the children of the cell

        Returns:
        - List of the children

        '''

        return self._children

    def empty(self):
        '''
        Empty the contents of a cell

        '''

        self.set_contents_and_value(None, None)

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
