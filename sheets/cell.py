import enum
from decimal import Decimal, DecimalException
from typing import Optional
from .formula_evaluator import Evaluator
import lark
from .cell_error import CellError, CellErrorType, cell_error_dict


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
        self.evaluator = evaluator
        self.parser = lark.Lark.open(
            'formulas.lark', start='formula', rel_to=__file__)

    def set_contents(self, input_str: Optional[str]):
        '''
        Set the contents of the cell.

        Arguments:
        - input_str: str - specifications to set new cell value

        '''

        try:
            # Remove leading and trailing whitespace
            inp = input_str.strip()
            self.contents = inp

            # Check if there is a leading single quote, set to STRING type
            if inp[0] == "'":
                self.value = inp[1:].strip() # Remove whitespace again

            # Check if there is a leading equal sign, set to FORMULA type
            # and evaluate
            elif inp[0] == "=":
                tree = self.parser.parse(inp)
                eval = self.evaluator.transform(tree).children[0]

                if isinstance(eval, CellError):
                    self.contents = self.get_string_from_error(
                        eval.get_type().value)
                    
                self.value = eval

            # Otherwise set to NUMBER type - works for now, will need to change
            # if we can have other cell types
            else:
                self.value = Decimal(inp) 

        except DecimalException as d:
            # needs to be checked
            self.value = inp
        
        except lark.exceptions.LarkError as l:
            self.value = CellError(CellErrorType.PARSE_ERROR, 
                            detail='Unable to parse entry', exception = l)
            self.contents = '#ERROR!'


    def empty(self):
        '''
        Empty the contents of a cell

        '''

        self.contents = None
        self.value = None


    def get_string_from_error(self, cell_error_type: CellErrorType) -> str:
        '''
        to-do
        '''
        
        return cell_error_dict[cell_error_type]
