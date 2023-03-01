'''
Cell

This module holds the basic functionality of dealing with an individual cell.

See the Workbook and Sheet modules for implementation.

Global Variables:
- RESTRICTED_VALUES (List[Decimal]) - values that a cell cannot be set to
    numerically

Classes:
- Cell

    Local Variables:
    - PARSER (Any) - Lark parser for formula evaluation

    Methods:
    - get_loc(object) -> str
    - get_contents(object) -> str
    - get_value(object) -> str
    - get_parser_and_evaluator(object) -> Tuple[Any, Evaluator]
    - set_contents_and_value(object, Optional[str], Optional[Any]) -> None
    - set_contents(object, Optional[str]) -> None
    - get_children(object) -> List[Tuple[str, str]]]
    - empty(object) -> None
    - set_circular_error(object) -> None
    - get_shifted_contents(object, Tuple[int, int]) -> str

'''


import re
from decimal import Decimal, DecimalException
from typing import Optional, List, Tuple, Any

import lark
from lark import Visitor, Tree

from .evaluator import Evaluator
from .cell_error import CellError, CellErrorType, CELL_ERRORS
from .utils import get_loc_from_coords, get_coords_from_loc


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
        '''
        Initialize a Cell Tree Visitor

        Arguments:
        - sheet: str - sheet to visit children in

        '''

        self.children = set()
        self.sheet = sheet

    def cell(self, tree: Tree) -> None:
        '''
        Get a cell from a Tree object

        Arguments:
        - tree: Tree - tree containing cell information

        '''

        if len(tree.children) == 2:
            cell_sheet = str(tree.children[0])
            if cell_sheet[0] == "'":
                cell_sheet = cell_sheet[1:-1]
            cell = str(tree.children[1]).replace('$', '')
        else:
            cell_sheet = self.sheet
            cell = str(tree.children[0]).replace('$','')
        self.children.add((cell_sheet, cell))


class Cell:
    '''
    A cell containing values of CellType and their string contents.

    This class represents an individual cell in a spreadsheet.
    Stores the string contents as well as the value
    Stores the type of the value as a CellType

    '''

    PARSER = lark.Lark.open('formulas.lark', start='formula',
                rel_to=__file__)

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

    def get_parser_and_evaluator(self) -> Tuple[Any, Evaluator]:
        '''
        Get the parser and evaluator

        Returns:
        - tuple with [parser, evaluator]

        '''

        return Cell.PARSER, self._evaluator

    def set_contents_and_value(self, contents: Optional[str],
            value: Optional[Any]) -> None:
        '''
        Set the contents and value fields of the cell

        Arguments:
        - contents: Optional[str] - contents to set, can be None
        - value: Optional[Any] - value to set, can be None

        '''

        self._contents = contents
        self._value = value

    def set_contents(self, input_str: Optional[str]) -> None:
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
                self.set_contents_and_value(contents, inp[1:])

            # Check if there is a leading equal sign, set to FORMULA type
            # and evaluate
            elif inp[0] == "=":
                parser, evaluator = self.get_parser_and_evaluator()
                tree = parser.parse(inp)
                visitor = _CellTreeVisitor(str(evaluator.get_working_sheet()))
                visitor.visit(tree)
                self._children = list(visitor.children)
                evaluator = evaluator.transform(tree).children[0]
                # Handle when referencing an empty cell only
                evaluator = Decimal('0') if evaluator is None else evaluator
                self.set_contents_and_value(contents, evaluator)

            elif inp.upper() in list(CELL_ERRORS.values()):
                e_type = [i[0] for i in list(CELL_ERRORS.items()) if \
                    i[-1]==inp.upper()]
                e_type = CellErrorType(e_type[0])
                self.set_contents_and_value(contents, CellError(e_type, '',
                    None))

            # Otherwise set to NUMBER type, if this fails throws
            # Decimal Exception
            else:
                temp_value = Decimal(inp)
                if temp_value in RESTRICTED_VALUES:
                    temp_value = inp
                self.set_contents_and_value(contents, temp_value)

        except DecimalException:
            self.set_contents_and_value(contents, inp)

        except lark.exceptions.LarkError as e:
            value = CellError(CellErrorType.PARSE_ERROR,
                            detail='Unable to parse entry', exception = e)
            self.set_contents_and_value(contents, value)

    def get_children(self) -> List:
        '''
        Gets the children of the cell

        Returns:
        - List of the children

        '''

        return self._children

    def empty(self) -> None:
        '''
        Empty the contents of a cell

        '''

        self.set_contents_and_value(None, None)

    def set_circular_error(self) -> None:
        '''
        Set the value of a cell to be a circular reference

        '''

        self._value = CellError(CellErrorType.CIRCULAR_REFERENCE,
                                "Cell is in a circular reference.")

    def get_shifted_contents(self, coord_shift: Tuple[int, int]) -> str:
        '''
        Shifts source cell contents to be target cell contents.  Handles
        absolute/mixed/relative referencing.

        Arguments:
        - source_contents: str - contents of source cell **TODO
        - coord_shift: Tuple[int, int] - diff between source & target cell

        '''

        source_contents = self.get_contents()
        source_value = self.get_value()

        # check if source cell contents are None or not formula, return
        if source_contents is None or source_contents[0] != "=":
            return source_contents

        # check if source cell value is a Parse Error, return
        if isinstance(source_value, CellError) and source_value.get_type() == \
            CellErrorType.PARSE_ERROR:
            return source_contents

        # remove strings for case we have '= ... & "sheet!A1"
        split = source_contents.split('\"')

        # Handler for regex substitution
        def subberoo(match):
            beg, col, row, end = match.groups()

            # Check for absolute col ref
            if col[0] == '$':
                c_shift=0
                col = col[1:]
                c_mark = '$'
            else:
                c_shift= coord_shift[0]
                c_mark = ''

            # Check for absolute row ref
            if row[0] == '$':
                r_shift = 0
                row = row[1:]
                r_mark = '$'
            else:
                r_shift = coord_shift[1]
                r_mark = ''

            x, y = get_coords_from_loc(col+row)
            x += c_shift
            y += r_shift

            try:
                loc = get_loc_from_coords((x,y))
            except ValueError:
                return f'{beg}#REF!{end}'

            split = re.split(r'(\d+)', loc.upper())
            return f'{beg}{c_mark}{split[0]}{r_mark}{split[1]}{end}'

        new_contents = ''
        for i, substring in enumerate(split):
            if i % 2 == 0:
                substring = re.sub(
                    r'([^ ])([\+\-\\\*\&])([^ ])', r'\1 \2 \3', substring)
                new_contents += re.sub(
                    r'([\ \-+\\\*=&!])(\$?[A-Za-z]+)(\$?[1-9][0-9]*)([^!]|$)',
                    subberoo, substring)
            else:
                new_contents += f'"{substring}"'

        return new_contents
