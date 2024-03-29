'''
Cell Error

This module holds the basic functionality of an individual cell error.

See the Cell module for implementation.

Global Variables:
- CELL_ERRORS (Dict[int, str]) - dictionary mapping an int to the string
    literal for that error type

Classes:
- CellErrorType (Enum)

   enum.Enum class that enumerates the 6 error types.
   For more information, see the CellErrorType class

- CellError

    Methods:
    - get_type(object) -> CellErrorType
    - get_detail(object) -> str
    - get_exception(object) -> Optional[Exception]

'''

import enum
from typing import Optional


CELL_ERRORS = {
    1: '#ERROR!',
    2: '#CIRCREF!',
    3: '#REF!',
    4: '#NAME?',
    5: '#VALUE!',
    6: '#DIV/0!'
}


class CellErrorType(enum.Enum):
    '''
    This enum specifies the kinds of errors that spreadsheet cells can hold.
    '''

    # A formula doesn't parse successfully ("#ERROR!")
    PARSE_ERROR = 1

    # A cell is part of a circular reference ("#CIRCREF!")
    CIRCULAR_REFERENCE = 2

    # A cell-reference is invalid in some way ("#REF!")
    BAD_REFERENCE = 3

    # Unrecognized function name ("#NAME?")
    BAD_NAME = 4

    # A value of the wrong type was encountered during evaluation ("#VALUE!")
    TYPE_ERROR = 5

    # A divide-by-zero was encountered during evaluation ("#DIV/0!")
    DIVIDE_BY_ZERO = 6


class CellError:
    '''
    This class represents an error value from user input, cell parsing, or
    evaluation.

    '''

    def __init__(self, error_type: CellErrorType, detail: str,
                 exception: Optional[Exception] = None):
        self._error_type = error_type
        self._detail = detail
        self._exception = exception


    def get_type(self) -> CellErrorType:
        ''' 
        Get the category of the cell error

        '''

        return self._error_type


    def get_detail(self) -> str:
        '''
        More detail about the cell error
        
        '''

        return self._detail


    def get_exception(self) -> Optional[Exception]:
        '''
        If the cell error was generated from a raised exception, this is the
        exception that was raised.  Otherwise this will be None.

        '''

        return self._exception


    def __str__(self) -> str:
        return f'ERROR[{self._error_type}, "{self._detail}"]'


    def __repr__(self) -> str:
        return self.__str__()
