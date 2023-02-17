'''
Sheets

Package that contains all of the functionality for a spreadsheet engine.

Imported Modules:
- Workbook
- CellError
- CellErrorType

Attributes:
- version (str) - current version number

'''


from .workbook import Workbook
from .cell_error import CellError, CellErrorType

# pylint: disable=invalid-name

# We need to have a lowercase version field, and thus are disabling the
# check for snake casing.

version = '1.2'

# pylint: enable=invalid-name

# We enable the check for snake casing

__all__ = ['Workbook', 'CellError', 'CellErrorType']
