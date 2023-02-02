# Initialization script
from .workbook import Workbook
from .cell_error import CellError, CellErrorType

version = "1.0"

__all__ = [Workbook, CellError, CellErrorType]