from decimal import Decimal
import io, json

from sheets import *

# Initialize a Workbook
wb1 = Workbook()

# Create new sheets and delete sheets.
# Optionally save the sheet name and index.
# Optionally give a sheet name.
wb1.new_sheet('Sheet2')
wb1.del_sheet('Sheet2')
sheet_idx, sheet_name = wb1.new_sheet()

# Set a cells contents to numerical, string, and boolean literals
# Cell location labeling (A1) is consistent with that of Excel - 
# the letter refers to the column and the number refers to the row
wb1.set_cell_contents(sheet_name, 'A1', '3.1')
wb1.set_cell_contents(sheet_name, 'A2', 'True')
wb1.set_cell_contents(sheet_name, 'A3', 'string')
# This will be set to the numerical value 3.1
wb1.set_cell_contents(sheet_name, 'D5', '3.1')
# While this will be set to the string "3.1"
wb1.set_cell_contents(sheet_name, 'D6', '\'3.1')
# Cell contents can be reset
wb1.set_cell_contents(sheet_name, 'A1', 'False')
# Cells can also be set to empty
wb1.set_cell_contents(sheet_name, 'D6', '')
wb1.set_cell_contents(sheet_name, 'D7', '')

# Set a cells contents to a formula
# Formulas are marked by the = sign at the beginning of the input string
# Formulas can include all of the following examples
# All of the same literals as before:
wb1.set_cell_contents(sheet_name, 'A1', '=1')
wb1.set_cell_contents(sheet_name, 'A1', '="a string"')
wb1.set_cell_contents(sheet_name, 'A1', '=True')
# Addition and subtraction
wb1.set_cell_contents(sheet_name, 'A1', '=1+1')
wb1.set_cell_contents(sheet_name, 'A1', '=1-1')
# Multiplication and division
wb1.set_cell_contents(sheet_name, 'A1', '=2*2')
wb1.set_cell_contents(sheet_name, 'A1', '=2/2')
# String concatenation
wb1.set_cell_contents(sheet_name, 'A1', '="a string"&"another string"')
# Unary operations and parenthesis
wb1.set_cell_contents(sheet_name, 'A1', '=-(20+3)')
# Boolean comparisons
wb1.set_cell_contents(sheet_name, 'A1', '=2<3')
wb1.set_cell_contents(sheet_name, 'A1', '=True=False')
# Cell references
wb1.set_cell_contents(sheet_name, 'A1', '=A1+b2')
wb1.set_cell_contents(sheet_name, 'A1', '=A99 > B3')
# Function calls - the list of included functions is available in the README
wb1.set_cell_contents(sheet_name, 'A1', '=AND(True, True)')
wb1.set_cell_contents(sheet_name, 'A1', '=SUM(A2, A5)')
# Cell ranges are still not entirely functional, but the idea is that you
# can use them like so
wb1.set_cell_contents(sheet_name, 'A1', '=SUM(A1:A8)')

# Access cell values and contents
# Contents will be a string representation of whatever value is passed when
# setting the cell's contents
# Values will be any type, representing the evaluated result of the given
# contents, numeric values will be a Decimal object
wb1.set_cell_contents(sheet_name, 'A1', '=2')
contents = wb1.get_cell_contents(sheet_name, 'A1')
value = wb1.get_cell_value(sheet_name, 'A1')
assert contents == '=2'
assert value == Decimal('2')

# Setting a cell contents to anything that causes internal errors will set the
# cell to be a CellError with a descriptive type.  All error types are described
# in the README.
wb1.set_cell_contents(sheet_name, 'A1', '="no closing quotes')
contents = wb1.get_cell_contents(sheet_name, 'A1')
value = wb1.get_cell_value(sheet_name, 'A1')
assert contents == '="no closing quotes'
assert isinstance(value, CellError)
assert value.get_type() == CellErrorType.PARSE_ERROR

wb1.set_cell_contents(sheet_name, 'A1', '=1/0')
contents = wb1.get_cell_contents(sheet_name, 'A1')
value = wb1.get_cell_value(sheet_name, 'A1')
assert contents == '=1/0'
assert isinstance(value, CellError)
assert value.get_type() == CellErrorType.DIVIDE_BY_ZERO

# On a larger scale, you are able to perform the following operations:

# List all of the sheets in the workbook
list_of_sheets = wb1.list_sheets()
assert list_of_sheets == [sheet_name]

# Get the extent of a sheet (# of columns, # of rows)
extent = wb1.get_sheet_extent(sheet_name)
assert extent == (4, 5)

# Rename a sheet
wb1.rename_sheet(sheet_name, 'NewSheet')
sheet_name = 'NewSheet'

# Move a sheet, such that its index amongst other sheets in the workbook is
# altered. Indexing is preserved.
_, sheet_name2 = wb1.new_sheet()
list_of_sheets = wb1.list_sheets()
assert list_of_sheets == [sheet_name, sheet_name2]
wb1.move_sheet(sheet_name2, 0)
list_of_sheets = wb1.list_sheets()
assert list_of_sheets == [sheet_name2, sheet_name]

# Copy a sheet, with all of its cells
wb1.copy_sheet(sheet_name)
list_of_sheets = wb1.list_sheets()
assert list_of_sheets == [sheet_name2, sheet_name, f'{sheet_name}_1']

# Move a group of cells, setting the old locations to be empty
wb1.set_cell_contents(sheet_name, 'A1', '=1')
value1 = wb1.get_cell_value(sheet_name, 'A1')
value2 = wb1.get_cell_value(sheet_name, 'A2')
assert (value1, value2) == (Decimal(1), True)
wb1.move_cells(sheet_name, 'A1', 'A1', 'A2')
value1 = wb1.get_cell_value(sheet_name, 'A1')
value2 = wb1.get_cell_value(sheet_name, 'A2')
assert (value1, value2) == (None, Decimal(1))

# Copy a group of cells, leaving the old locations as they are
wb1.copy_cells(sheet_name, 'A1', 'A2', 'B2')
value1 = wb1.get_cell_value(sheet_name, 'A1')
value2 = wb1.get_cell_value(sheet_name, 'A2')
assert (value1, value2) == (None, Decimal(1))
value1 = wb1.get_cell_value(sheet_name, 'B2')
value2 = wb1.get_cell_value(sheet_name, 'B3')
assert (value1, value2) == (None, Decimal(1))

# Sort a region of cells
wb1.set_cell_contents(sheet_name, 'A1', '2')
wb1.sort_region(sheet_name, 'A1', 'A2', [1])
value1 = wb1.get_cell_value(sheet_name, 'A1')
value2 = wb1.get_cell_value(sheet_name, 'A2')
assert (value1, value2) == (Decimal(1), Decimal(2))

# You can also save your current workbook as a TextIO in JSON format
wb1.del_sheet(sheet_name)
wb1.del_sheet(sheet_name2)
wb1.del_sheet(f'{sheet_name}_1')
wb1.new_sheet('S1')
wb1.set_cell_contents('S1', 'A1', '=1')

with io.StringIO('') as fp:
    wb1.save_workbook(fp)
    fp.seek(0)
    json_act = json.load(fp)
    json_exp = {
        'sheets':[
            {
                'name':'S1',
                'cell-contents':{
                    'A1': '=1'
                }
            }
        ]
    }
    assert json_act == json_exp

# And you can load in a TextIO to load another workbook
with io.StringIO('') as fp:
    json.dump({
        'sheets':[
            {
                'name':'S1',
                'cell-contents':{
                    'A1': '=1'
                }
            }
        ]
    }, fp)
    fp.seek(0)
    wb2 = Workbook.load_workbook(fp)
    contents = wb2.get_cell_contents('S1', 'A1')
    value = wb2.get_cell_value('S1', 'A1')
    assert contents == '=1'
    assert value == Decimal('1')
