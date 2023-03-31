# cs130
This repository contains all of the course work for CS 130.

All work seen is completed by Dallas Taylor (dtaylor@caltech.edu), Kyle McGraw (kmcgraw@caltech.edu), and Benjamin Juarez (bjuarez@caltech.edu).


**Caltech Honor Code**: No member of the Caltech community shall take unfair advantage of any other member of the Caltech community.

Per Caltech Honor Code, do not view/download these files if you plan to take (or are currently taking) this course

## Spreadsheet Engine

This course was composed of 5 two-week projects (similar to a Sprint) that culminated into one spreadsheet engine similar to that of Excel or Google Sheets.  Here, we will describe the usage and testing of the engine.

### Terminal Functionality

Before initializing a workbook, all necessary dependencies must be installed.  In order to do this, navigate to the repository's directory and run the following commands (note that it is recommended you use a virtual environment):

`cd cs130`
`make setup`

All necessary libraries will now be installed.  For specific information on what is being used, see `requirements.txt`.

Now, enter into a python environment and import the Sheets module:

`import sheets`

Example usage of the Sheets module can be seen in `examples.py`.

### Testing

If any edits are made to the Sheets module, basic functionality can be tested from the terminal.  All of the make targets are specified in the `Makefile`, but here are some of the most useful targets:

- to run all tests: `make test`
- to test just the Evaluator: `make test-evaluator`
- to test a performance test titled `x`: `make test-performance-x`
- to lint the Sheets module: `make lint-all`
- to lint all functionality tests: `make lint-test-all`
- to lint a specific file titled `x`: `make lint-x`

Note, to adjust the linter to accept different style choices, you can edit `.pylintrc` in the main `cs130` directory.

### Notes

Supported functions for this spreadsheet engine include the following:
- AND
- OR
- NOT
- XOR
- EXACT
- IF
- IFERROR
- CHOOSE
- ISBLANK
- ISERROR
- VERSION
- INDIRECT
- MIN
- MAX
- SUM
- AVERAGE
- HLOOKUP
- VLOOKUP

Cell Error types include the following, with this order of importance:
1. Parse Error (#ERROR!)
2. Circular Reference (#CIRCREF!)
3. Bad Reference (#BADREF!)
4. Bad Name (#NAME?)
5. Type Error (#VALUE!)
6. Divide By Zero Error (#DIV/0!)
