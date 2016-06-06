"""
Stores spreadsheet error values

Note that these are not python errors, rather they simulate
Spreadsheet errors that can be caused by formulas, and are stored in
cells similar to floats or strings.
"""


class FormulaError:
    """
    Abstract Formula Error class
    """
    string = '#ERROR!'


class FormulaValueError(FormulaError):
    """
    Raised when improper operations take place, such as when a string
    is subtracted from a float.
    """
    string = '#VALUE!'


class FormulaNameError(FormulaError):
    """
    Raised when a function is called, but the name of the function is
    not recognized.
    There should be a dictionary of all known spreadsheet
    functions, and this value will be returned if a function is called
    that is not present within it.
    """
    string = '#NAME?'


class FormulaRefError(FormulaError):
    """
    Raised when a reference is made to a cell that cannot be found.
    It could be attempting to access a file or sheet that can not be
    found, or otherwise failing to return the value.
    """
    string = '#REF!'


class FormulaDivZeroError(FormulaError):
    """
    Someone divided by zero.
    """
    string = '#DIV/0!'
