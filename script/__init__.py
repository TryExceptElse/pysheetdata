"""
This stores variables and related things for python scripts that are
run in a cell

Included / to be included:
    value  # the script cell's value
    formula  # the script cell's formula
    text  # the script cell's text - will start as the script text

    this_row  # list of the cells in the script cell's row
    this_column  # list of the cells in the script cell's column

"""


class _CellReferencer:
    """
    returns cells from references in script
    """

    def __init__(self, lookup_f):
        self.lookup_f = lookup_f

    def __getitem__(self, item):
        return self.lookup_f(item)

cell = None
cells = None
