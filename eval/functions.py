"""
functions that mirror those in excel, and operate (ideally) the same

imported by eval.parser

todo:
"""

import eval.storage as global_file


function_replace = {
    'if': 'functions.py_if',
    'sum': 'sum',  # I think this should work. if not.. add 'py_sum' fun
    'count': 'functions.py_count',
    'counta': 'functions.py_counta',
}


def py_if(args):

    s = 'if %s:\n' \
        '   global_file.function_return = %s\n' \
        'else:\n' \
        '   glob_file.function_return =  %s' % (args[0], args[1], args[2])

    exec(s)

    return global_file.function_return


def py_count(args):
    count = 0
    for arg in args:
        if arg in global_file.formulas:
            ref = global_file.formulas[arg]

            def count(cell_ob):
                if cell_ob.data_type == 'float' or cell_ob.data_type == 'int':
                    return 1
                else:
                    return 0

            if ref.__class__.__name__ == 'Cell':
                count += count(ref)
            else:
                for cell in ref.cells:
                    count += count(cell)
    return count


def py_counta(args):
    count = 0
    for arg in args:
        if arg in global_file.formulas:
            ref = global_file.formulas[arg]

            def count(cell_ob):
                if cell_ob.has_contents:
                    return 1
                else:
                    return 0

            if ref.__class__.__name__ == 'Cell':
                count += count(ref)
            else:
                for cell in ref.cells:
                    count += count(cell)
    return count