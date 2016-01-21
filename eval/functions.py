"""
functions that mirror those in excel, and operate (ideally) the same

imported by eval.parser
"""

import eval.storage as global_file


function_replace = {
    'if': 'functions.py_if',
    'sum': 'sum'  # I think this should work. if not.. add 'py_sum' fun
}


def py_if(args):

    s = 'if %s:\n' \
        '   global_file.function_return = %s\n' \
        'else:\n' \
        '   glob_file.function_return =  %s' % (args[0], args[1], args[2])

    exec(s)

    return global_file.function_return


"""
todo:
"""