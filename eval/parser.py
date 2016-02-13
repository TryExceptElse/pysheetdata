"""
functions for evaluating spreadsheet functions

primary function is parse, which the rest revolves around

evaluate should be called with the full string by a parent program

A note on exec:
    This uses the exec function repeatedly, and where possible, use of it
    should be minimized, but the intention of this is only meant to be run
    on trusted spreadsheets. Future development of this may focus on it being
    more secure, but the primary goal is simply to evaluate the most common
    functions, regardless the ability for code to be injected.

Another note:
    this whole thing could stand to be redone
"""

# import spreadsheet mirroring functions
import eval.functions as functions
import eval.translate as translate
import eval.storage as global_file  # historical reasons for name


__author__ = 'user0'


def evaluate(s, reference_dictionary=None):
    # if included, reference dictionary is a dictionary of relevant
    # cell references.
    # alternatively, if reference_dictionary is None, it is presumed
    # that it is not needed to replace references with values in the
    # formula. The reference_type arg, if none, defaults to 'sheet'

    if s[0] == '=':
        # get rid of the equals sign at the beginning of the formula
        s = s[1:]
        # send reference dictionary to storage
        global_file.formulas = reference_dictionary
    # I feel like I'm forgetting something else here
    return parse(s)


def parse(s, function=None):
    # returns evaluation of formula via recursive function;
    # before this function is run, dependencies should be
    # identified and evaluated

    replace = {}
    it = 0
    level = 0

    # replace references with cell values
    s = s.lower()
    # for formula in global_file.formulas:
    #     if formula in s:
    #         s = s.replace(formula, str(
    #             global_file.formulas[formula].return_value()))

    # replace values with python equivalents
    # ('^' with '**' for example)
    s = translate.spreadsheet_replace(s)

    # evaluate formula
    for char in s:
        if char == '(':
            level += 1
            if level == 1:
                parent_start = it
        if char == ')':
            level -= 1
            if level == 0:
                parent_close = it
                prefix = get_prefix(s, parent_start)
                body = s[parent_start + 1: parent_close]
                formula = '{}({})'.format(prefix, body)
                replace[formula] = str(parse(prefix, body))
                verbose('replacing {} with {}'.format(formula,
                                                      replace[formula]))

        it += 1

    # replace strings
    for entry in replace:

        s = s.replace(entry, replace[entry])

    # depending on the presence of a function, either simply evaluate,
    # or use a function from functions
    if function:
        # if function is in the replacement dictionary,
        # replace it with that entry
        if function in functions.function_replace:
            function = functions.function_replace[function]
        else:
            print('function %s was not in function dictionary') % function
        # function just stopped sounding like a word

        # insert the formula in a python-readable format
        body_strings = s.split(',')  # this is used below
        exec_string = '%s(body_strings)' % function
    else:
        # replace references with values and find result
        for reference in global_file.formulas:
            while reference in s:
                s = s.replace(reference, str(
                    global_file.formulas[reference].value()
                ))
        exec_string = s

    exec_string = eval_append(exec_string)

    verbose(exec_string)
    exec(exec_string)

    return global_file.returned


def get_prefix(formula_string, start):

    alpha = 'abcdefghijklmnopqrstuvwxyz'
    number = '.0123456789'

    prefix = ''

    string_position = start - 1

    while True:
        character = formula_string[string_position]

        if string_position >= 0:
            if character in alpha or character in number:
                prefix = character + prefix
            else:
                return prefix
        else:
            return prefix

        string_position -= 1


def eval_append(s):
    prefix = 'global_file.returned = '
    return prefix + s


def verbose(s):
    # if verbose setting, print s
    if global_file.verbose:
        print(s)
