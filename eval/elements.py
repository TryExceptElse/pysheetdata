"""
Handles elements making up an excel formula.
"""

from operator import add, sub, mul, truediv

from math import pow


class FormulaError(Exception):
    pass


class Element:
    """
    Uppermost abstract class for any item residing within a
    spreadsheet function.
    """

    def __init__(self, formula_string, operator_string):
        self.string = formula_string  # string of element
        self.operator_string = operator_string if operator_string is not \
            None else '+'

    def __str__(self):
        """
        Returns string representation of element
        :return: String value of element.
        """
        return str(self.string)

    def __add__(self, other):
        pass

    @property
    def operator(self):
        assert self.operator_string in OPERATORS, \
            "operator string %s not in operators: %s" % \
            (self.operator_string, ", ".join(OPERATORS.values()))
        return OPERATORS[self.operator_string]


class Bracket(Element):
    """
    Bracket class, handles portion of spreadsheet formula contained
    within brackets.
    """
    def __init__(self, formula_string, operator=None):
        super().__init__(formula_string, operator)
        self.sub_elements = []

    @property
    def value(self):
        """
        Gets value of this element
        :return: Element value. May be String, float, or int
        """
        return _evaluate(self.string)


class Formula(Bracket):
    """
    Class representing the overall formula of a spreadsheet cell.
    Operates similarly to an unqualified 'evaluation' bracket
    """

    def __init__(self, formula_string):
        super().__init__(formula_string, '+')

    @property
    def value(self):
        if self.string.startswith("="):
            return Bracket(self.string[1:]).value
        if self.string.startswith('of:='):
            return Bracket(self.string[4:]).value


class Function(Bracket):
    """
    Function with arguments in brackets.
    """
    def __init__(self, function_string, operator):
        super().__init__(function_string, operator)
        # function name string preceding brackets
        self.function_name_string, self.args = _parse_function_string(
            self.string)


class Reference(Element):
    """
    Acts like a basic element, but gets its value from a separate cell.
    May be either a single reference, or a range.
    """


def _evaluate(formula_string):
    """
    Evaluates a string to return its value
    :param formula_string:
    :return:
    """
    # get list of element objects composing the formula string
    elements = _get_elements_of_string(formula_string)
    # for order of operations level...
    for level in OPERATOR_LEVELS:
        # for each pair of elements...
        for i, element in enumerate(elements):
            next_el = elements[i + 1]
            # if second element's operator is in the current
            # operations level...
            if next_el.operator in level:
                # apply operator to the pair.
                elements[i] = next_el.operator(element, next_el)
                # remove next element since it has
                # been combined with the current index
                elements.pop(i + 1)
    assert len(elements) == 1, "after operations have all been" \
                               "applied, there should only be" \
                               "one element remaining. Actual" \
                               "elements list: %s" % elements
    return elements[0].value  # return last element's value


def _get_elements_of_string(formula_string):
    """List of elements composing the bracket
    :return: list of elements in order of occurrence.
    """
    elements = []
    i = 0
    operator_string = None
    while i < formula_string.length():
        char = formula_string[i]
        if char in OPERATORS:
            if operator_string is not None:
                raise FormulaError('Double operators present in'
                                   'formula: %s' % formula_string)
            operator_string = char
        elif char == '[':
            end_i = _find_paired_closing_bracket(formula_string, i)
            elements.append(Reference(formula_string[i + 1:end_i],
                                      operator_string))
            operator_string = None
            i = end_i
        elif char == '"':
            end_i = _find_end_quote_index(formula_string, i)
            elements.append(Element(formula_string[i + 1:end_i],
                                    operator_string))
            operator_string = None
            i = end_i
        elif char in ARABIC_NUMERALS or char == '.':
            end_i = _find_end_index_of_number(formula_string, i)
            elements.append(Element(formula_string[i:end_i + 1],
                                    operator_string))
            operator_string = None
            i = end_i
        elif char in LATIN_CHARACTERS:
            end_i = _find_function_end_index(formula_string, i)
            elements.append(Function(formula_string[i:end_i + 1],
                                     operator_string))
            operator_string = None
            i = end_i
        elif char == '(':
            end_i = _find_paired_closing_bracket(formula_string, i)
            elements.append(Bracket(formula_string[i + 1:end_i],
                                    operator_string))
            operator_string = None
            i = end_i
        i += 1
    return elements


def _find_end_index_of_number(formula_string, start_index):
    """
    Finds ending index of the number beginning at the passed index
    :param formula_string: formula string containing number for be found
    :param start_index: starting index int of the number
    :return: int end index of the number, may == start index.
    """
    for i, char in enumerate(formula_string[start_index + 1:]):
        # if index is not a digit, or decimal point...
        if char not in ARABIC_NUMERALS and char != '.':
            # consider the number string to have ended.
            return i + start_index
    # if the end of the string is hit:
    else:
        return formula_string.length - 1


def _parse_function_string(function_string):
    """
    Parses function's string to return the string of its name
    and args.
    :return: name string, tuple of arg strings
    """
    # find name string
    name_end_i = None
    name = None
    for i, char in enumerate(function_string):
        if char not in LATIN_CHARACTERS and char not in ARABIC_NUMERALS:
            name_end_i = i - 1
            name = function_string[:i]
            break
    assert isinstance(name_end_i, int)
    # find args string
    args_string = None
    for i, char in enumerate(function_string[name_end_i:]):
        if char is '(':
            args_end_i = _find_paired_closing_bracket(function_string,
                                                      i + name_end_i - 1)
            args_string = function_string[i:args_end_i + 1]
    # parse args string
    assert isinstance(name, str)
    assert isinstance(args_string, str)
    args = _separate_by_commas(args_string)
    return name, args


def _find_function_end_index(formula_string, start_index):
    """
    Finds end index of function and its arguments
    :param formula_string: string containing function
    :param start_index: index int at which function starts
    :return: index int at which function ends (closing bracket)
    """
    for i, char in enumerate(formula_string[start_index:]):
        if char in OPERATORS:
            raise FormulaError("could not parse formula %s" % formula_string)
        if char == '(':
            return _find_paired_closing_bracket(formula_string, 
                                                start_index + i)


def _find_end_quote_index(string, quote_start_index):
    """
    Returns the ending quote mark index for a passed starting
    quote index
    :param string: string containing quote.
    :param quote_start_index: int index of starting quote mark.
    :return: int index of ending quote mark.
     """
    assert string[quote_start_index] in QUOTES
    start_quote_char = string[quote_start_index]
    for i, char in enumerate(string[quote_start_index:]):
        if char == start_quote_char:
            return i + quote_start_index


def _find_paired_closing_bracket(string, open_bracket_index):
    """
    Finds index at which bracket opened at open_bracket_index
    is closed
    :param string: string containing brackets
    :param open_bracket_index: index at which bracket opens
    :return: int index at which bracket closes
    """
    brackets = {
        '<': '>',
        '[': ']',
        '{': '}',
        '(': ')',
    }
    start_bracket = string[open_bracket_index]
    assert start_bracket in brackets, 'character at index %s in "%s"; "%s" ' \
                                      "is not a bracket ('%s')" \
                                      % (open_bracket_index, string,
                                         string[open_bracket_index],
                                         "', '".join(brackets.keys()))
    close_bracket = brackets[start_bracket]
    bracket_depth = 0
    i = open_bracket_index + 1  # string index
    while i < len(string):
        char = string[i]
        if char in QUOTES:
            i = _find_end_quote_index(string, i)
        elif char == start_bracket:
            bracket_depth += 1
        elif char == close_bracket:
            if bracket_depth == 0:
                return i
            else:
                bracket_depth -= 1
        i += 1


def _separate_by_commas(args_string):
    """
    Separates the passed formula by commas.
    Commas appearing in quotes, or within parenthesis will not be used
    :param args_string: complete args string.
    :return: tuple of strings, one for each arg.
    """
    # first check if args string is empty,
    # if so, return an empty list.
    if not args_string:  # empty Strings are false
        return ()
    # otherwise, look through args for a comma and separate the
    # values there.
    # This should not consider commas that are contained in quotes
    # or inside brackets.
    args = []
    i = 0
    last_comma_index = 0
    while i < len(args_string):
        char = args_string[i]
        if char == '"':
            i = _find_end_quote_index(args_string, i)
        elif char == '(':
            i = _find_paired_closing_bracket(args_string, i)
        elif i == ',':
            args.append(args_string[last_comma_index, i])
        i += 1
    else:
        args.append(args_string[last_comma_index, len(args_string)])
    return tuple(args)


OPERATORS = {
    '^': pow,
    '*': mul,
    '/': truediv,
    '+': add,
    '-': sub,
}
OPERATOR_LEVELS = (
    ('^',),  # tuple of one
    ('*', '/'),
    ('+', '-')
)
QUOTES = "'", '"'
ARABIC_NUMERALS = '0123456789'
LATIN_CHARACTERS = 'abcdefghijklmnopqrstuvwxyz' \
                   'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
