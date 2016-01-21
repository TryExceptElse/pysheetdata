"""
dictionary and functions for making excel format formulas readable for
python

imported by eval.parser
"""

__author__ = 'user3'


def spreadsheet_replace(s):
    for entry in replace_dict:
        while entry in s:
            s = s.replace(entry, replace_dict[entry])
    return s

replace_dict = {
    '^': '**',
    'pi()': '3.14159265358979',  # same value as used by spreadsheets
}
