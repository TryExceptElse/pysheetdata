"""

=======================================================================
WARNING: PYSHEETDATA IS NOT SECURE AGAINST MALICIOUSLY CONSTRUCTED DATA
=======================================================================
it is intended to be used only with trusted data.

note on file names:
    file names including ' : [ or ] are likely to mess things up, as
    they tend to do to other spreadsheet programs as well

given a ods file -
    extract
    parse contents file
        for each instance of cell name:
            add cell to list
                get contents_string
                from string:
                    get value
                    get raw_formula
                    get formula from raw_formula
                        (use brackets to find reference)  - (so much easier)
                    get xy ref
                    get a1 ref

                    each cell needs (verbatim):

                    return_value()
                    return_raw_formula
    return library or file

scripting variables:
    cells[spreadsheet reference] - gets a cell object, or matrix of objects
    sheet.rows[y]
    sheet.columns
"""

import zipfile
import xml.etree.ElementTree

from copy import deepcopy

import eval.storage as storage
import eval.parser as parser
import script.script as script


class SheetDataError(Exception):
    pass


class Cell:
    # in some sheets there are going to be a LOT of these, should be
    # coded as such.
    # 78,000 plus in some sheets that I'll be using this for,
    # and I'm sure others will have far more
    def __init__(self, cell_data, position, sheet, mode='standard'):
        self.mode = mode
        # changing mode allows different amounts of data to be stored
        # for each cell. speed vs memory usage
        with cell_data as self.cell_data:
            # cell data is ditched after init
            if self.mode == 'standard' or self.mode == 'formatted':
                self.position = position  # (x, y)
                self.sheet = sheet
                if self.cell_data:
                    self.has_contents = True
                else:
                    self.has_contents = False
                if self.has_contents:
                    self.data_type = self.cell_data['office:value-type']
                    self.text = self.cell_data['text']
                    if 'office:value' in self.cell_data:
                        self.cached_value = self.cell_data['office:value']
                        self.value = deepcopy(self.cached_value)
                    else:
                        self.cached_value = None
                        self.value = None
                    if 'table:formula' in self.cell_data:
                        self.raw_formula = self.cell_data['table:formula']
                    else:
                        self.raw_formula = None
                    if self.text[:3] == 'py=':
                        self.is_script = True
                    else:
                        self.is_script = False
                # dictionary of dependencies with the cell's reference
                # in self.raw_formula as key
                self.dependencies = self.find_dependencies()
                self.a1 = a1_from_xy(self.position)
            # todo: test if cell data is still present after init,
            # if so use alternate means to save space
            if self.mode == 'formatted':
                pass  # formatting stuff to load if needed.

    def return_value(self):
        if self.has_contents:
            if self.value is None:
                self.evaluate()
            return self.value
        else:
            return None

    def return_script(self):
        # considering simply adding 'self.script' to cell
        if self.is_script:
            return self.text[3:]
        else:
            return None

    def find_dependencies(self):
        # returns dictionary of referenced cells with reference as key
        # if reference is a range (like 'A1:B2') returns matrix
        if not self.has_contents or (self.raw_formula is None and
                                     self.return_script() is None):
            return {}
        else:
            dependencies = {}
            # look in formula first
            if self.raw_formula is not None:
                # find start of a reference
                is_quote = False
                start = None
                for x in range(0, len(self.raw_formula)):
                    if self.raw_formula[x] == "'":
                        is_quote = not is_quote
                    elif self.raw_formula[x] == '[' and not is_quote:
                        # reference found!
                        start = x
                    elif self.raw_formula[x] == ']' and not is_quote:
                        # set reference string
                        reference = self.raw_formula[start + 1: x - 1]
                        # set default values, will be changed in a
                        # moment if needed
                        dependencies['[' + reference + ']'] = \
                            self.cell_from_reference(reference)
            if self.return_script() is not None:
                start = None
                script_s = self.return_script()
                for x in range(0, len(script_s)):
                    # try:
                        if script_s[x: x + 5] == 'cells[':
                            start = x
                        elif script_s[x] == ']' and \
                                        start is not None:
                            reference = script_s[start + 6:x - 1]
                            start = None
                            dependencies['cells[' +reference + ']'] = \
                                self.cell_from_reference(reference)
                    # except:  # todo: find out what error happens
                        # here and specify that.
                        # pass
            return dependencies

    def cell_from_reference(self, reference):
        reference_type = 'cell'  # default reference
        range_start = None
        range_end = None
        # now that there's a reference, iterate through it
        # to see if there's a ':' outside of quotes
        # denoting a range of cells
        if ':' in reference:
            index = find_unquoted(':', reference)
            if index is not None:
                reference_type = 'range'
                range_start = reference[:index - 1]
                range_end = reference[index + 1:]
        if reference_type == 'cell':
            reference_parts = break_apart_reference(reference)
            return self.sheet.book.library.return_cell(reference_parts)
        elif reference_type == 'range':
            start_parts = break_apart_reference(range_start)
            start_parts = self.complete_reference_parts(
                    start_parts)
            end_parts = break_apart_reference(range_end)
            end_parts = self.complete_reference_parts(
                    end_parts)
            return CellRange(
                self.sheet.book.library.return_cell(
                        start_parts),
                self.sheet.book.library.return_cell(
                        end_parts)
            )

    def complete_reference_parts(self, parts):
        sheet_default = self.sheet.name
        book_default = self.sheet.book.file_name
        length = len(parts)
        if length == 1:
            return [book_default, sheet_default] + parts
        elif length == 2:
            return [book_default] + parts
        elif length == 3:
            return parts
        else:
            raise SheetDataError('reference parts outside 1-3 range')

    def evaluate(self, recursive=True, scripts=True):
        if self.has_contents:
            if recursive:
                for cell in self.dependencies:
                    self.dependencies[cell].evaluate()
            if scripts and self.return_script():
                self.run_script(self.dependencies)
            elif self.raw_formula:
                evaluation = parser.evaluate(self.raw_formula,
                                             self.dependencies)
                # if returned is value, put it in self.value
                # if that doesn't work, put it in self.string
                try:
                    self.value = float(evaluation)
                except TypeError:
                    try:
                        self.text = str(evaluation)
                    except TypeError:
                        pass
        else:
            return None

    def run_script(self, dependencies):
        # dependencies is list of cells that are used
        if self.has_contents and self.is_script:
            # set script vars that can be called by the script
            script.value = None
            script.formula = None
            script.text = self.text
            script.cells = dependencies
            script_string = self.text[3:len(self.text)]
            exec(script_string)
        else:
            return ''


class CellRange:
    def __init__(self, start_cell, end_cell):
        sheet_list = start_cell.sheet.book.sheet_list

        # create 3d matrix of cells inside the range (sheet, row, cell)
        indexes = {
            'sheets': [sheet_list.index(start_cell.sheet.name),
                       sheet_list.index(end_cell.sheet.name)],
            'rows': [start_cell.position[1], end_cell.position[1]],
            'cells': [start_cell.position[0], end_cell.position[0]]
        }
        for pair in indexes:
            indexes[pair].sort()
        self.matrix = []
        for sheet in start_cell.sheet.book.sheet_list[indexes['sheets'][0]:
                                                      indexes['sheets'][1]]:
            sheet_matrix = []
            self.matrix.append(sheet_matrix)
            for row in sheet.matrix[indexes['rows'][0]:indexes['rows'][1]]:
                row_matrix = []
                sheet_matrix.append(row_matrix)
                for cell in row[indexes['cells'][0]:indexes['cells'][1]]:
                    row_matrix.append(cell)

    def return_value(self):
        return sum([sum([sum([cell.return_value() for cell in row])
                   for row in sheet]) for sheet in self.matrix])

    def find_dependencies(self):
        d = {}
        for sheet in self.matrix:
            for row in sheet:
                for cell in row:
                    d.update(cell.find_dependencies())
        return d

    def evaluate(self):  # evaluates each cell in matrix
        [[[cell.evaluate() for cell in row]
          for row in sheet]
         for sheet in self.matrix]


class Book:
    def __init__(self, library, file_name):
        self.file_name = file_name
        self.library = library
        self.sheets = {}
        self.sheet_list = []  # list of sheets, in order of file


class Sheet:
    def __init__(self, book, name, matrix):
        self.name = name
        self.book = book
        self.matrix = matrix  # matrix of cells


class Library:
    def __init__(self, list_of_addresses, recursive_loading=False):
        self.list = list_of_addresses
        self.files = [File(address, self) for address in self.list]
        self.books = {}
        self.recursive = recursive_loading

        [file.load_cells() for file in self.files]

    def load_books(self):
        for book_name in self.list:
            self.load_book(book_name)

    def load_book(self, book_name):
        self.files[book_name] = File(book_name)
        self.files[book_name].load_cells()

    def return_cell(self, *strings):
        # takes either xy or a1 cell ref
        # book, sheet, (row, cell) or (cell a1 ref)
        if strings[0] not in self.books:
            try:
                if self.recursive:
                    self.load_book(strings[0])
                else:
                    raise SheetDataError('cell referenced cell not in loaded'
                                         'book')
            except:
                raise SheetDataError('could not find referenced book')
        return self.books[strings[0]].return_cell(strings[1:])


class File:
    def __init__(self, file_address, library=None):
        self.file = file_address
        self.body = None
        self.library = library

        if self.file[:7] == 'file://':
            self.file = self.file[7:]

    def load_cells(self, only_sheet=None):
        # load cell instances into
        #
        # dictionary (library)
        # of dictionaries (book)
        # of lists (sheet)
        # of lists (row)
        # of objects (cell)

        with zipfile.ZipFile(self.file).open('content.xml') as content:
            data = xml.etree.ElementTree.fromstring(content)
            book_dict = {}
            book = Book(self, self.file)
            self.library.books[self.file] = book
            # find each child which is a sheet ('table:table')
            # if looking for a specific sheet, make sure it's that one
            for sheet in [child for child in data if
                          child.tag == 'table:table' and
                          (child.find('name') == only_sheet or
                           only_sheet is None)]:
                sheet_name = sheet.find('name')
                book.sheet_list.append(sheet_name)
                rows = []
                book_dict[sheet_name] = rows
                sheet_instance = Sheet(book, sheet_name, rows)
                book.sheets[sheet_name] = sheet_instance
                y = 0
                for row in [child for child in sheet if
                            child.tag == 'table:table-row']:
                    cells = []
                    rows.append(cells)
                    x = 0
                    for cell in [child for child in row if
                                 child.tag == 'table:table-cell']:
                        cell_data = {}
                        for key, value in cell.items():
                            cell_data[key] = value
                        cell_data['text'] = cell.find('text')
                        # this may raise an error, if so, try-except it
                        position = x, y
                        cells.append(Cell(cell_data, position, sheet_instance))


def a1_from_xy(pos):
    # returns excel-type cell given X, Y position
    # A corresponds to 0 in x-axis
    # 1 corresponds to 0 in y-axis

    x_ref = x_to_letter(pos[0] + 1)
    y_ref = str(pos[1] + 1)

    return x_ref + y_ref


def x_to_letter(x):
    # takes in a number and returns it as letter(s)
    # ex.: 1 = 'a', 26 = 'aa'

    l = []
    s = ''

    # while x > 26 divides it by 26 and adds the remainder to l
    while x > 26:
        l.append(x % 26)
        x //= 26

    l.append(x % 26)

    # since the above worked from right
    # (smallest to largest place digit)
    # this reverses the order
    l = l[::-1]

    # convert each i in l into the corresponding letter
    for i in l:
        s += storage.ENGLISH_ALPHABET[i - 1]

    return s


def break_apart_reference(s):
    # takes a reference string and breaks it down into a list of
    # file name, sheet name, cell name
    parts = []

    # iterate backwards through string looking for '.' then '#$'
    dot = find_unquoted('.', s, True)
    if dot is None:
        parts.append(s)
        return parts
    elif dot is 0:
        parts.append(s[dot + 1:])
        return parts
    else:
        parts.append(s[dot + 1:])
    # now try to find the #$
    hd = find_unquoted('#$', s, True)
    if hd is None:
        parts.append(s[:dot - 1])
    else:
        parts.append(s[hd + 1: dot - 1])
        parts.append(s[:hd - 1])
    # reverse order to go book, sheet, cell rather than the reverse
    return list(reversed(parts))


def find_unquoted(target, string, back=False, list_mode=False):
    # finds unquoted target in string and returns the position of the
    # first character
    motion = 1
    index = 0
    target_length = len(target)
    if back:
        motion *= -1
        index = len(string) - 1
    if list_mode:
        matches = []
        while 0 <= index < len(string):
            if string[index] == "'":
                index += motion
                while string[index] != "'":
                    index += motion
            elif string[index: string + target_length - 1] == target:
                matches.append(index)
            index += motion
            if not 0 <= index < len(string) - target_length:
                return
    else:
        while string[index: index + target_length - 1] != target:
            if string[index] == "'":
                index += motion
                while string[index] != "'":
                    index += motion
            index += motion
            if not 0 <= index < len(string) - target_length:
                return
        return index

