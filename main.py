"""

=======================================================================
WARNING: PYSHEETDATA IS NOT SECURE AGAINST MALICIOUSLY CONSTRUCTED DATA
=======================================================================
it is intended to be used only with trusted data.

note on file names:
    file names including ' : [ or ] are likely to mess things up, as
    they tend to do to other spreadsheet programs as well (due to
    the way file names and references are annotated in the xml file)

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
import lxml.etree as etree

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
        self._cell_element = cell_data
        self._position = position
        self._cached_position = position
        self.sheet = sheet
        # stores new values, without applying them to _cell_element
        self._new_attrib = {}
        self._new_text = None

        # may be added in the future to speed up program
        # self._dependencies
        # self._dependants

    @property
    def library(self):
        try:
            return self.sheet.book.library
        except AttributeError:
            return None

    @property
    def file(self):
        return self.sheet.book.file

    @property
    def book(self):
        return self.sheet.book

    @property
    def map(self):
        return self.file.map

    @property
    def has_contents(self):
        if self._cell_element:
            return True
        else:
            return False

    @property
    def data_type(self):
        return self.get('office:value-type')

    @data_type.setter
    def data_type(self, value):
        self.set('office:value-type', value)

    @property
    def value(self):
        # at the moment, reevaluates all cells in dependency tree
        # in the future, once dependants and change flags are
        # instituted, will only reevaluate if a cell in the tree has
        # changed
        self.evaluate()
        return self.get('office:value')

    @value.setter
    def value(self, value):
        self.set('office:value', value)

    @property
    def cached_value(self):
        # always returns original value from the xml file
        return self.get('office:value', True)

    @property
    def text(self):
        # text getter/setter is different from other attributes because
        # it is stored separately in the xml file
        return self._cell_element.text

    @text.setter
    def text(self, string):
        # not using self.set because text is stored separately
        if string != self.text:
            self._new_text = string

    @property
    def cached_text(self):
        return self._cell_element.text

    @property
    # this is the formula as modified to appear as it does as typed
    # by a user in the spreadsheet program
    def formula(self):
        return self.get('table:formula')

    @property
    # this is the formula as stored in the xml file, with formatting
    # left in
    def raw_formula(self):
        return self.get('table:formula')

    @property
    def is_script(self):
        if self.text.startswith(PYSCRIPT_FLAG):
            return True
        else:
            return False

    @property
    def script(self):
        if self.is_script:
            return self.text[len(PYSCRIPT_FLAG):]

    @script.setter
    def script(self, script_string):
        self.text = PYSCRIPT_FLAG + script_string

    @property
    def a1(self):
        return a1_from_xy(self.position)

    @a1.setter
    def a1(self, a1_string):
        self.position = xy_from_a1(a1_string)

    @property
    def position(self):
        return self._position

    @position.setter
    def position(self, tuple_or_list):
        self.position = (tuple_or_list[0], tuple_or_list[1])

    @property
    def cached_position(self):
        return self._cached_position

    @property
    # returns the dependencies of self.
    # may just have them be stored as standard
    # I -don't think- this needs to be sent over to _new_attrib, but
    # if that proves to be the case, this will need to be updated.
    def dependencies(self):
        return self.find_dependencies()

    def get(self, string, cached=False):
        # get method called by property getters that use the attrib
        # dictionary
        # check if attribute has a new value in self._new_attrib
        # before getting it from the element tree attrib dictionary
        string = self.ns(string)
        if string in self._new_attrib and not cached:
            return self._new_attrib[string]
        else:
            return self._cell_element.get(string)

    def set(self, key, entry):
        # set method called by properties using the attrib dict
        # if new val is different from original, put in new_attrib
        # dict, rather than editing loaded element tree. This allows
        # changes to be reverted
        key = self.ns(key)
        if entry != self.get(key):
            self._new_attrib[key] = entry

    def return_value(self):
        # candidate for deletion once code has been reconstituted to not
        # need these non-property getters/setters
        return self.value

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
                # fixing text prop will fix the above error warning
                # (if no warning, I forgot to remove this after fixing)
                for x in range(0, len(script_s)):
                    # try:
                        if script_s[x: x + 5] == 'cells[':
                            start = x
                        elif script_s[x] == ']' and \
                                start is not None:
                            reference = script_s[start + 6:x - 1]
                            start = None
                            dependencies['cells[' + reference + ']'] = \
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
        if 0 < length <= 3:
            additions = [book_default, sheet_default]
            return additions[0: 3 - length] + parts
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

    def ns(self, string):
        return self.map.ns(string)


class CellRange:
    # range of cells intended for evaluator - acts similar to cell
    # in that it returns (total) value, dictionary of all dependencies,
    # etc
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
    # gets passed relevant dictionary of elements by file.load()
    # should not load cells until that function is called
    # then pass on relevant sub-dictionary to sheets that are created
    # so on
    def __init__(self, library, file, element_tree):
        self.file = file
        self._element_tree = element_tree
        self.library = library
        self.sheets = {}
        self.sheet_list = []  # list of sheets, in order of file

    @property
    def file_name(self):
        return self.file.file_id

    def return_cell(self, *strings):
        if strings[0] in self.sheets:
            pass
        elif strings[0] in self.sheet_list:
            if self.library.recursive:
                self.library.files[self.file_name].load_cells(strings[0])
            else:
                raise SheetDataError('cell referenced a cell not in loaded'
                                     'sheet \'%s\'' % strings[0])
        else:
            raise SheetDataError('could not find referenced sheet \'' +
                                 str(strings[0]) + '\'')
        return self.sheets[strings[0]].return_cell(strings[1:])

    def __getitem__(self, item):
        if item in self.sheets:
            pass
        elif item in self.sheet_list:
            if self.library.recursive:
                self.load(item)
            else:
                raise KeyError(str(item) + ' is in Book ' + self.file_name +
                               ' but not loaded, and recursive loading is'
                               ' not enabled')
        else:
            extra_message = None
            for sheet in self.sheets:
                if sheet.lower() == item.lower():
                    extra_message = 'a sheet named ' + sheet + ' exists ' \
                         'however. sheets are case sensitive'
            raise KeyError(str(item) + ' is not in Book ' + self.file_name +
                           extra_message)
        return self.sheets[item]

    def load(self, sheet_to_load):
        # load sheet of name (arg) or else load everything
        # from xml module, response to simple 'if self._element_tree';
        # "FutureWarning: The behavior of this method will change in
        # future versions. Use specific 'len(elem)' or
        # 'elem is not None' test instead."
        if self._element_tree is not None:
            [self.add_sheet(Sheet(self, sheet))
             for sheet in self._element_tree.findall(self.ns('table:table'))
             if (sheet.attrib[self.ns('table:name')] == sheet_to_load or
                 sheet_to_load is None)]
        # check

    def add_sheet(self, sheet):
        self.sheets[sheet.name] = sheet

    def ns(self, string):
        return self.file.map.ns(string)


class Sheet:
    def __init__(self, book, element_tree):
        self._rows = []
        self._columns = []
        self._tree = element_tree
        self._attributes = self._tree.attrib
        self._loaded = False
        self.book = book
        self.name = [self._attributes[attribute]
                     for attribute in self._attributes
                     if not attribute.endswith('style-name') and
                     attribute.endswith('name')][0]

    def load(self):
        # row_elements = [element for element in self._tree if
        #                 element.tag.endswith('row')]
        # self._rows = [self._rows.append(Row(self, row_elements[y], y)) for
        #               y in range(0, len(row_elements) - 1)]
        row_elements = self._tree.findall(self.ns('table:table-row'))
        [self._rows.append(Row(self, y, row_elements[y])) for y in
         range(0, len(row_elements))]

    def return_cell(self, *strings):
        if len(strings) == 1:
            x, y = xy_from_a1(strings[0])
        else:
            x = strings[0]
            y = strings[1]
        return self._rows[y].return_cell(x)

    def __getitem__(self, item):
        # if item is tuple or list, return the referenced cell
        # if item is string, convert it from a1 and return the cell
        # if item is int, return the referenced row
        name = item.__class__.__name__
        # if self is not yet loaded, do that now
        self.load()
        if name == 'tuple' or name == 'list':
            x, y = item
        elif name == 'str':
            x, y = xy_from_a1(item)
        else:
            try:
                y = int(item)
            except:
                raise KeyError(str(item) + ' is not a valid key for sheet ' +
                               self.name)
            else:
                return self._rows[y]
        return self._rows[y][x]

    def ns(self, string):
        return self.book.file.map.ns(string)


class Row:
    def __init__(self, sheet, y, tree):
        self.y = y
        self.sheet = sheet
        self._tree = tree
        self._cells = []
        self._loaded = False

    def return_cell(self, x):
        if not self._cells:
            self.load()
        return self._cells[x]

    def __getitem__(self, item):
        if not self._loaded:
            self.load()
        return self._cells[item]

    def load(self):
        cell_elements = [element for element in self._tree
                         if element.tag.endswith('cell')]
        [self._cells.append(Cell(cell_elements[x], (x, self.y), self.sheet))
         for x in range(0, len(cell_elements) - 1)]


class Column:
    pass  # not used yet, here because it exists in spreadsheet xml
    # files and is used for formatting + possibly other uses. May be
    # utilized in the future


class Library:
    def __init__(self, list_of_addresses, recursive_loading=False):
        self.list = list_of_addresses
        self.files = [File(address, self) for address in self.list]
        self.books = {}
        self.recursive = recursive_loading

        [file.load() for file in self.files]

    def load_books(self):
        for book_name in self.list:
            self.load_book(book_name)

    def load_book(self, book_name):
        self.files[book_name] = File(book_name)
        self.files[book_name].load()

    def return_cell(self, *strings):
        # takes either xy or a1 cell ref
        # book, sheet, (row, cell) or (cell a1 ref)
        if strings[0] in self.books:
            book = self.books[strings[0]]
        elif '/' not in strings[0] and \
             any([book_name.endswith(strings[0]) for book_name in self.books]):
            book = [self.books[book_name] for book_name in self.books if
                    book_name.endswith(strings[0])][0]
        else:
            try:
                if self.recursive:
                    self.load_book(strings[0])
                    book = self.books[strings[0]]
                else:
                    raise SheetDataError('cell referenced cell not in loaded'
                                         'book')
            except:
                error_string = 'could not find referenced book' + str(
                        strings[0])
                raise SheetDataError(error_string)
        return book.return_cell(strings[1:])

    def __getitem__(self, item):
        # the new improved method for returning cells / books / sheets
        # should supplant return_cell method
        if item in self.books:
            book = self.books[item]
        elif '/' not in item and \
             any([book_name.endswith(item) for book_name in self.books]):
            # if there are no backspaces in item string, and one of the
            # library's books ends with the string 'item,'
            # that's the book
            book = [self.books[book_name] for book_name in self.books if
                    book_name.endswith(item)][0]
            # this has the potential for bugs, should be refined in the
            # future
        elif self.recursive:
            try:
                self.load_book(item)
                book = self.books[item]
            except:
                error_string = 'could not find referenced book' + str(
                       item)
                raise SheetDataError(error_string)
        else:
            error_string = 'cell referenced cell not in loaded'\
                                         'book' + str(item)
            raise SheetDataError(error_string)
        return book


class File:
    def __init__(self, file_address, library=None):
        self.file_id = file_address
        self.body = None
        self.map = None
        self.library = library

        if self.file_id[:7] == 'file://':
            self.file_id = self.file_id[7:]

        if library is None:
            self.library = Library([file_address])

    def load(self, sheet_name=None):
        # load cell instances into
        #
        # dictionary (library)
        # of dictionaries (book)
        # of lists (sheet)
        # of lists (row)
        # of objects (cell)

        # tree structure:
        # content
        #   body
        #       spreadsheet
        #           table
        #               table-row
        #                   table-cell

        # create element tree by extracting xml file and parsing it
        with zipfile.ZipFile(self.file_id).open('content.xml') as content:
            data = etree.parse(content).getroot()

            # get prefix map
            self.map = NSMap(data.nsmap)

            # set library.books [file_address]
            # to book of appropriate data
            if self.file_id not in self.library.books:
                body_element = data.find(self.ns('office:body'))
                spreadsheet_element = body_element.find(self.ns(
                        'office:spreadsheet'))
                self.library.books[self.file_id] = \
                    Book(self.library, self, spreadsheet_element)

            self.library.books[self.file_id].load(sheet_name)

    def ns(self, string):
        return self.map.ns(string)


class NSMap:
    # simple object for dealing with namespace mapping
    def __init__(self, map_dictionary):
        self.dict = map_dictionary

    def ns(self, s):
        # return string with prefix converted to namespace
        # string:tag --> {namespace}tag
        for x in range(0, len(s) - 1):
            if s[x] == ':':
                prefix = s[:x]
                suffix = s[x + 1:]
                s = '{%s}%s' % (self.dict[prefix], suffix)
                break
        return s


def a1_from_xy(pos):
    # returns excel-type cell given X, Y position
    # A corresponds to 0 in x-axis
    # 1 corresponds to 0 in y-axis

    x_ref = x_to_letter(pos[0] + 1)
    y_ref = str(pos[1] + 1)

    return x_ref + y_ref


def xy_from_a1(s):
    # returns x, y from inputted 'a1' string
    x = None
    y = None
    for a in range(0, len(s)):
        if s[a - 1] in storage.ENGLISH_ALPHABET and \
                s[a] in storage.ARABIC_NUMBERS:
            x = s[:a]
            y = s[a:]
            break

    if x is None:
        raise SheetDataError('could not parse reference to %s' % s)

    x = letter_to_x(x)
    y = int(y) - 1

    return x, y


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


def letter_to_x(s):
    # converts letter or letter series to arabic number value
    # example: a --> 0, ab --> 28
    s = s.lower()[::-1]
    val = -1  # starts at -1 so that a, when evaluated as 1, returns 0
    for x in range(0, len(s)):
        val += (ord(s[x]) - 96) * 26 ** x
    return val


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
    # this function is used for finding a string in a parent string
    # but only if it is not enclosed in quotes - useful for finding a
    # cell reference in a formula, where you want to find 'cell['
    # but not when used as in a script saying
    # print('you can access cells by typing cell[<reference>]')
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

PYSCRIPT_FLAG = 'py='  # flag denoting start of python script
