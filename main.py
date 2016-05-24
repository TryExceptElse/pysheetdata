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
    return library or file

scripting variables:
    cells[spreadsheet reference] - gets a cell object, or matrix of objects
    sheet.rows[y]
    sheet.columns
"""

import zipfile
import lxml.etree as etree
import os.path as path

import eval.storage as storage
import eval.parser as parser
import script


class SheetDataError(Exception):
    pass


class LibComponent:
    """
    abstract library component class, sub-classed by cell, sheet, book, lib
    """
    def __init__(self):
        self.parent = None
        self._settings = {}
        self._sheet = None
        self._book = None
        self._lib = None

    def __str__(self):
        return '%s %s' % (self.class_name, self.identifier)

    @property
    def identifier(self):
        raise NotImplementedError('LibComponent should only be inherited from')

    @property
    def class_name(self):
        return self.__class__.__name__

    @property
    def inc_set(self):
        return self._settings.get('inc_set', True)

    @property
    def included(self):
        if all([parent.inc_set for parent in self.parents]):
            return True

    @property
    def parents(self):
        # returns list of parent row, column, sheet, book, lib
        return [getattr(self, parent_s) for parent_s in
                ['row', 'column', 'sheet', 'book', 'lib']
                if getattr(self, parent_s) is not None]

    #######################
    # parent object getters
    # iterates through hierarchy in order and attempts to find
    # appropriate instance

    @property
    def row(self):
        try:
            return self._row  # only cells have this attr
        except AttributeError:
            raise NotImplementedError('%s does not have a row parent obj'
                                      % self)

    @property
    def column(self):
        try:
            return self._column  # see above
        except AttributeError:
            raise NotImplementedError('%s does not have a column parent obj'
                                      % self)

    @property
    def sheet(self):
        return self.get_parent('Sheet')

    @property
    def book(self):
        return self.get_parent('Book')

    @property
    def lib(self):
        return self.get_parent('Library')

    @property
    def library(self):
        # backwards compatibility, hurrah.
        return self.get_parent('Library')

    @property
    def file(self):
        return self.get_parent('Book').file

    @file.setter
    def file(self, file):
        self.get_parent('Book').file = file

    def get_parent(self, hierarchy_s=None):
        # return parent in hierarchy of name 'hierarchy_s'
        # if none, returns direct parent of inst.
        if self.class_name == hierarchy_s:
            return self
        parent_attr = getattr(self, '_parent', None)
        if hierarchy_s is None:
            hierarchy_s = 'parent'
        if parent_attr is None:
            raise NotImplementedError(self.class_name + ' does not have a ' +
                                      hierarchy_s)
        if hierarchy_s is None:
            return self.parent
        else:
            return parent_attr.get_parent(hierarchy_s)

    def ns(self, string):
        # returns namespace instance for lib comp
        return self.file.map.ns(string)


class Cell(LibComponent):
    # in some sheets there are going to be a LOT of these, should be
    # coded as such.
    # 78,000 plus in some sheets that I'll be using this for,
    # and I'm sure others will have far more
    def __init__(self, cell_data, position, sheet, mode='standard'):
        super().__init__()
        self.mode = mode
        # changing mode allows different amounts of data to be stored
        # for each cell. speed vs memory usage
        self._cell_element = cell_data
        self._position = position
        self._cached_position = position
        self._parent = sheet
        # stores new values, without applying them to _cell_element
        self._new_attrib = {}
        self._new_text = None
        self._new_script = None
        self._change_flag = False  # not yet implemented
        # output vars

        # may be added in the future to speed up program
        # self._dependencies
        # self._dependants

    @property
    def identifier(self):
        return self.a1

    @property
    def has_contents(self):
        if self._cell_element is not None:  # for etree reasons, cannot
                # just use self._cell_element
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
    def content(self):
        # will be refactored as 'value' to return either string
        # or float value
        if self.data_type == 'string':
            self.evaluate()
            return self.text
        else:
            return self.value

    @content.setter
    def content(self, content):
        try:
            self.value = float(content)
        except ValueError:
            self.text = content

    @property
    def value(self):
        # at the moment, reevaluates all cells in dependency tree
        # in the future, once dependants and change flags are
        # instituted, will only reevaluate if a cell in the tree has
        # changed
        self.evaluate()
        return float(self.get('office:value'))

    @value.setter
    def value(self, value):
        # note: this only returns a numerical value only
        # this will be refactored soon to return either str or float
        self.set('office:value', float(value))
        self.data_type = 'float'

    @property
    def cached_value(self):
        # always returns original value from the xml file
        return float(self.get('office:value', True))

    @property
    def text(self):
        # text getter/setter is different from other attributes because
        # it is stored separately in the xml file
        if self._new_text is not None:
            return self._new_text
        return self._cell_element[0].text

    @text.setter
    def text(self, string):
        # not using self.set because text is stored separately
        # (ask whoever started that standard, no idea why)
        self.data_type = 'string'
        if string != self.text:
            self._new_text = string

    @property
    def cached_text(self):
        return self._cell_element[0].text  # get text directly from ET

    @property
    # this is the formula as modified to appear as it does as typed
    # by a user in the spreadsheet program (...eventually)
    def formula(self):
        return self.get('table:formula')[3:]

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
        # if a new script has been set, use that
        # otherwise, use the cached text to see if self is script.
        # scripts themselves can change self text value.
        if self._new_script:
            return self._new_script
        if self.is_script:
            return self.text[len(PYSCRIPT_FLAG):]

    @script.setter
    def script(self, script_string):
        self._new_script = script_string

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
    def column(self):
        return self.sheet.columns[self.position[0]]

    @property
    def row(self):
        return self.sheet[self.position[1]]

    @property
    # returns dictionary of dependencies of self.
    # referencing string is key
    # may just have them be stored as standard
    # I -don't think- this needs to be sent over to _new_attrib, but
    # if that proves to be the case, this will need to be updated.
    def dependencies(self):
        return self.find_dependencies()

    #############################
    # Output props
    @property
    def included(self):
        # if all parent objects are set to be included in output,
        # returns true as cell is to be included in output file
        if all([self.book.inc_set,
                self.sheet.inc_set,
                self.row.inc_set,
                self.column.inc_set]):
            return True

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
        if entry != self.get(key):
            ns_key = self.ns(key)
            self._new_attrib[ns_key] = entry

    def find_dependencies(self):
        # returns dictionary of referenced cells with reference as key
        # if reference is a range (like 'A1:B2') returns matrix
        if not self.has_contents or (self.raw_formula is None and
                                     self.script is None):
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
                        reference = self.raw_formula[start + 1: x]
                        # set default values, will be changed in a
                        # moment if needed
                        dependencies['[' + reference + ']'] = \
                            self.cell_from_reference(reference)
            if self.script is not None:
                start = None
                script_s = self.script
                for x in range(0, len(script_s)):
                    if script_s[x: x + 6] == 'cells[':
                        start = x
                    elif script_s[x] == ']' and \
                            start is not None:
                        reference = script_s[start + 7:x - 1]
                        start = None
                        dependencies['cells[' + reference + ']'] = \
                            self.cell_from_reference(reference)
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
            reference_parts = self.complete_reference_parts(*reference_parts)
            return self.library[reference_parts]
        elif reference_type == 'range':
            start_parts = break_apart_reference(range_start)
            start_parts = self.complete_reference_parts(
                    start_parts)
            end_parts = break_apart_reference(range_end)
            end_parts = self.complete_reference_parts(
                    end_parts)
            return CellRange(
                self.sheet.book.library[start_parts],
                self.sheet.book.library[end_parts]
            )

    def complete_reference_parts(self, *parts):
        # unfilled parts of reference are completed with
        # self sheet and book
        sheet_default = self.sheet.name
        book_default = self.sheet.book.file_name
        length = len(parts)
        if 0 < length <= 3:  # if length is within possible ref. depths
            additions = [book_default, sheet_default]
            # add parts of reference that are missing from reference
            return additions[0: 3 - length] + list(parts)
        else:
            raise SheetDataError('len of reference parts outside 1-3 range')

    def evaluate(self, recursive=True, scripts=True):
        if self.has_contents:
            if recursive:
                for cell in self.dependencies:
                    self.dependencies[cell].evaluate()
            if scripts and self.script:
                self.run_script()
            elif self.raw_formula:
                evaluation = parser.evaluate(self.formula,
                                             self.dependencies)
                # if returned is value, put it in self.value
                # if that doesn't work, put it in self.string
                try:
                    self.value = float(evaluation)
                except ValueError:
                    try:
                        self.text = str(evaluation)
                    except ValueError:
                        print('could not set new value')
                        pass

    def run_script(self, auto_import=True):
        # dependencies is list of cells that are used

        class CellReferencer:
            def __init__(self, lookup_formula):
                self.lookup = lookup_formula

            def __getitem__(self, reference_string):
                return self.lookup(reference_string)

        if self.has_contents and self.is_script:
            # set script vars that can be called by the script
            script.cell = self
            script.cells = CellReferencer(self.cell_from_reference)
            script_string = self.script
            if auto_import:
                script_string = 'from script import *\n' + self.script
            exec(script_string)
        else:
            return ''


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
        self._matrix = []
        for sheet in start_cell.sheet.book.sheet_list[indexes['sheets'][0]:
                                                      indexes['sheets'][1]]:
            sheet_matrix = []
            self._matrix.append(sheet_matrix)
            for row in sheet.matrix[indexes['rows'][0]:indexes['rows'][1]]:
                row_matrix = []
                sheet_matrix.append(row_matrix)
                for cell in row[indexes['cells'][0]:indexes['cells'][1]]:
                    row_matrix.append(cell)

    @property
    def identifier(self):
        return '%s %s' % (self.start_cell.a1, self.end_cell.a1)

    @property
    def start_cell(self):
        return self.matrix[1][1][1]

    @property
    def end_cell(self):
        return self.matrix[-1][-1][-1]

    @property
    def matrix(self):
        return self._matrix

    @matrix.setter
    def matrix(self, matrix):
        self._matrix = matrix

    @property
    def cells(self):
        cells = []
        [[[cells.append(cell) for cell in row] for row in sheet] for sheet in
         self.matrix]
        return cells

    @property
    def value(self):
        return sum([sum([sum([cell.value for cell in row])
                   for row in sheet]) for sheet in self.matrix])

    @property
    def dependencies(self):
        # returns dictionary of dependencies
        # reference string is key, referenced cell or range is value
        d = {}
        for sheet in self.matrix:
            for row in sheet:
                for cell in row:
                    d.update(cell.dependencies)
        return d

    def evaluate(self):  # evaluates each cell in matrix
        [[[cell.evaluate() for cell in row]
          for row in sheet]
         for sheet in self.matrix]


class Row(LibComponent):
    def __init__(self, sheet, y, tree):
        super().__init__()
        self.y = y
        self._parent = sheet
        self._tree = tree
        self._cells = []
        self._loaded = False
        self._settings = {}

    def __getitem__(self, item):
        if not self._loaded:
            self.load()
        return self._cells[item]

    @property
    def identifier(self):
        return self.y

    @property
    def list(self):
        return [cell.content for cell in self._cells]

    def load(self):
        cell_elements = self._tree.findall(self.ns('table:table-cell'))
        [self._cells.append(Cell(cell_elements[x], (x, self.y), self.sheet))
         for x in range(0, len(cell_elements))]
        self._loaded = True


class Column(LibComponent):
    # used by spreadsheet xml for storing formatting, also useful for
    # references.
    def __init__(self, sheet, x, tree):
        super().__init__()
        self.x = x
        self.parent = sheet
        self._tree = tree

    def __getitem__(self, y):
        # unlike row, does not store cells.
        return self.sheet[(self.x, y)]

    @property
    def identifier(self):
        return x_to_letter(self.x)


class Sheet(LibComponent):
    def __init__(self, book, element_tree):
        super().__init__()
        self._rows = []
        self._columns = []
        self._tree = element_tree
        self._attributes = self._tree.attrib
        self._loaded = False
        self._parent = book
        self.name = [self._attributes[attribute]
                     for attribute in self._attributes
                     if not attribute.endswith('style-name') and
                     attribute.endswith('name')][0]
        self._settings = {}

    def __getitem__(self, item):
        # if item is tuple or list, return the referenced cell
        # if item is string, convert it from a1 and return the cell
        # if item is int, return the referenced row
        # if self is not yet loaded, do that now
        if not self._loaded:
            self.load()
        if isinstance(item, (list, tuple)):
            x, y = item
        elif isinstance(item, str):
            x, y = xy_from_a1(item)
        else:
            try:
                y = item
            except:
                raise KeyError(str(item) + ' is not a valid key for sheet ' +
                               self.name)
            else:
                return self._rows[y]
        return self._rows[y][x]

    @property
    def identifier(self):
        return self.name

    @property
    def columns(self):
        return self._columns

    @property
    def matrix(self):
        """
        returns a matrix of sheet data
        :return: matrix
        """
        return [row.list for row in self._rows]

    def load(self):
        # loads rows and columns into sheet lists

        # load rows
        row_elements = self._tree.findall(self.ns('table:table-row'))
        [self._rows.append(Row(self, y, row_elements[y])) for y in
         range(0, len(row_elements))]
        self._loaded = True

        # load columns
        column_elements = self._tree.findall(self.ns('table:table-column'))
        x = 0
        for column_tree in column_elements:
            repeat = column_tree.get(self.ns(
                    'table:number-columns-repeated'))
            if repeat:
                for x in range(0, int(repeat)):
                    self._columns.append(Column(self, x, column_tree))
                    x += 1
            else:
                self._columns.append(Column(self, x, column_tree))
                x += 1

    def find(self, column_or_row, name_to_find,
             name_index=0, case_sensitive=False):
        if column_or_row == 'row' or column_or_row == 'rows':
            list_to_search = self._rows
        elif column_or_row == 'column' or column_or_row == 'columns':
            list_to_search = self._columns
        else:
            raise KeyError('list to search should be \'row\' or \'column\', '
                           'not %s' % column_or_row)
        if not case_sensitive:
            name_to_find = name_to_find.lower()
        result = None
        for row in list_to_search:
            row_name = row[name_index].text
            if not case_sensitive:
                row_name = row_name.lower()
            if row_name == name_to_find:
                if result is None:
                    result = row
                else:
                    raise KeyError('multiple rows with name \'%s\'' %
                                   name_to_find)


class Book(LibComponent):
    # gets passed relevant dictionary of elements by file.load()
    # should not load cells until that function is called
    # then pass on relevant sub-dictionary to sheets that are created
    # so on
    def __init__(self, library, file, element_tree):
        super().__init__()
        self._file = file
        self._element_tree = element_tree
        self._parent = library
        self._settings = {}
        self.sheets = {}
        self.sheet_list = []  # list of sheets, in order of file

    @property
    def identifier(self):
        return path.split(self._file.file_id)[-1]

    @property
    def file_name(self):
        return self.file.file_id

    @property
    def file(self):
        return self._file

    @file.setter
    def file(self, file):
        self._file = file

    @property
    def dictionary(self):
        """
        creates and returns dictionary of book content
        :return: content dictionary
        """
        return {sheet.name: sheet.matrix for sheet in self.sheets}

    def __getitem__(self, item):
        item = strip_quotes(item.lower())
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
            raise KeyError(str(item) + ' is not in Book ' + self.file_name)
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

    def write(self, name, overwrite):
        """
        writes content dictionary to a file of name
        :param name: name to write as.
        :param overwrite: Bool, write over existing file of name
        """

    def add_sheet(self, sheet):
        self.sheets[sheet.name.lower()] = sheet


class Library(LibComponent):
    def __init__(self, list_of_addresses, recursive_loading=False):
        super().__init__()
        self.list = list_of_addresses
        self.files = {address: File(address, self) for address in self.list}
        self.books = {}
        self.recursive = recursive_loading
        self._settings = {}

        [self.files[file].load() for file in self.files]

    @property
    def identifier(self):
        return 'base'

    def load_books(self):
        for book_name in self.list:
            self.load_book(book_name)

    def load_book(self, book_name):
        self.files[book_name] = File(book_name, self)
        self.files[book_name].load()

    def __getitem__(self, item):
        # the new improved method for returning cells / books / sheets
        # if item is list, break apart and recall self
        # (allows passing of lists of getters)
        if isinstance(item, (list, tuple)):
            return self[item[0]][item[1]][item[2]]
        item = remove_file_prefix(strip_quotes(item))
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
                error_string = 'could not find referenced book ' + str(
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

    def ns(self, s):
        return self.map.ns(s)


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
    s = s.lower()
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

    # since the above worked from the right
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
        parts.append(s[:dot])
    else:
        parts.append(s[hd + 2: dot])
        parts.append(s[:hd])
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
            elif string[index: index + target_length] == target:
                matches.append(index)
            index += motion
        return matches
    else:
        while string[index: index + target_length] != target:
            if string[index] == "'":
                index += motion
                while string[index] != "'":
                    index += motion
            index += motion
            if not 0 <= index < len(string):
                return
        return index

PYSCRIPT_FLAG = 'py='  # flag denoting start of python script


def strip_quotes(string):
    def is_quote(char):
            if char == '\'' or char == '\"':
                return True
            else:
                return False

    if is_quote(string[0]) and is_quote(string[-1]):
        string = string[1:-1]
    return string


def remove_file_prefix(string):
    pre = 'file://'
    if string.startswith(pre):
        string = string[len(pre):]
    return string
