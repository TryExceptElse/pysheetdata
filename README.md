=======================================================================
pysheetdata
=======================================================================
For extracting and working with spreadsheet data


WARNING: PYSHEETDATA IS NOT SECURE AGAINST MALICIOUSLY CONSTRUCTED DATA

it is intended to be used only with data that is trusted to not have been built to exploit potential security holes

Updates:
===

As of 2/13/16 - basic functionality achieved.

Can load workbooks or groups of workbooks in 'libraries'

For any cell in a book, can get the cell data type, value, string, formula, or...

Can recognize and run python scripts in a cell whose contents begins with 'py='

Can parse cell formulas to return their value. Very few excel functions are currently supported, but more will be added in the future

While basic functions of this module are complete, for the most part it should be considered highly unfinished, and relied (or not relied) upon with this in mind



Basic use:
=======================================================================

Libraries / Books / Sheets
=======================================================================

Library(list of addresses, enable recursive loading) : initialize Library

list of addresses: list of addresses of books that are to be loaded into the library
recursive loading: allow books that have cells with references to non-loaded books to load those into the library



example:

lib = Library([address1, address2, etc], True)

cell = lib[book_address]['Sheet1']['d2'] : return cell object in library

cell = lib[book_address]['Sheet1'][(3, 1)] : return same cell as above, in xy format

cell = lib[book_address]['Sheet1'][1][3] : return same cell again, via getting first row, then cell in row: note the order has reversed

Cells
=======================================================================

useful cell properties:

.library : If present, returns the library that has loaded the cell, otherwise returns none

.book : returns the book object the cell belongs to

.has_contents : returns true if the cell has contents

.data_type : returns 'string' or 'float' depending on the stored data. Note that cells with scripts will still return 'string'

.content : returns the cell string or value depending on which is present.

.value : returns the float value of the cell, if present. will evaluate the cell formula and dependencies before returning

.cached_value : returns the value of the cell as stored in the workbook. useful if the cell's value has been changed

.text : returns the text of the cell - will return the value as it appears in a spreadsheet if it is a float

.cached_text : returns the cell's text that is stored in the workbook

.raw_formula : returns the cell's formula as it appears in the workbook xml file. this is not the same as it appears to users.

.is_script : returns true if the cell's text begins with a script flag

.a1 : returns the a1 style position of the cell within its sheet

.position : returns the x, y position of the cell within its sheet. In this format, 'a1' = (0, 0)

.cached_position : returns the position of the cell as it was loaded from the workbook

.dependencies : returns the dependencies of the cell, either that are in its formula, or a script that uses cells['ref']

note on dependencies:

Dependencies can be either a cell, or a range object. Ranges have less data available than a cell.




Ranges
=======================================================================

Ranges can be within a single column or row, or even 3d ranges that stretch across multiple sheets.

everything in this module is likely to contain bugs, but ranges particuarly so as they have not yet been tested. Wait for 
another commit or two. Or help out and get these working smoothly!

The properties available for ranges are:

.matrix : returns the matrix of cells in a 3 dimensional matrix of sheet, row, cell. When using, remember the sheet dimension

.cells : returns a list of all cells contained in the range

.value : returns the combined sum of all cells within the range

.dependencies : returns the list of cells and ranges that are dependencies of any of the cells in the range


