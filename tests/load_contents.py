import main
import os

test_address = os.path.abspath('testdata/test2.ods')

test_library = main.Library(['file:///media/user0/raid1a1/docs/pycharm/'
                             'pysheetdata/test/testdata/test2.ods'], True)

for sheet in test_library.books:
    print(sheet)

print(test_library.return_cell('test1.ods', 'sheet1', 'd2'))
