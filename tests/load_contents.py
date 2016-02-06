import main
import os

test_address = os.path.abspath('testdata/test1.ods')

test_library = main.Library(['file:///media/user0/raid1a1/docs/pycharm/'
                             'pysheetdata/test/testdata/test1.ods'], True)

for sheet in test_library.books:
    print(sheet)

print(test_library['test1.ods']['Sheet1']['d2'])
