import main

test_library = main.Library(['file:///home/user3/PycharmProjects/pysheetdata/'
                             'testdata/test1.ods'], True)

for sheet in test_library.books:
    print(sheet)

# print(test_library.return_cell(['test1.ods', 'sheet1', 'd2']))
