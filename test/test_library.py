from unittest import TestCase

import main
import os


class TestLibrary(TestCase):
    def test___getitem__(self):

        test_address = os.path.abspath('testdata/test1.ods')

        test_library = main.Library([test_address], True)

        for sheet in test_library.books:
            print(sheet)

        print(test_library['test1.ods']['Sheet1']['d2'])