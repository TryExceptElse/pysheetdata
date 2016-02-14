"""
This should grow to be by far the largest test file -
should test most of the methods of the cell class
"""

from unittest import TestCase
import main

import os


class TestCell(TestCase):
    # should try to keep these in the same order as main
    def test_evaluate(self):

        # get address of test book
        test_address = os.path.abspath('testdata/test1.ods')

        # create library object
        test_library = main.Library([test_address], True)

        tests = [
            (test_library[test_address]['Sheet1']['b2'].content, 'cell1'),
            (test_library[test_address]['Sheet1']['c2'].content, 'cell2'),
            (test_library[test_address]['Sheet1']['d2'].content, 'cell1'),
            (test_library[test_address]['Sheet1']['e2'].content, 'test cell'),
        ]

        for test in tests:
            if test[0] != test[1]:
                self.fail()
