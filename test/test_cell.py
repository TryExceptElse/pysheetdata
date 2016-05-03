"""
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

    def test_dependencies_returns_cell_dependencies_from_formula_content(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        self.assertIn("['('.A1]",
                      test_library[test_address]['sheet1']['e2'].dependencies)

    def test_dependencies_returns_cell_dependencies_from_script_content(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        self.assertIn("cells[a1]",
                      test_library[test_address]['sheet1']['h2'].dependencies)

    def test_cell_can_store_and_return_values(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        testing_cell = test_library[test_address]['sheet1']['h2']
        testing_cell.value = 1
        self.assertEqual(1, testing_cell.value)

    def test_cached_value_returns_xml_value_despite_new_value_stored(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        testing_cell = test_library[test_address]['sheet1']['i2']
        testing_cell.value = 2468.
        self.assertEqual(1234, testing_cell.cached_value)

    def test_cell_can_store_and_return_text(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        testing_cell = test_library[test_address]['sheet1']['h2']
        testing_cell.text = 'a string'
        self.assertEqual('a string', testing_cell.text)

    def test_cached_text_returns_xml_text_despite_new_text_stored(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        testing_cell = test_library[test_address]['sheet1']['i1']
        testing_cell.text = 'new string'
        self.assertEqual('test cell', testing_cell.cached_text)

    def test_cell_can_set_and_return_content_property_string(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        testing_cell = test_library[test_address]['sheet1']['h2']
        testing_cell.content = 'content string'
        self.assertEqual('content string', testing_cell.content)

    def test_cell_can_set_and_return_content_property_float(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        testing_cell = test_library[test_address]['sheet1']['h2']
        testing_cell.content = 4.4
        self.assertEqual(4.4, testing_cell.content)

    def test_setting_string_content_sets_text_property(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        testing_cell = test_library[test_address]['sheet1']['h2']
        testing_cell.content = 'content string'
        self.assertEqual('content string', testing_cell.text)

    def test_setting_float_content_sets_value_property(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        testing_cell = test_library[test_address]['sheet1']['h2']
        testing_cell.content = 1234
        self.assertEqual(1234, testing_cell.value)

    def test_pyscript_cell_returns_correct_referenced_value(self):
        test_address = os.path.abspath('testdata/test1.ods')
        test_library = main.Library([test_address], True)
        self.assertEqual('cell1',
                         test_library[test_address]['sheet1']['h2'].content)
