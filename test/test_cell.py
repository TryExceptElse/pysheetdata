"""
should test most of the methods of the cell class
"""

from unittest import TestCase
import main

import os


class TestCell(TestCase):
    
    def setUp(self):
        # get address of test book
        self.ta = os.path.abspath('testdata/test1.ods')
        # create library object
        self.tl = main.Library([self.ta], True)
        
    def tearDown(self):
        self.ta = None
        self.tl = None
    
    # should try to keep these in the same order as main
    def test_evaluate(self):
        tests = [
            (self.tl[self.ta]['Sheet1']['b2'].content, 'cell1'),
            (self.tl[self.ta]['Sheet1']['c2'].content, 'cell2'),
            (self.tl[self.ta]['Sheet1']['d2'].content, 'cell1'),
            (self.tl[self.ta]['Sheet1']['e2'].content, 'test cell'),
        ]

        for test in tests:
            if test[0] != test[1]:
                self.fail()

    def test_dependencies_returns_cell_dependencies_from_formula_content(self):
        self.assertIn("['('.A1]",
                      self.tl[self.ta]['sheet1']['e2'].dependencies)

    def test_dependencies_returns_cell_dependencies_from_script_content(self):
        self.assertIn("cells[a1]",
                      self.tl[self.ta]['sheet1']['h2'].dependencies)

    def test_cell_can_store_and_return_values(self):
        testing_cell = self.tl[self.ta]['sheet1']['h2']
        testing_cell.value = 1
        self.assertEqual(1, testing_cell.value)

    def test_cached_value_returns_xml_value_despite_new_value_stored(self):
        testing_cell = self.tl[self.ta]['sheet1']['i2']
        testing_cell.value = 2468.
        self.assertEqual(1234, testing_cell.cached_value)

    def test_cell_can_store_and_return_text(self):
        testing_cell = self.tl[self.ta]['sheet1']['h2']
        testing_cell.text = 'a string'
        self.assertEqual('a string', testing_cell.text)

    def test_cached_text_returns_xml_text_despite_new_text_stored(self):
        testing_cell = self.tl[self.ta]['sheet1']['i1']
        testing_cell.text = 'new string'
        self.assertEqual('test cell', testing_cell.cached_text)

    def test_cell_can_set_and_return_content_property_string(self):
        testing_cell = self.tl[self.ta]['sheet1']['h2']
        testing_cell.content = 'content string'
        self.assertEqual('content string', testing_cell.content)

    def test_cell_can_set_and_return_content_property_float(self):
        testing_cell = self.tl[self.ta]['sheet1']['h2']
        testing_cell.content = 4.4
        self.assertEqual(4.4, testing_cell.content)

    def test_setting_string_content_sets_text_property(self):
        testing_cell = self.tl[self.ta]['sheet1']['h2']
        testing_cell.content = 'content string'
        self.assertEqual('content string', testing_cell.text)

    def test_setting_float_content_sets_value_property(self):
        testing_cell = self.tl[self.ta]['sheet1']['h2']
        testing_cell.content = 1234
        self.assertEqual(1234, testing_cell.value)

    def test_pyscript_cell_returns_correct_referenced_value(self):
        self.assertEqual('cell1',
                         self.tl[self.ta]['sheet1']['h2'].content)
        
    def test_inclduded_returns_correct_value(self):
        pass
