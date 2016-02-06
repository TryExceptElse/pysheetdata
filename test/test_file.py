from unittest import TestCase
import os

import main


class TestFile(TestCase):
    def test_load(self):

        test_address = os.path.abspath('testdata/test2.ods')

        self.test_file = main.File(test_address)

        self.test_file.load()





