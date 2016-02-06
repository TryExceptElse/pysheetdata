from unittest import TestCase

import main
import os


class TestLibrary(TestCase):
    def test___getitem__(self):

        test_address = os.path.abspath('testdata/test1.ods')

        test_address_b = os.path.abspath('testdata/test2.ods')

        test_library = main.Library([test_address], True)
        
        results = [
            test_library['test1.ods']['Sheet1']['d2'],
            test_library[test_address]['Sheet1'][(3, 1)],
            test_library[test_address]['Sheet1'][1][3],
            test_library[test_address]['Sheet2']['a1'],
            test_library[test_address]['(']['a1'],
            test_library[test_address_b]['Sheet1']['b1'],
        ]

        if not results[0] == results[1] == results[2]:
            self.fail('equivalent references do not return the same cell')

        text_results = [entry.text for
                        entry in results]

        text_correct = ['cell1', 'cell1', 'cell1', 'cell2', 'test cell',
                        'referencer']

        if text_results != text_correct:
            for x in range(0, len(text_correct)):
                response = text_results[x]
                correct = text_results[x]
                extra = None
                if response != correct:
                    extra = '(error)'
                print('response: %s, correct: %s %s' %
                      (response, correct, extra))
            self.fail('incorrect text entries are being returned')

