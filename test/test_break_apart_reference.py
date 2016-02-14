from unittest import TestCase
import main


class TestBreak_apart_reference(TestCase):
    def test_break_apart_reference(self):

        # test function from main to establish that it correctly
        # breaks down references

        test_cases = [
            'Sheet2.A1',
            'A1',
            'file_ref_goes_here#$Sheet2.A1'
        ]

        correct = [
            ['Sheet2', 'A1'],
            ['A1'],
            ['file_ref_goes_here', 'Sheet2', 'A1']
        ]

        for x in range(0, len(test_cases)):
            test_result = main.break_apart_reference(test_cases[x])
            correct_result = correct[x]
            if test_result != correct_result:
                self.fail('test %s result %s did not match correct return %s' %
                          (x, test_result, correct_result))
