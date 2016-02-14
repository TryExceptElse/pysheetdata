from unittest import TestCase
import main


class TestFind_unquoted(TestCase):
    def test_find_unquoted(self):
        test_string_pairs = [
            ('Sheet2.A1', '.', False, False),
            ('Sheet2.A1', '.', True, False),
            ('Sheet2.A1', '.', False, True),
            ('Sheet2.A1', '.', True, True),
            ('A1', '.', False, False),  # should return None
            ('this is outside a quote & \' and this & is not \'', 'this',
             False, True),
            ('this is outside a quote & \' and this & is not \'', 'this',
             True, True),
            ('this should return \'None\'', 'None', False, False),
            ('this should return \'None\'', 'None', False, True),
            ('this should return \'None\'', 'None', True, True),
            ('multiples! multiples! multiples!', 'multiples!', True, True),
            ('multiples! multiples! multiples!', 'multiples!', False, True),
            ('file_ref_goes_here#$Sheet2.A1', '#$', True, False),
        ]

        correct_results = [
            6,
            6,
            [6],
            [6],
            None,
            [0],
            [0],
            None,
            [],
            [],
            [22, 11, 0],
            [0, 11, 22],
            18,
        ]

        for x in range(0, len(test_string_pairs)):
            test = test_string_pairs[x]
            correct = correct_results[x]
            test_result = main.find_unquoted(test[1], test[0], test[2],
                                             test[3])
            if test_result != correct:
                self.fail('test %s result %s did not match correct %s' %
                          (x, test_result, correct))
