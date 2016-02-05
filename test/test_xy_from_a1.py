from unittest import TestCase

from main import xy_from_a1


class TestXy_from_a1(TestCase):

    def test_xy_from_a1(self):
        error_flag = False

        a1 = ['a1', 'b4', 'aa10', 'bc123']
        a1_result = []
        correct_a1 = [(0, 0), (1, 3), (26, 9), (54, 122)]

        [a1_result.append(xy_from_a1(test)) for test in a1]

        for x in range(0, len(correct_a1)):
            result = a1_result[x]
            correct = correct_a1[x]
            if result == correct:
                response = 'correct'
            else:
                response = 'error. correct:' + str(correct)
                error_flag = True
            print('%s --> %s (%s)' % (a1[x], result, response))

        if error_flag:
            self.fail('one or more incorrect conversions occurred')
