from unittest import TestCase

from main import a1_from_xy


class TestA1_from_xy(TestCase):
    def test_xy_from_a1(self):
        error_flag = False

        xy = [(0, 0), (1, 1), (27, 3), (60, 60)]
        xy_result = []
        correct_xy = ['a1', 'b2', 'ab4', 'bi61']

        [xy_result.append(a1_from_xy(test)) for test in xy]

        for x in range(0, len(correct_xy)):
            result = xy_result[x]
            correct = correct_xy[x]
            if result == correct:
                response = 'correct'
            else:
                response = 'error. correct:' + str(correct)
                error_flag = True

            print('%s --> %s (%s)' % (xy[x], result, response))

        if error_flag:
            self.fail('one or more incorrect conversions occurred')