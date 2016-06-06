from unittest import TestCase

import eval.elements as el


class TestFormula(TestCase):
    pass


class TestBracket(TestCase):
    def test_find_paired_bracket_returns_correct_index_with_single_layer(self):
        test_string = '5 * (2 * 7)'
        self.assertEqual(10, el._find_paired_closing_bracket(test_string, 4))

    def test_find_paired_bracket_returns_correctly_with_multiple_layers(self):
        test_string = '5*(2* (5 * (2 + (5 + 4))))'
        self.assertEqual(24, el._find_paired_closing_bracket(test_string, 6))
