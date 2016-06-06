from unittest import TestCase

import eval.elements as el


class TestElement(TestCase):
    def test_adding_integers_works_correctly(self):
        sum_element = el.Combined(5, '+') + el.Combined(6, '+')
        self.assertEqual(11, sum_element.value)

    def test_adding_strings_works_correctly(self):
        sum_element = el.Combined('Monty ') + el.Combined('Python', '+')
        self.assertEqual('Monty Python', sum_element.value)

    def test_adding_booleans_works_correctly(self):
        sum_element = el.Combined(True) + el.Combined(True, '+')
        self.assertEqual(2, sum_element.value)

    def test_adding_errors_works_correctly(self):


    def test_subtracting_integer_elements_works_correctly(self):
        remainder_element = el.Combined(8, '+') + el.Combined(6, '-')
        self.assertEqual(2, remainder_element.value)


class TestFormula(TestCase):
    pass


class TestFunctions(TestCase):
    def test_find_paired_bracket_returns_correct_index_with_single_layer(self):
        test_string = '5 * (2 * 7)'
        self.assertEqual(10, el._find_paired_closing_bracket(test_string, 4))

    def test_find_paired_bracket_returns_correctly_with_multiple_layers(self):
        test_string = '5*(2* (5 * (2 + (5 + 4))))'
        self.assertEqual(24, el._find_paired_closing_bracket(test_string, 6))