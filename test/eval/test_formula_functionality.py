from unittest import TestCase

from eval.elements import Formula


class TestFormula(TestCase):
    def test_single_float_element_formula_returns_correctly(self):
        self.assertEqual(5, Formula('=5'))

    def test_multi_element_formula_returns_correctly(self):
        self.assertEqual(5, Formula('=B1 + 5 * (8^4)'))

    def test_multi_element_formula_returns_correctly_no_spaces(self):
        self.assertEqual(5, Formula('=B1+5*(8^4)'))
