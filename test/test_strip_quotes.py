from unittest import TestCase

from main import strip_quotes


class TestStripQuotes(TestCase):
    def test_strip_quotes_removes_quotes_from_string_margins(self):
        self.assertEqual('string_with quotes',
                         strip_quotes('\'string_with quotes\''))

    def test_strip_quotes_removes_double_quotes_from_string_edges(self):
        self.assertEqual('string', strip_quotes('"string"'))

    def test_strip_quotes_does_not_remove_from_center_of_string(self):
        self.assertEqual('string_""with_quotes',
                         strip_quotes('string_""with_quotes'))

    def test_strip_quotes_returns_passed_string_if_no_quotes(self):
        self.assertEqual('string', strip_quotes('string'))
