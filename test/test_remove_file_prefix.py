from main import remove_file_prefix
from unittest import TestCase


class TestRemoveFilePrefix(TestCase):
    def test_remove_file_prefix_removes_file_string_prefix(self):
        self.assertEqual('/testdata/test1.ods',
                         remove_file_prefix('file:///testdata/test1.ods'))
