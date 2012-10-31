import os
import unittest

from ..utils import fileslist


PROJECT_ROOT = os.path.join(os.path.dirname(__file__), '..')


class UtilsTests(unittest.TestCase):

    def setUp(self):
        self.files = fileslist(PROJECT_ROOT, ('.docx', '.xlsx'))

    def tearDown(self):
        del self.files

    def test_fileslist_count(self):
        self.assertEqual(len(self.files), 2)

    def test_fileslist_xlsx(self):
        self.assertIn(
            os.path.abspath(os.path.join(PROJECT_ROOT,
                                         'fixtures/xlsx/test.xlsx')),
            self.files
        )

    def test_fileslist_docx(self):
        self.assertIn(
            os.path.abspath(os.path.join(PROJECT_ROOT,
                                         'fixtures/docx/test.docx')),
            self.files
        )
