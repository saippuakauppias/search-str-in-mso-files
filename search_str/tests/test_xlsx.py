# -*- coding: utf-8 -*-

import os
import unittest

from ..mso import XLSXFile


PROJECT_ROOT = os.path.join(os.path.dirname(__file__), '..')


class SimpleXLSXTests(unittest.TestCase):

    def setUp(self):
        filepath = os.path.join(PROJECT_ROOT, 'fixtures/xlsx/test.xlsx')
        filepath = os.path.abspath(filepath)
        self.table = XLSXFile(filepath)

    def tearDown(self):
        del self.table

    def test_zipfile(self):
        self.assertTrue(self.table.is_zipfile())

    def test_shared_strings(self):
        self.table.open()
        strings = self.table.strings
        self.assertEqual(len(strings), 2)

    def test_sheets(self):
        self.table.open()
        sheets = self.table.sheets
        self.assertEqual(len(sheets), 3)
        self.assertEqual(sheets.keys(), ['1', '3', '2'])
        self.assertEqual(sheets.values(),
                         [u'ListN1', u'Лист3', u'Лист номер 2'])

    def test_search_latin_symbols(self):
        self.table.open()
        self.assertIn('English string', self.table.get_text())

    def test_search_cyrillic_symbols(self):
        self.table.open()
        self.assertIn('Русский', self.table.get_text())

    def test_search_values(self):
        self.table.open()
        # values
        self.assertIn('5', self.table._get_values())
        self.assertIn('7', self.table._get_values())
        # compiled formula
        self.assertIn('35', self.table._get_values())
