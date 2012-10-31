# -*- coding: utf-8 -*-

import os
import unittest

from ..mso import DOCXFile


PROJECT_ROOT = os.path.join(os.path.dirname(__file__), '..')


class SimpleDOCXTests(unittest.TestCase):

    def setUp(self):
        filepath = os.path.join(PROJECT_ROOT, 'fixtures/docx/test.docx')
        filepath = os.path.abspath(filepath)
        self.doc = DOCXFile(filepath)

    def tearDown(self):
        del self.doc

    def test_zipfile(self):
        self.assertTrue(self.doc.is_zipfile())

    def test_paragraphs_count(self):
        self.doc.open()
        paragraphs = list(self.doc._get_paragraphs())
        self.assertEqual(len(paragraphs), 3)

    def test_text_elements(self):
        self.doc.open()
        paragraphs = list(self.doc._get_paragraphs())

        text_element = list(self.doc._get_text_for_paragraph(paragraphs[0]))
        self.assertEqual(len(text_element), 1)

        text_element = list(self.doc._get_text_for_paragraph(paragraphs[1]))
        self.assertEqual(len(text_element), 0)

        text_element = list(self.doc._get_text_for_paragraph(paragraphs[2]))
        self.assertEqual(len(text_element), 1)

    def test_search_latin_symbols(self):
        self.doc.open()
        self.assertIn('English string', self.doc.get_text())

    def test_search_cyrillic_symbols(self):
        self.doc.open()
        self.assertIn('Русский', self.doc.get_text())
