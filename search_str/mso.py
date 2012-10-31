import zipfile
from lxml import etree

from mso_constants import DOCX_XPATH, XLSX_XPATH, XLSX_KEYS


class MSOFile(object):

    def __init__(self, filename):
        self.filename = filename

    def is_zipfile(self):
        return zipfile.is_zipfile(self.filename)

    def open(self):
        raise NotImplementedError(
            'This method must be implemented by subclasses'
        )

    def get_text(self):
        raise NotImplementedError(
            'This method must be implemented by subclasses'
        )


class DOCXFile(MSOFile):

    def open(self):
        zipped_doc = zipfile.ZipFile(self.filename)
        xml_document = zipped_doc.read('word/document.xml')
        self.document = etree.fromstring(xml_document)

    def get_text(self):
        paragraphs_list = []
        for p_element in self._get_paragraphs():
            texts_list = []
            for t_element in self._get_text_for_paragraph(p_element):
                if t_element.text:
                    texts_list.append(t_element.text)
            paragraphs_list.append(u''.join(texts_list))
        return u'\n'.join(paragraphs_list).encode('utf-8')

    def _get_paragraphs(self):
        return self.document.iterfind(DOCX_XPATH['paragraph'])

    def _get_text_for_paragraph(self, paragraph):
        return paragraph.iterfind(DOCX_XPATH['text'])


class XLSXFile(MSOFile):

    def open(self):
        self.zipped_table = zipfile.ZipFile(self.filename)

        xml_shared_strings = self.zipped_table.read('xl/sharedStrings.xml')
        shared_strings = etree.fromstring(xml_shared_strings)
        self.strings = self._parse_shared_strings(shared_strings)

        xml_workbook = self.zipped_table.read('xl/workbook.xml')
        workbook = etree.fromstring(xml_workbook)
        self.sheets = self._parse_workbook(workbook)

    def get_text(self):
        data_list = self.strings + self.sheets.values() + self._get_values()
        return u' '.join(data_list).encode('utf-8')

    def _get_values(self):
        values_list = []
        for sheet_id in self.sheets.keys():
            xml_sheet = self.zipped_table.read(
                'xl/worksheets/sheet{0}.xml'.format(sheet_id)
            )
            sheet = etree.fromstring(xml_sheet)
            for v_element in sheet.iterfind(XLSX_XPATH['value']):
                values_list.append(v_element.text)
#            values_list = list(set(values_list))
        return values_list

    def _parse_shared_strings(self, shared_strings):
        strings_list = []
        for t_element in shared_strings.iterfind(XLSX_XPATH['text']):
            strings_list.append(t_element.text)
        return strings_list

    def _parse_workbook(self, workbook):
        sheets_dict = {}
        for sheet_element in workbook.iterfind(XLSX_XPATH['sheet']):
            if sheet_element.get(XLSX_KEYS['r:id']):
                sheet_id = sheet_element.get(XLSX_KEYS['r:id'])[3:]
            else:
                sheet_id = sheet_element.get(XLSX_KEYS['sheetId'])

            sheet_name = u''
            if sheet_element.get('name'):
                sheet_name = sheet_element.get('name')

            sheets_dict.update({sheet_id: sheet_name})
        return sheets_dict
