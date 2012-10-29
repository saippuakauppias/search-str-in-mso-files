import zipfile
from lxml import etree


DOCX_XPATH = {
    'paragraph':
        './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p',
    'text':
        './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t',
    'tab':
        './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab'
}


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
        for p_element in self.document.findall(DOCX_XPATH['paragraph']):
            texts_list = []
            for t_element in p_element.findall(DOCX_XPATH['text']):
                if t_element.text:
                    texts_list.append(t_element.text)
            paragraphs_list.append(u''.join(texts_list))
        return u'\n'.join(paragraphs_list).encode('utf-8')
