DOCX_XMLNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XLSX_XMLNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
XLSX_XMLNS_R = 'http://schemas.openxmlformats.org/officeDocument/'\
               '2006/relationships'


DOCX_XPATH = {
    'paragraph': './/{%s}p' % DOCX_XMLNS,
    'text': './/{%s}t' % DOCX_XMLNS,
    'tab': './/{%s}tab' % DOCX_XMLNS
}


XLSX_XPATH = {
    'text': './/{%s}t' % XLSX_XMLNS,
    'sheet': './/{%s}sheet' % XLSX_XMLNS,
    'value': './/{%s}v' % XLSX_XMLNS
}


XLSX_KEYS = {
    'r:id': '{%s}id' % XLSX_XMLNS_R,
    'sheetId': 'sheetId',
}
