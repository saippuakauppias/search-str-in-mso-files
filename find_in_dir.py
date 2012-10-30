from mso import DOCXFile, XLSXFile
from utils import fileslist


def main():
    # TODO: upper case in filetype
    files = fileslist('/Users/saippuakauppias/Desktop', ('.docx', '.xlsx'))
    for filename in files:
        if filename.endswith('.docx'):
            reader = DOCXFile(filename)
        else:
            reader = XLSXFile(filename)

        if not reader.is_zipfile():
            print u"Error: File '{0}' is bad zip archive".format(filename)
            continue
        reader.open()
        text = reader.get_text()
        print text

if __name__ == '__main__':
    main()
