from optparse import OptionParser

from mso import DOCXFile, XLSXFile
from utils import fileslist


def main(directory, search):
    result = []
    files = fileslist(directory, ('.docx', '.xlsx'))
    for filename in files:
        if filename.endswith('.docx'):
            reader = DOCXFile(filename)
        else:
            reader = XLSXFile(filename)

        if not reader.is_zipfile():
            err = "Error: File '{0}' is bad zip archive".format(filename)
            result.append(err)
            continue
        reader.open()
        if search in reader.get_text():
            result.append("File '{0}' contains string".format(filename))

    if result:
        print '\n'.join(result)
    else:
        print 'Search string not found'


if __name__ == '__main__':
    parser = OptionParser()
    parser.add_option('-d', '--dir', dest='dir',
                      help='Directory for search string in MSO files')
    parser.add_option('-s', '--search', dest='search',
                      help='Search string')
    options, _ = parser.parse_args()

    if not (options.dir and options.search):
        print parser.format_help()
    else:
        main(options.dir, options.search)
