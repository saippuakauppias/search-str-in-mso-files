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
            result.append("\033[1mError\033[0m: File '\033[4m{0}\033[0m'"
                          " is bad zip archive".format(filename))
            continue
        reader.open()
        if search in reader.get_text():
            result.append("File '\033[4m{0}\033[0m' \033[1mcontains"
                          " string\033[0m".format(filename))

    if result:
        print '\n'.join(result)
    else:
        print '\033[1mSearch string not found\033[0m'


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
