search-str-in-mso-files
======================

Small script for search string in many Microsoft Office files (like *.docx and *.xlsx).


How it works
------------

Search string in many MS Office files (*.docx, *.xlsx) in directory:

    $ python find_in_dir.py -d path/to/dir -s 'search string'
    File '/path/to/dir/first_doc.docx' contains string
    File '/path/to/dir/last_table.xlsx' contains string


Code structure
--------------

mso.py:

    class MSOFile(filename):
        __init__(filename)
        is_zipfile -> bool (check zipfile.is_zipfile)
        open -> write in self.document xml tree
        get_text -> unicode

    class DOCXFile(MSOFile):
        new implementations of some methods

    class XLSXFile(MSOFile):
        new implementations of some methods

find-in-dir.py:

    optparse!
