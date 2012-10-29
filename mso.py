import zipfile
import lxml


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
        pass
