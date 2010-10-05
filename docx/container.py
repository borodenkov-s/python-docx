import collections
import os
import zipfile
from docx import TEMPLATE_DIR
from docx.utils import make_element

class DocxFile(object):
    def __init__(self, file, template=None):
        self.container=zipfile.ZipFile(file, "w")
        self._add_template(template if template is not None else TEMPLATE_DIR)
        self.document=make_element("document")
        self.body=make_element("body")
        self.document.append(self.body)


    def _add_template(self, template):
        """Copy a template document to our container."""

        for (dirpath, dirnames, filenames) in os.walk(template):
            for filename in filenames:
                path=os.path.join(dirpath, filename)
                self.container.write(path, os.path.relpath(path, template))

