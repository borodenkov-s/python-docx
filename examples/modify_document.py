"""
Creating a docx document from scratch and adding some elements to it.
"""
import os
import sys
import re

# adding the parent directory to PATH
path = os.path.abspath(os.path.join(os.path.dirname(__file__),".."))
sys.path.append(path)

from docx.document import DocxDocument
from docx.elements import *
from docx.meta import CoreProperties, WordRelationships


#doc = DocxDocument('modify.docx')
doc = DocxDocument('modify.docx')

# Replacing a string of text with another one.
doc.replace('This is a sample document', 'This is a modified document')

# replacing placeholder with picture
pic_paragraph = picture(doc,'python_logo.png','This is a test description')
doc.replace('IMAGE', pic_paragraph)

# Adding something to the end of the document.
doc.add(heading('Adding another element to the end of this document.',1))

# saving the new document
doc.save('modified_document.docx')
