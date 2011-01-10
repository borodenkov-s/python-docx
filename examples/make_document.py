"""
Creating a docx document from scratch and adding some elements to it.
"""
import os
import sys

# adding the parent directory to PATH
path = os.path.abspath(os.path.join(os.path.dirname(__file__),".."))
sys.path.append(path)

from docx.document import DocxDocument
from docx.elements import *
from docx.meta import CoreProperties, WordRelationships

if __name__ == '__main__': 
    # creating a new document with a template dir specified
    template_path = os.path.abspath(os.path.join(os.path.dirname(__file__),'template'))
    doc = DocxDocument(template_dir=template_path)

    # appending various elements to the newly created document.
    doc.add(heading('''Welcome to Python's docx module''',1)  )   
    doc.add(heading('Make and edit docx in 200 lines of pure Python',2))
    doc.add(paragraph('The module was created when I was looking for a Python support for MS Word .doc files on PyPI and Stackoverflow. Unfortunately, the only solutions I could find used:'))

    # Add a numbered list
    for point in ['''COM automation''','''.net or Java''','''Automating OpenOffice or MS Office''']:
        doc.add(paragraph(point,style='ListNumber'))
    doc.add(paragraph('''For those of us who prefer something simpler, I made docx.''')) 
        
    doc.add(heading('Making documents',2))
    doc.add(paragraph('''The docx module has the following features:'''))

    # Add some bullets
    for point in ['Paragraphs','Bullets','Numbered lists','Multiple levels of headings','Tables','Document Properties']:
        doc.add(paragraph(point,style='ListBullet'))

    doc.add(paragraph('Tables are just lists of lists, like this:'))
    # Append a table
    doc.add(table([['A1','A2','A3'],['B1','B2','B3'],['C1','C2','C3']]))

    doc.add(heading('Editing documents',2))
    doc.add(paragraph('Thanks to the awesomeness of the lxml module, we can:'))
    for point in ['Search and replace','Extract plain text of document','Add and delete items anywhere within the document']:
        doc.add(paragraph(point,style='ListBullet'))

    pic_paragraph = picture(doc,'python_logo.png','This is a test description')
    doc.add(pic_paragraph)

    doc.replace('the awesomeness','the goshdarned awesomeness') 

    # Add a pagebreak
    doc.add(pagebreak(type='page', orient='portrait'))

    doc.add(heading('Ideas? Questions? Want to contribute?',2))
    doc.add(paragraph('''Email <python.docx@librelist.com>'''))

    # Setting the meta properties of the document. This part is work in progress. If you create a document from scratch, you
    # we to set them manually. There are also WebSettings, AppProperties and ContentTypes. At the moment they are initialized
    # on creating a new DocxDocument. Take a look into meta.py for the various classes available. Plan is to work on those meta classes and expand their functionality to the user later on.
    doc.core_properties = CoreProperties(
                        title='Creating a docx document in Python from scratch.',
                        creator='markmywords',
                        keywords=['python', 'MS Office', 'docx']
                      )

    # Finally, we save the document.
    doc.save('Welcome to the Python docx module.docx')
