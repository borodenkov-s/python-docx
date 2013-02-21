#!/usr/bin/env python2.6
'''
This file makes an docx (Office 2007) file from scratch, showing off most of python-docx's features.

If you need to make documents from scratch, use this file as a basis for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
<<<<<<< HEAD
'''
<<<<<<< HEAD
<<<<<<< HEAD
from docx.dsl import *
=======
"""

from docx.document import *

if __name__ == '__main__':
    # Default set of relationshipships - the minimum components of a document
    relationships = relationshiplist()

    # Make a new document tree - this is the main part of a Word document
    document = newdocument()

    # This xpath location is where most interesting content lives
    body = document.xpath('/w:document/w:body', namespaces=NSPREFIXES)[0]
>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c

if __name__ == '__main__':        
    # Make a new document
    start_doc(meta={
        'title': 'Python docx demo',
        'subject': 'A practical example of making docx from Python',
        'creator': 'Mike MacCana',
        'keywords': ['python','Office Open XML','Word']
    })
    
    # Append two headings and a paragraph
    h1("Welcome to Python's docx module")   
    h2('Make and edit docx in 200 lines of pure Python')
    
    p('The module was created when I was looking for a Python support for MS Word .doc files on PyPI and Stackoverflow. Unfortunately, the only solutions I could find used:')

    # Add a numbered list
<<<<<<< HEAD
    with ol() as li:
        li('COM automation')
        li('.net or Java')
        li('Automating OpenOffice or MS Office')
        
    p('For those of us who prefer something simpler, I made docx.') 
    
    h2('Making documents')
    
    p('The docx module has the following features:')
=======
    points = ['COM automation',
              '.net or Java',
              'Automating OpenOffice or MS Office'
             ]
    for point in points:
        body.append(paragraph(point, style='ListNumber'))
    body.append(paragraph('For those of us who prefer something simpler, I '
                          'made docx.'))
    body.append(heading('Making documents', 2))
    body.append(paragraph('The docx module has the following features:'))
>>>>>>> 7bb35b0877215421b39515eba9d27f9829b6bc14

    # Add some bullets
    with ul() as li:
        li("Paragraphs")
        li("Bullets")
        li("Numbered lists")
        li('Multiple levels of headings')
        li('Tables')
        li('Document Properties')

    p('Tables are just lists of lists, like this:')
    
    # Append a table
<<<<<<< HEAD
    with table() as tr:
        with tr() as td:
            td("A1")
            td("A2")
            td("A3")
        with tr() as td:
            td("B1")
            td("B2")
            td("B3")
        with tr() as td:
            td("C1")
            td("C2")
            td("C3")

    h2('Editing documents')
    
    p('Thanks to the awesomeness of the lxml module, we can:')
    
    with ul() as li:
        li('Search and replace')
        li('Extract plain text of document')
        li('Add and delete items anywhere within the document')
        
    """
=======
    tbl_rows = [['A1', 'A2', 'A3'],
                ['B1', 'B2', 'B3'],
                ['C1', 'C2', 'C3']]
    body.append(table(tbl_rows))

    body.append(heading('Editing documents', 2))
    body.append(paragraph('Thanks to the awesomeness of the lxml module, '
                          'we can:'))
    points = ['Search and replace',
             'Extract plain text of document',
             'Add and delete items anywhere within the document'
             ]
    for point in points:
        body.append(paragraph(point, style='ListBullet'))

>>>>>>> 7bb35b0877215421b39515eba9d27f9829b6bc14
=======
import shutil
from copy import deepcopy
from tempfile import TemporaryFile, mkdtemp
import docx as dx
=======
from docx.dsl import *
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1

if __name__ == '__main__':        
    # Make a new document
    start_doc(meta={
        'title': 'Python docx demo',
        'subject': 'A practical example of making docx from Python',
        'creator': 'Mike MacCana',
        'keywords': ['python','Office Open XML','Word']
    })
    
    # Append two headings and a paragraph
    h1("Welcome to Python's docx module")   
    h2('Make and edit docx in 200 lines of pure Python')
    
    p('The module was created when I was looking for a Python support for MS Word .doc files on PyPI and Stackoverflow. Unfortunately, the only solutions I could find used:')

    # Add a numbered list
    with ol() as li:
        li('COM automation')
        li('.net or Java')
        li('Automating OpenOffice or MS Office')
        
    p('For those of us who prefer something simpler, I made docx.') 
    
    h2('Making documents')
    
    p('The docx module has the following features:')

    # Add some bullets
    with ul() as li:
        li("Paragraphs")
        li("Bullets")
        li("Numbered lists")
        li('Multiple levels of headings')
        li('Tables')
        li('Document Properties')

    p('Tables are just lists of lists, like this:')
    
    # Append a table
    with table() as tr:
        with tr() as td:
            td("A1")
            td("A2")
            td("A3")
        with tr() as td:
            td("B1")
            td("B2")
            td("B3")
        with tr() as td:
            td("C1")
            td("C2")
            td("C3")

<<<<<<< HEAD
>>>>>>> c2c09b66b47efe1922d5dd4f03e52eec0a06ad15
=======
    h2('Editing documents')
    
    p('Thanks to the awesomeness of the lxml module, we can:')
    
    with ul() as li:
        li('Search and replace')
        li('Extract plain text of document')
        li('Add and delete items anywhere within the document')
        
    """
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    # Add an image
    relationships,picpara = picture(relationships,'image1.png','This is a test description')
    docbody.append(picpara)
<<<<<<< HEAD
<<<<<<< HEAD
    """
    
    # Search and replace
    print 'Searching for something in a paragraph ...',
    if doc.search('the awesomeness'): 
        print 'found it!'
    else: 
        print 'nope.'

    print 'Searching for something in a heading ...',
    if doc.search('200 lines'): 
        print 'found it!'
    else: 
        print 'nope.'
    
    print 'Replacing ...',
    doc.replace('the awesomeness','the goshdarned awesomeness')
    print 'done.'

    # Add a pagebreak
<<<<<<< HEAD
    br()
=======
    body.append(pagebreak(type='page', orient='portrait'))

    body.append(heading('Ideas? Questions? Want to contribute?', 2))
    body.append(paragraph('Email <python.docx@librelist.com>'))

    # Create our properties, contenttypes, and other support files
    title = 'Python docx demo'
    subject = 'A practical example of making docx from Python'
    creator = 'Mike MacCana'
    keywords = ['python', 'Office Open XML', 'Word']

    coreprops = coreproperties(title=title, subject=subject, creator=creator,
                               keywords=keywords)
    appprops = appproperties()
    contenttypes = contenttypes()
    websettings = websettings()
    wordrelationships = wordrelationships(relationships)
>>>>>>> 7bb35b0877215421b39515eba9d27f9829b6bc14

    h2('Ideas? Questions? Want to contribute?')
    p('Email <python.docx@librelist.com>')    
    
    # Save our document
<<<<<<< HEAD
    write_docx('Welcome to the Python docx module.docx')
=======
    savedocx(document, coreprops, appprops, contenttypes, websettings,
             wordrelationships, 'Welcome to the Python docx module.docx')
>>>>>>> 7bb35b0877215421b39515eba9d27f9829b6bc14
=======

    docbody.append(dx.paragraph([
        ('hello', {}),
        ('2', {'vertAlign': 'superscript'}),
        ]))

    # Append a table with special properties and cells
    spec_cell = dx.paragraph([('2', {'vertAlign': 'superscript'})])
    t_prop_margin = dx.makeelement('tblCellMar')
    for margin_type in ['top', 'left', 'right', 'bottom']:
        t_prop_margin.append(dx.makeelement(margin_type, attributes={'w': '0', 'type': 'dxa'}))
    CELL_SIZE = 12*30 # twenties of a point
    docbody.append(dx.table([['A1',
                              {'content': spec_cell, 'style': {'vAlign': {'val': 'top'},
                                                               'shd': {'fill': '777777'}}},
                              ('A3', 'ttt')],
                             ['B1','B2','B3'],
                             ['C1','C2','C3']],
        heading=False,
        colw=[CELL_SIZE]*3,
        cwunit='dxa', # twenties of a point
        borders={'all': {'color': 'AAAAAA'}},
        celstyle=[{'align': 'center', 'vAlign': {'val': 'center'}}]*3,
        rowstyle={'height': CELL_SIZE},
        table_props={'jc': {'val': 'center'},
                     '__margin__': t_prop_margin,
                     },
    ))
=======
    """
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    
    # Search and replace
    print 'Searching for something in a paragraph ...',
    if doc.search('the awesomeness'): 
        print 'found it!'
    else: 
        print 'nope.'

    print 'Searching for something in a heading ...',
    if doc.search('200 lines'): 
        print 'found it!'
    else: 
        print 'nope.'
    
    print 'Replacing ...',
    doc.replace('the awesomeness','the goshdarned awesomeness')
    print 'done.'

    # Add a pagebreak
    br()

    h2('Ideas? Questions? Want to contribute?')
    p('Email <python.docx@librelist.com>')    
    
    # Save our document
<<<<<<< HEAD
    dx.savedocx(document,coreprops,appprops,my_contenttypes,my_websettings,relationships,'Welcome to the Python docx module.docx')
<<<<<<< HEAD
>>>>>>> c2c09b66b47efe1922d5dd4f03e52eec0a06ad15
=======
>>>>>>> f3612b9ebedbb122f611b6c6e8cb3f13c4a46481
=======
    write_docx('Welcome to the Python docx module.docx')
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
