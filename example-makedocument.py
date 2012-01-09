#!/usr/bin/env python2.6
'''
This file makes an docx (Office 2007) file from scratch, showing off most of python-docx's features.

If you need to make documents from scratch, use this file as a basis for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
'''
from docx.dsl import *

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

    h2('Editing documents')
    
    p('Thanks to the awesomeness of the lxml module, we can:')
    
    with ul() as li:
        li('Search and replace')
        li('Extract plain text of document')
        li('Add and delete items anywhere within the document')
        
    """
    # Add an image
    relationships,picpara = picture(relationships,'image1.png','This is a test description')
    docbody.append(picpara)
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
    br()

    h2('Ideas? Questions? Want to contribute?')
    p('Email <python.docx@librelist.com>')    
    
    # Save our document
    write_docx('Welcome to the Python docx module.docx')