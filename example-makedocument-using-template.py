#!/usr/bin/env python2.6
'''
This file makes an docx (Office 2007) file from scratch, showing off most of python-docx's features.

If you need to make documents from scratch, use this file as a basis for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
'''
import tempfile
from docx import *

def template_setup(filepath):
    z = zipfile.ZipFile(filepath, 'r')
    template_dir = tempfile.mkdtemp(prefix='docx-')
    z.extractall(template_dir)
    set_template(template_dir)

if __name__ == '__main__':        
<<<<<<< HEAD:example-makedocument-using-template.py
    # setup template files
    template_setup('template.docx')

    # Default set of relationshipships - these are the minimum components of a document
    relationships = relationshiplist()
=======
>>>>>>> a908b4bfef36fd5de73df8ac8b543a34d24498a4:example-makedocument.py

    # Make a new document tree - this is the main part of a Word document
    doc = newdocx(title='Python docx demo',subject='A practical example of making docx from Python',creator='Mike MacCana',keywords=['python','Office Open XML','Word'])
    
    # This xpath location is where most interesting content lives 
    
    # Append two headings and a paragraph
    doc = append(doc, heading('''Welcome to Python's docx module''',1)  )   
    doc = append(doc, heading('Make and edit docx in 200 lines of pure Python',2))
    doc = append(doc, paragraph('The module was created when I was looking for a Python support for MS Word .doc files on PyPI and Stackoverflow. Unfortunately, the only solutions I could find used:'))

    # Add a numbered list
    for point in ['''COM automation''','''.net or Java''','''Automating OpenOffice or MS Office''']:
<<<<<<< HEAD:example-makedocument-using-template.py
        docbody.append(paragraph(point,style='List Number'))
    docbody.append(paragraph('''For those of us who prefer something simpler, I made docx.''')) 
=======
        doc = append(doc, paragraph(point,style='ListNumber'))
    doc = append(doc, paragraph('''For those of us who prefer something simpler, I made docx.''')) 
>>>>>>> a908b4bfef36fd5de73df8ac8b543a34d24498a4:example-makedocument.py
    
    doc = append(doc, heading('Making documents',2))
    doc = append(doc, paragraph('''The docx module has the following features:'''))

    # Add some bullets
    for point in ['Paragraphs','Bullets','Numbered lists','Multiple levels of headings','Tables','Document Properties']:
<<<<<<< HEAD:example-makedocument-using-template.py
        docbody.append(paragraph(point,style='List Bullet'))
=======
        doc = append(doc, paragraph(point,style='ListBullet'))
>>>>>>> a908b4bfef36fd5de73df8ac8b543a34d24498a4:example-makedocument.py

    doc = append(doc, paragraph('Tables are just lists of lists, like this:'))
    # Append a table
    doc = append(doc, table([['A1','A2','A3'],['B1','B2','B3'],['C1','C2','C3']]))

    doc = append(doc, heading('Editing documents',2))
    doc = append(doc, paragraph('Thanks to the awesomeness of the lxml module, we can:'))
    for point in ['Search and replace','Extract plain text of document','Add and delete items anywhere within the document']:
<<<<<<< HEAD:example-makedocument-using-template.py
        docbody.append(paragraph(point,style='List Bullet'))
=======
        doc = append(doc, paragraph(point,style='ListBullet'))
>>>>>>> a908b4bfef36fd5de73df8ac8b543a34d24498a4:example-makedocument.py
        
    # Add an image
    doc = addpicture(doc,'image1.png','This is a test description')
 
    # Search and replace
    print 'Searching for something in a paragraph ...',
    if searchdocx(doc, 'the awesomeness'): print 'found it!'
    else: print 'nope.'
    
    print 'Searching for something in a heading ...',
    if searchdocx(doc, '200 lines'): print 'found it!'
    else: print 'nope.'
    
    print 'Replacing ...',
    doc = replacedocx(doc,'the awesomeness','the goshdarned awesomeness') 
    print 'done.'

    # Add a pagebreak
    doc = append(doc, pagebreak(type='page', orient='portrait'))

    doc = append(doc, heading('Ideas? Questions? Want to contribute?',2))
    doc = append(doc, paragraph('''Email <python.docx@librelist.com>'''))

    # Save our document
<<<<<<< HEAD:example-makedocument-using-template.py
    savedocx(document,coreprops,appprops,contenttypes,websettings,wordrelationships,'Welcome to the Python docx module.docx')
=======
    savedocx(doc,'Welcome to the Python docx module.docx')
>>>>>>> a908b4bfef36fd5de73df8ac8b543a34d24498a4:example-makedocument.py
