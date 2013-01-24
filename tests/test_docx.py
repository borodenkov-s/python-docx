#!/usr/bin/env python2.6
'''
Test docx module
'''

# Set up the path
from os.path import abspath, join, dirname
import sys
PROJECT_ROOT = abspath(join(dirname(__file__), ".."))
sys.path.insert(0, PROJECT_ROOT)

# do the imports for testing
import os
import lxml
from docx import *

TEST_FILE = 'ShortTest.docx'
IMAGE1_FILE = 'image1.png'

# --- Setup & Support Functions ---
def setup_module():
    '''Set up test fixtures'''
    import shutil
    if IMAGE1_FILE not in os.listdir('.'):
        shutil.copyfile(join(os.path.pardir,IMAGE1_FILE), IMAGE1_FILE)
    testnewdocument()

def teardown_module():
    '''Tear down test fixtures'''
    if TEST_FILE in os.listdir('.'):
        os.remove(TEST_FILE)

def simpledoc():
    '''Make a docx (document, relationships) for use in other docx tests'''
    document = Docx()
    document.append(heading('Heading 1',1)  )   
    document.append(heading('Heading 2',2))
    document.append(paragraph('Paragraph 1'))
    for point in ['List Item 1','List Item 2','List Item 3']:
        document.append(paragraph(point,style='ListNumber'))
    document.append(pagebreak(breaktype='page'))
    document.append(paragraph('Paragraph 2')) 
    document.append(table([['A1','A2','A3'],['B1','B2','B3'],['C1','C2','C3']]))
    document.append(pagebreak(breaktype='section', orient='portrait'))
    
    #relationships,picpara = picture(relationships,IMAGE1_FILE,'This is a test description')
    #document.append(picpara)
    document.append(pagebreak(breaktype='section', orient='landscape'))
    document.append(paragraph('Paragraph 3'))
    return document


# --- Test Functions ---
def testsearchandreplace():
    '''Ensure search and replace functions work'''
    document = Docx(TEST_FILE)
    assert document.search('ing 1')
    assert document.search('ing 2')
    assert document.search('graph 3')
    assert document.search('ist Item')
    assert document.search('A1')
    if document.search('Paragraph 2'): 
        document.replace('Paragraph 2','Whacko 55') 
    assert document.search('Whacko 55')
    
def testtextextraction():
    '''Ensure text can be pulled out of a document'''
    document = Docx(TEST_FILE)
    paratextlist = document.text
    assert len(paratextlist) > 0

def testunsupportedpagebreak():
    '''Ensure unsupported page break types are trapped'''
    document = Docx().document
    docbody = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    try:
        docbody.append(pagebreak(breaktype='unsup'))
    except ValueError:
        return # passed
    assert False # failed
    
def testnewdocument():
    '''Test that a new document can be created'''
    doc = simpledoc()
    doc.title = 'Python docx testnewdocument'
    doc.description = 'A short example of making docx from Python'
    doc.author = 'Alan Brooks'
    doc.save(TEST_FILE)

def testopendocx():
    '''Ensure an etree element is returned'''
    assert isinstance(Docx(TEST_FILE).document,lxml.etree._Element)

def testmakeelement():
    '''Ensure custom elements get created'''
    testelement = makeelement('testname',attributes={'testattribute':'testvalue'},tagtext='testtagtext')
    assert testelement.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testname'
    assert testelement.attrib == {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testattribute': 'testvalue'}
    assert testelement.text == 'testtagtext'

def testparagraph():
    '''Ensure paragraph creates p elements'''
    testpara = paragraph('paratext',style='BodyText')
    assert testpara.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    pass
    
def testtable():
    '''Ensure tables make sense'''
    testtable = table([['A1','A2'],['B1','B2'],['C1','C2']])
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    assert testtable.xpath('/ns0:tbl/ns0:tr[2]/ns0:tc[2]/ns0:p/ns0:r/ns0:t',namespaces={'ns0':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})[0].text == 'B2'

if __name__=='__main__':
    import nose
    nose.main()