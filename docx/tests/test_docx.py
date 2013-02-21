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
<<<<<<< HEAD:docx/tests/test_docx.py
        shutil.copyfile(os.path.join(os.path.pardir,IMAGE1_FILE), IMAGE1_FILE)
    testsavedocument()
=======
        shutil.copyfile(join(os.path.pardir,IMAGE1_FILE), IMAGE1_FILE)
    testnewdocument()
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f:tests/test_docx.py

def teardown_module():
    '''Tear down test fixtures'''
    if TEST_FILE in os.listdir('.'):
        os.remove(TEST_FILE)

def simpledoc():
    '''Make a docx (document, relationships) for use in other docx tests'''
<<<<<<< HEAD:docx/tests/test_docx.py
    doc = newdocx('Python docx testnewdocument','A short example of making docx from Python','Alan Brooks',['python','Office Open XML','Word'])

    document = getdocument(doc)
    relationships = getrelationshiplist(doc)

    docbody = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    docbody.append(heading('Heading 1',1)  )
    docbody.append(heading('Heading 2',2))
    docbody.append(paragraph('Paragraph 1'))
    for point in ['List Item 1','List Item 2','List Item 3']:
        docbody.append(paragraph(point,style='ListNumber'))
    docbody.append(pagebreak(type='page'))
    docbody.append(paragraph('Paragraph 2'))
    docbody.append(table([['A1','A2','A3'],['B1','B2','B3'],['C1','C2','C3']]))
    docbody.append(pagebreak(type='section', orient='portrait'))
    relationships,picpara = picture(relationships,IMAGE1_FILE,'This is a test description')
    docbody.append(picpara)
    docbody.append(pagebreak(type='section', orient='landscape'))
    docbody.append(paragraph('Paragraph 3'))

    doc['word/document.xml'] = document
    doc['word/_rels/document.xml.rels'] = wordrelationships(relationships)
    return doc
=======
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
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f:tests/test_docx.py


# --- Test Functions ---
def testsearchandreplace():
    '''Ensure search and replace functions work'''
<<<<<<< HEAD:docx/tests/test_docx.py
    doc = simpledoc()
    document = getdocument(doc)
    docbody = getdocbody(document)
    assert search(docbody, 'ing 1')
    assert search(docbody, 'ing 2')
    assert search(docbody, 'graph 3')
    assert search(docbody, 'ist Item')
    assert search(docbody, 'A1')
    if search(docbody, 'Paragraph 2'):
        docbody = replace(docbody,'Paragraph 2','Whacko 55')
    assert search(docbody, 'Whacko 55')

=======
    document = Docx(TEST_FILE)
    assert document.search('ing 1')
    assert document.search('ing 2')
    assert document.search('graph 3')
    assert document.search('ist Item')
    assert document.search('A1')
    if document.search('Paragraph 2'): 
        document.replace('Paragraph 2','Whacko 55') 
    assert document.search('Whacko 55')
    
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f:tests/test_docx.py
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
<<<<<<< HEAD:docx/tests/test_docx.py

def testsavedocument():
    '''Tests a new document can be saved'''
    document = simpledoc()
    savedocx(document, TEST_FILE)

def testgetdocument():
    '''Ensure an etree element is returned'''
    doc = opendocx(TEST_FILE)
    document = getdocument(doc)
    if isinstance(document,lxml.etree._Element):
        pass
    else:
        assert False
=======
    
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
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f:tests/test_docx.py

def testmakeelement():
    '''Ensure custom elements get created'''
    testelement = Element('testname',attributes={'testattribute':'testvalue'},tagtext='testtagtext')
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
