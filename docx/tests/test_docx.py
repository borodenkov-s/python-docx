#!/usr/bin/env python2.6
"""
Test docx module
<<<<<<< HEAD:docx/tests/test_docx.py
"""
import os
import lxml
from docx.document import *
=======
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
from docx.utils import *


# --- Some fixture and testing values ---
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1:tests/test_docx.py

TEST_FILE = 'ShortTest.docx'
IMAGE1_FILE = 'image1.png'

RELATIONS = """
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Id="rId1" Target="numbering.xml"/>
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Id="rId2" Target="styles.xml"/>
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Id="rId3" Target="settings.xml"/>
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Id="rId4" Target="webSettings.xml"/>
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Id="rId5" Target="fontTable.xml"/>
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Id="rId6" Target="theme/theme1.xml"/>
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Id="rId7" Target="header1.xml"/>
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Id="rId8" Target="footer.xml"/>
    </Relationships>
    """

# --- Setup & Support Functions ---
def setup_module():
    """Set up test fixtures"""
    import shutil
    if IMAGE1_FILE not in os.listdir('.'):
        shutil.copyfile(join(os.path.pardir,IMAGE1_FILE), IMAGE1_FILE)
    testnewdocument()

def teardown_module():
    """Tear down test fixtures"""
    if TEST_FILE in os.listdir('.'):
        os.remove(TEST_FILE)

def simpledoc():
<<<<<<< HEAD:docx/tests/test_docx.py
    """Make a docx (document, relationships) for use in other docx tests"""
    relationships = relationshiplist()
    document = newdocument()
    docbody = document.xpath('/w:document/w:body', namespaces=NSPREFIXES)[0]
    docbody.append(heading('Heading 1',1)  )   
    docbody.append(heading('Heading 2',2))
    docbody.append(paragraph('Paragraph 1'))
    for point in ['List Item 1','List Item 2','List Item 3']:
        docbody.append(paragraph(point,style='ListNumber'))
    docbody.append(pagebreak(typeOfBreak='page'))
    docbody.append(paragraph('Paragraph 2')) 
    docbody.append(table([['A1','A2','A3'],['B1','B2','B3'],['C1','C2','C3']]))
    docbody.append(pagebreak(typeOfBreak='section', orient='portrait'))
    relationships,picpara = picture(relationships,IMAGE1_FILE,'This is a test description')
    docbody.append(picpara)
    docbody.append(pagebreak(typeOfBreak='section', orient='landscape'))
    docbody.append(paragraph('Paragraph 3'))
    return (document, docbody, relationships)
=======
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

    picpara = picture(document, IMAGE1_FILE,
    'This is a test description')
    document.append(picpara)
    document.append(pagebreak(breaktype='section', orient='landscape'))
    document.append(paragraph('Paragraph 3'))
    return document
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1:tests/test_docx.py


# --- Test Functions ---
def testsearchandreplace():
<<<<<<< HEAD:docx/tests/test_docx.py
    """Ensure search and replace functions work"""
    document, docbody, relationships = simpledoc()
    docbody = document.xpath('/w:document/w:body', namespaces=NSPREFIXES)[0]
    assert search(docbody, 'ing 1')
    assert search(docbody, 'ing 2')
    assert search(docbody, 'graph 3')
    assert search(docbody, 'ist Item')
    assert search(docbody, 'A1')
    if search(docbody, 'Paragraph 2'): 
        docbody = replace(docbody,'Paragraph 2','Whacko 55') 
    assert search(docbody, 'Whacko 55')
    
def testtextextraction():
    """Ensure text can be pulled out of a document"""
    document = opendocx(TEST_FILE)
    paratextlist = getdocumenttext(document)
    assert len(paratextlist) > 0

def testunsupportedpagebreak():
    """Ensure unsupported page break types are trapped"""
    document = newdocument()
    docbody = document.xpath('/w:document/w:body', namespaces=NSPREFIXES)[0]
    try:
        docbody.append(pagebreak(typeOfBreak='unsup'))
=======
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
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1:tests/test_docx.py
    except ValueError:
        return # passed
    assert False # failed

def testnewdocument():
<<<<<<< HEAD:docx/tests/test_docx.py
    """Test that a new document can be created"""
    document, docbody, relationships = simpledoc()
    coreprops = coreproperties('Python docx testnewdocument','A short example of making docx from Python','Alan Brooks',['python','Office Open XML','Word'])
    savedocx(document, coreprops, appproperties(), contenttypes(), websettings(), wordrelationships(relationships), TEST_FILE)

def testopendocx():
    """Ensure an etree element is returned"""
    if isinstance(opendocx(TEST_FILE),lxml.etree._Element):
        pass
    else:
        assert False
=======
    '''Test that a new document can be created'''
    doc = simpledoc()
    doc.title = 'Python docx testnewdocument'
    doc.description = 'A short example of making docx from Python'
    doc.author = 'Alan Brooks'
    doc.save(TEST_FILE)

def testopendocx():
    '''Ensure an etree element is returned'''
    assert isinstance(Docx(TEST_FILE).document,lxml.etree._Element)
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1:tests/test_docx.py

def testmakeelement():
    """Ensure custom elements get created"""
    testelement = makeelement('testname',attributes={'testattribute':'testvalue'},tagtext='testtagtext')
    assert testelement.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testname'
    assert testelement.attrib == {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testattribute': 'testvalue'}
    assert testelement.text == 'testtagtext'

def testparagraph():
    """Ensure paragraph creates p elements"""
    testpara = paragraph('paratext',style='BodyText')
    assert testpara.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    pass
def testparagraphwithtab():
    '''Ensure paragraph creates the correct xml tree when passing content \t'''
    """<ns0:p xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <ns0:pPr>
            <ns0:pStyle ns0:val="BodyText" />
            <ns0:jc ns0:val="left" />
          </ns0:pPr>
          <ns0:r>
            <ns0:rPr>
              <ns0:rFonts ns0:ascii="Times New Roman" ns0:hAnsi="Times New Roman" ns0:eastAsia="Times New Roman" ns0:cs="Times New Roman" />
              <ns0:sz ns0:val="24" />
              <ns0:szCs ns0:val="24" />
            </ns0:rPr>
            <ns0:t>TitreA</ns0:t>
          </ns0:r>
          <ns0:r>
            <ns0:rPr>
              <ns0:rFonts ns0:ascii="Times New Roman" ns0:hAnsi="Times New Roman" ns0:eastAsia="Times New Roman" ns0:cs="Times New Roman" />
              <ns0:sz ns0:val="24" />
              <ns0:szCs ns0:val="24" />
            </ns0:rPr>
            <ns0:tab />
          </ns0:r>
          <ns0:r>
            <ns0:rPr>
              <ns0:rFonts ns0:ascii="Times New Roman" ns0:hAnsi="Times New Roman" ns0:eastAsia="Times New Roman" ns0:cs="Times New Roman" />
              <ns0:sz ns0:val="24" />
              <ns0:szCs ns0:val="24" />
            </ns0:rPr>
            <ns0:t>text after tab</ns0:t>
          </ns0:r>
        </ns0:p>
    """
    testpara = paragraph('TitreA\ttext after tab', style='BodyText',)

    assert testpara.getchildren()[1].getchildren()[1].text == 'TitreA'
    assert testpara.getchildren()[2].getchildren()[1].tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab'
    assert testpara.getchildren()[3].getchildren()[1].text == 'text after tab'



def testtable():
    """Ensure tables make sense"""
    testtable = table([['A1','A2'],['B1','B2'],['C1','C2']])
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    assert testtable.xpath('/ns0:tbl/ns0:tr[2]/ns0:tc[2]/ns0:p/ns0:r/ns0:t',namespaces={'ns0':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})[0].text == 'B2'

def testnew_id():
    ''' Ensure we can get unique id for relationship'''
    relationshiplist = lxml.etree.fromstring(RELATIONS)
    assert new_id(relationshiplist) == 'rId9'

if __name__=='__main__':
    import nose
    nose.main()
