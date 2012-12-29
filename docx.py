#!/usr/bin/env python2.6
# -*- coding: utf-8 -*-
'''
Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and 'Office OpenXML' by Microsoft)

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
'''

from copy import deepcopy
import logging
from lxml import etree
try:
    from PIL import Image
except ImportError:
    import Image

import zipfile
import shutil
from distutils import dir_util
import re
import time
import os
from os.path import join, basename
from StringIO import StringIO

log = logging.getLogger(__name__)

# Record template directory's location which is just 'template' for a docx
# developer or 'site-packages/docx-template' if you have installed docx
TEMPLATE_DIR = join(os.path.dirname(__file__), 'docx-template') # installed
if not os.path.isdir(TEMPLATE_DIR):
    TEMPLATE_DIR = join(os.path.dirname(__file__), 'template') # dev

TEMP_TEMPLATE_DIR = join(os.path.dirname(__file__), '.temp_template_dir')

# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily.
nsprefixes = {
    # Text Content
    'mv': 'urn:schemas-microsoft-com:mac:vml',
    'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o': 'urn:schemas-microsoft-com:office:office',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    # Properties (core and extended)
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/',
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    # Content Types (we're just making up our own namespaces here to save time)
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships (we're just making up our own namespaces here to save time)
    'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
    }

image_relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
hlink_relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' 

def opendocx(file):
    '''Open a docx file, return a document XML tree'''
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml')
    document = etree.fromstring(xmlcontent)
    return document

def opendocxheader(file):
    '''Open a docx file, return the header XML tree'''
    mydoc = zipfile.ZipFile(file)
    try:
        xmlcontent = mydoc.read('word/header2.xml')
    except :
        xmlcontent = mydoc.read('word/header1.xml')
        
    header = etree.fromstring(xmlcontent)
    return header

def newdocument():
    build_temp_template_layout()

    document = makeelement('document')
    document.append(makeelement('body'))
    return document

def build_temp_template_layout():
    clear_temp_template_layout()
    dir_util.copy_tree(TEMPLATE_DIR, TEMP_TEMPLATE_DIR)

def clear_temp_template_layout():
    if os.path.exists(TEMP_TEMPLATE_DIR):
        if os.path.isdir(TEMP_TEMPLATE_DIR):
            shutil.rmtree(TEMP_TEMPLATE_DIR)
        else:
            os.remove(TEMP_TEMPLATE_DIR)

def makeelement(tagname,tagtext=None,nsprefix='w',attributes=None,attrnsprefix=None):
    '''Create an element & return it'''
    # Deal with list of nsprefix by making namespacemap
    namespacemap = None
    if isinstance(nsprefix, list):
        namespacemap = {}
        for prefix in nsprefix:
            namespacemap[prefix] = nsprefixes[prefix]
        nsprefix = nsprefix[0] # FIXME: rest of code below expects a single prefix
    if nsprefix:
        namespace = '{'+nsprefixes[nsprefix]+'}'
    else:
        # For when namespace = None
        namespace = ''
    newelement = etree.Element(namespace+tagname, nsmap=namespacemap)
    # Add attributes with namespaces
    if attributes:
        # If they haven't bothered setting attribute namespace, use an empty string
        # (equivalent of no namespace)
        if not attrnsprefix:
            # Quick hack: it seems every element that has a 'w' nsprefix for its tag uses the same prefix for it's attributes
            if nsprefix == 'w':
                attributenamespace = namespace
            else:
                attributenamespace = ''
        else:
            attributenamespace = '{'+nsprefixes[attrnsprefix]+'}'

        for tagattribute in attributes:
            newelement.set(attributenamespace+tagattribute, attributes[tagattribute])
    if tagtext:
        newelement.text = tagtext
    return newelement

def pagebreak(type='page', orient='portrait'):
    '''Insert a break, default 'page'.
    See http://openxmldeveloper.org/forums/thread/4075.aspx
    Return our page break element.'''
    # Need to enumerate different types of page breaks.
    validtypes = ['page', 'section']
    if type not in validtypes:
        raise ValueError('Page break style "%s" not implemented. Valid styles: %s.' % (type, validtypes))
    pagebreak = makeelement('p')
    if type == 'page':
        run = makeelement('r')
        br = makeelement('br',attributes={'type':type})
        run.append(br)
        pagebreak.append(run)
    elif type == 'section':
        pPr = makeelement('pPr')
        sectPr = makeelement('sectPr')
        if orient == 'portrait':
            pgSz = makeelement('pgSz',attributes={'w':'12240','h':'15840'})
        elif orient == 'landscape':
            pgSz = makeelement('pgSz',attributes={'h':'12240','w':'15840', 'orient':'landscape'})
        sectPr.append(pgSz)
        pPr.append(sectPr)
        pagebreak.append(pPr)
    return pagebreak

def paragraph(paratext,style='BodyText',breakbefore=False,jc='left',font='Times New Roman',fontsize=12):
    '''Make a new paragraph element, containing a run, and some text.
    Return the paragraph element.

    @param string jc: Paragraph alignment, possible values:
                      left, center, right, both (justified), ...
                      see http://www.schemacentral.com/sc/ooxml/t-w_ST_Jc.html
                      for a full list

    @param string font: Paragraph font family

    @param integer fontsize: Paragraph font size


    If paratext is a list, spawn multiple run/text elements.
    Support text styles (paratext must then be a list of lists in the form
    <text> / <style>. Style is a string containing a combination of 'bui' chars

    example
    paratext = [
        ('some bold text', {'style': 'b'}),
        ('some normal text', {'style': ''}),
        ('some italic underlined text', {'style': 'iu'}),
        ('some bold text with color and font size', {'style': 'b', 
                                                     'color': 'C00000', 
                                                     'sz': '32'})
    ]

    '''
    # Make our elements
    paragraph = makeelement('p')

    if not isinstance(paratext, list):
        paratext = [paratext]
    text = []
    for pt in paratext:
        if not isinstance(pt, (list,tuple)):
            pt = (pt, {})
        lines = pt[0].split('\n')
        for l in lines[:-1]:
            text.append((makeelement('t',tagtext=l), pt[1], True))  # with line break
        text.append((makeelement('t',tagtext=lines[-1]), pt[1], False))  # the last line, without line break        

    pPr = makeelement('pPr')
    pStyle = makeelement('pStyle',attributes={'val':style})
    pJc = makeelement('jc',attributes={'val':jc})
    pPr.append(pStyle)
    pPr.append(pJc)

    # Add the text to the run, and the run to the paragraph
    paragraph.append(pPr)
    for t in text:
        run = makeelement('r')
        rPr = makeelement('rPr')
        pFnt = makeelement('rFonts',attributes={'ascii':font,'cs':font,'eastAsia':font,'hAnsi':font})
        sz = makeelement('sz',attributes={'val':str(fontsize*2)})
        szCs = makeelement('szCs',attributes={'val':str(fontsize*2)})
        # Apply styles
        if t[1].has_key('style'):
            if t[1]['style'].find('b') > -1:
                b = makeelement('b')
                rPr.append(b)
            if t[1]['style'].find('u') > -1:
                u = makeelement('u',attributes={'val':'single'})
                rPr.append(u)
            if t[1]['style'].find('i') > -1:
                i = makeelement('i')
                rPr.append(i)
        for pr_key in t[1]:
            if not pr_key == 'style':
                pr = makeelement(pr_key,attributes={'val': t[1][pr_key]})
                rPr.append(pr)
        rPr.append(pFnt)
        rPr.append(sz)
        rPr.append(szCs)
        run.append(rPr)
        # Insert lastRenderedPageBreak for assistive technologies like
        # document narrators to know when a page break occurred.
        if breakbefore:
            lastRenderedPageBreak = makeelement('lastRenderedPageBreak')
            run.append(lastRenderedPageBreak)
        run.append(t[0])
        # Insert line break if there is multiple lines
        if t[2]:
            br = makeelement('br')
            run.append(br)
        paragraph.append(run)
    # Return the combined paragraph
    return paragraph

def getDefaultContentTypes():
    # FIXME - doesn't quite work...read from string as temp hack...
    #types = makeelement('Types',nsprefix='ct')
    types = etree.fromstring('''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>''')
    parts = {
        '/word/theme/theme1.xml':'application/vnd.openxmlformats-officedocument.theme+xml',
        '/word/fontTable.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
        '/docProps/core.xml':'application/vnd.openxmlformats-package.core-properties+xml',
        '/docProps/app.xml':'application/vnd.openxmlformats-officedocument.extended-properties+xml',
        '/word/document.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        '/word/settings.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
        '/word/numbering.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
        '/word/styles.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
        '/word/webSettings.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml'
        }
    for part in parts:
        types.append(makeelement('Override',nsprefix=None,attributes={'PartName':part,'ContentType':parts[part]}))
    # Add support for filetypes
    filetypes = {'rels':'application/vnd.openxmlformats-package.relationships+xml','xml':'application/xml','jpeg':'image/jpeg','gif':'image/gif','png':'image/png'}
    for extension in filetypes:
        types.append(makeelement('Default',nsprefix=None,attributes={'Extension':extension,'ContentType':filetypes[extension]}))
    return types

def getContentTypes(file=None):
    '''Get Content Types file from a given Word document'''
    if file and os.path.isfile(file):
        mydoc = zipfile.ZipFile(file)
        xmlcontent = mydoc.read('[Content_Types].xml')
        types = etree.fromstring(xmlcontent)
        return types
    else:
        return getDefaultContentTypes()

def heading(headingtext,headinglevel,lang='en'):
    '''Make a new heading, return the heading element'''
    lmap = {
        'en': 'Heading',
        'it': 'Titolo',
    }
    # Make our elements
    paragraph = makeelement('p')
    pr = makeelement('pPr')
    pStyle = makeelement('pStyle',attributes={'val':lmap[lang]+str(headinglevel)})
    run = makeelement('r')
    text = makeelement('t',tagtext=headingtext)
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)
    paragraph.append(run)
    # Return the combined paragraph
    return paragraph

def table(contents, tblstyle=None, tbllook={'val':'0400'}, heading=True, colw=None, cwunit='dxa', tblw=0, twunit='auto', borders={}, celstyle=None, rowstyle=None, table_props=None):
    '''Get a list of lists, return a table

        @param list contents: A list of lists describing contents
                              Every item in the list can be a string or a valid
                              XML element itself. It can also be a list. In that case
                              all the listed elements will be merged into the cell.
        @param string tblstyle: Specifies name of table style to override default if desired
        @param string tbllook: Specifies which elements of table style to to apply to this table,
                               e.g. {'firstColumn':'false', 'firstRow':'true'}, etc.
        @param bool heading: Tells whether first line should be treated as heading
                             or not
        @param list colw: A list of interger. The list must have same element
                          count of content lines. Specify column Widths in
                          wunitS
        @param string cwunit: Unit user for column width:
                                'pct': fifties of a percent
                                'dxa': twenties of a point
                                'nil': no width
                                'auto': automagically determined
        @param int tblw: Table width
        @param int twunit: Unit used for table width. Same as cwunit
        @param dict borders: Dictionary defining table border. Supported keys are:
                             'top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'all'
                             When specified, the 'all' key has precedence over others.
                             Each key must define a dict of border attributes:
                             color: The color of the border, in hex or 'auto'
                             space: The space, measured in points
                             sz: The size of the border, in eights of a point
                             val: The style of the border, see http://www.schemacentral.com/sc/ooxml/t-w_ST_Border.htm
        @param list celstyle: Specify the style for each colum, list of dicts.
                              supported keys:
                              'align': specify the alignment, see paragraph documentation,

        @return lxml.etree: Generated XML etree element
    '''
    table = makeelement('tbl')
    columns = len(contents[0])
    # Table properties
    tableprops = makeelement('tblPr')
    tablestyle = makeelement('tblStyle',attributes={'val':tblstyle if tblstyle else ''})
    tableprops.append(tablestyle)
    for attr in tableprops.iterchildren():
        if isinstance(attr, etree._Element):        
            tableprops.append(attr)            
        else:
            raise KeyError('what type of element to make?')
            prop = makeelement(k, attributes=attr)
            tableprops.append(prop)
            
    tablewidth = makeelement('tblW',attributes={'w':str(tblw),'type':str(twunit)})
    tableprops.append(tablewidth)
    if len(borders.keys()):
        tableborders = makeelement('tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            if b in borders.keys() or 'all' in borders.keys():
                k = 'all' if 'all' in borders.keys() else b
                attrs = {}
                for a in borders[k].keys():
                    attrs[a] = unicode(borders[k][a])
                borderelem = makeelement(b,attributes=attrs)
                tableborders.append(borderelem)
        tableprops.append(tableborders)
    tablelook = makeelement('tblLook',attributes=tbllook)
    tableprops.append(tablelook)
    table.append(tableprops)
    # Table Grid
    tablegrid = makeelement('tblGrid')
    for i in range(columns):
        tablegrid.append(makeelement('gridCol',attributes={'w':str(colw[i]) if colw else '2390'}))
    table.append(tablegrid)
    # Heading Row
    row = makeelement('tr')
    rowprops = makeelement('trPr')
    cnfStyle = makeelement('cnfStyle',attributes={'val':'000000100000'})
    rowprops.append(cnfStyle)
    row.append(rowprops)
    if heading:
        i = 0
        for heading in contents[0]:
            cell = makeelement('tc')
            # Cell properties
            cellprops = makeelement('tcPr')
            if colw:
                wattr = {'w':str(colw[i]),'type':cwunit}
            else:
                wattr = {'w':'0','type':'auto'}
            cellwidth = makeelement('tcW',attributes=wattr)
            cellstyle = makeelement('shd',attributes={'val':'clear','color':'auto','fill':'FFFFFF','themeFill':'text2','themeFillTint':'99'})
            cellprops.append(cellwidth)
            cellprops.append(cellstyle)
            cell.append(cellprops)
            # Paragraph (Content)
            if not isinstance(heading, (list, tuple)):
                heading = [heading,]
            for h in heading:
                if isinstance(h, etree._Element):
                    cell.append(h)
                else:
                    cell.append(paragraph(h,jc='center'))
            row.append(cell)
            i += 1
        table.append(row)
    # Contents Rows
    for contentrow in contents[1 if heading else 0:]:
        row = makeelement('tr')
        if rowstyle:
            rowprops = makeelement('trPr')
            if 'height' in rowstyle:
                rowHeight = makeelement('trHeight', attributes={'val': str(rowstyle['height']),
                                                                'hRule': 'exact'})
            rowprops.append(rowHeight)
            row.append(rowprops)
        i = 0
        for content_cell in contentrow:
            cell = makeelement('tc')
            # Properties
            cellprops = makeelement('tcPr')
            if colw:
                wattr = {'w':str(colw[i]),'type':cwunit}
            else:
                wattr = {'w':'0','type':'auto'}
            cellwidth = makeelement('tcW', attributes=wattr)
            cellprops.append(cellwidth)
            align = 'left'
            cell_spec_style = {}
            if celstyle:
                cell_spec_style = deepcopy(celstyle[i])
            if isinstance(content_cell, dict):
                cell_spec_style.update(content_cell['style'])                
                content_cell = content_cell['content']
            # spec. align property
            SPEC_PROPS = ['align',]
            if 'align' in cell_spec_style:
                align = celstyle[i]['align']
            # any property for cell, by OOXML specification
            for cs, attrs in cell_spec_style.iteritems():
                if cs in SPEC_PROPS:
                    continue
                cell_prop = makeelement(cs, attributes=attrs)
                cellprops.append(cell_prop)
            cell.append(cellprops)
            # Paragraph (Content)
            if not isinstance(content_cell, (list, tuple)):
                content_cell = [content_cell,]
            for c in content_cell:
                # cell.append(cellprops)
                if isinstance(c, etree._Element):
                    cell.append(c)
                else:
                    cell.append(paragraph(c, jc=align))
            row.append(cell)
            i += 1
        table.append(row)
    return table

def picture(relationshiplist, picpath, picdescription, pixelwidth=None,
            pixelheight=None, nochangeaspect=True, nochangearrowheads=True,
            template_dir=None):
    '''Take a relationshiplist, picture path, and return a paragraph containing the image
    and an updated relationshiplist'''
    # http://openxmldeveloper.org/articles/462.aspx
    # Create an image. Size may be specified, otherwise it will based on the
    # pixel size of image. Return a paragraph containing the picture'''
    # Copy the file into the media dir
    if not template_dir:
        template_dir = TEMP_TEMPLATE_DIR
    media_dir = join(template_dir, 'word', 'media')
    if not os.path.isdir(media_dir):
        os.makedirs(media_dir)
    shutil.copyfile(picpath, join(media_dir, picpath))

    # Check if the user has specified a size
    if not pixelwidth or not pixelheight:
        # If not, get info from the picture itself
        pixelwidth, pixelheight = Image.open(picpath).size[0:2]

    # OpenXML measures on-screen objects in English Metric Units
    # 1cm = 36000 EMUs
    emuperpixel = 12667
    width = str(pixelwidth * emuperpixel)
    height = str(pixelheight * emuperpixel)

    # Set relationship ID to the first available
    picid = '2'

    try:
        relid = (idx for idx, rel in enumerate(relationshiplist) if rel[1] == picpath).next() + 1

    except (StopIteration, IndexError):
        relationshiplist.append(makeelement('Relationship', attributes={'Id': 'rId' + str(len(relationshiplist)+1), 'Type': image_relationship,'Target': join('media', picpath)}, nsprefix=None))
        relid = len(relationshiplist)

    picrelid = 'rId' + str(relid)

    # There are 3 main elements inside a picture
    # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
    blipfill = makeelement('blipFill', nsprefix='pic')
    blipfill.append(makeelement('blip', nsprefix='a', attrnsprefix='r', attributes={'embed': picrelid}))
    stretch = makeelement('stretch', nsprefix='a')
    stretch.append(makeelement('fillRect', nsprefix='a'))
    blipfill.append(makeelement('srcRect', nsprefix='a'))
    blipfill.append(stretch)

    # 2. The non visual picture properties
    nvpicpr = makeelement('nvPicPr', nsprefix='pic')
    cnvpr = makeelement('cNvPr', nsprefix='pic',
                        attributes={'id': '0', 'name': 'Picture 1', 'descr': basename(picpath)})
    nvpicpr.append(cnvpr)
    cnvpicpr = makeelement('cNvPicPr', nsprefix='pic')
    cnvpicpr.append(makeelement('picLocks', nsprefix='a',
                    attributes={'noChangeAspect': str(int(nochangeaspect)),
                    'noChangeArrowheads': str(int(nochangearrowheads))}))
    nvpicpr.append(cnvpicpr)

    # 3. The Shape properties
    sppr = makeelement('spPr', nsprefix='pic', attributes={'bwMode': 'auto'})
    xfrm = makeelement('xfrm' ,nsprefix='a')
    xfrm.append(makeelement('off', nsprefix='a', attributes={'x': '0','y': '0'}))
    xfrm.append(makeelement('ext', nsprefix='a', attributes={'cx': width,'cy': height}))
    prstgeom = makeelement('prstGeom', nsprefix='a', attributes={'prst': 'rect'})
    prstgeom.append(makeelement('avLst', nsprefix='a'))
    sppr.append(xfrm)
    sppr.append(prstgeom)

    # Add our 3 parts to the picture element
    pic = makeelement('pic', nsprefix='pic')
    pic.append(nvpicpr)
    pic.append(blipfill)
    pic.append(sppr)

    # Now make the supporting elements
    # The following sequence is just: make element, then add its children
    graphicdata = makeelement('graphicData', nsprefix='a',
                              attributes={'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'})
    graphicdata.append(pic)
    graphic = makeelement('graphic', nsprefix='a')
    graphic.append(graphicdata)

    framelocks = makeelement('graphicFrameLocks', nsprefix='a', attributes={'noChangeAspect':'1'})
    framepr = makeelement('cNvGraphicFramePr', nsprefix='wp')
    framepr.append(framelocks)
    docpr = makeelement('docPr', nsprefix='wp',
                        attributes={'id': picid, 'name': 'Picture 1', 'descr':picdescription})
    effectextent = makeelement('effectExtent', nsprefix='wp',
                               attributes={'l': '25400', 't': '0', 'r': '0', 'b': '0'})
    extent = makeelement('extent', nsprefix='wp', attributes={'cx': width,'cy': height})
    inline = makeelement('inline',
                         attributes={'distT': "0", 'distB': "0", 'distL': "0", 'distR': "0"},
                         nsprefix='wp')
    inline.append(extent)
    inline.append(effectextent)
    inline.append(docpr)
    inline.append(framepr)
    inline.append(graphic)
    drawing = makeelement('drawing')
    drawing.append(inline)
    run = makeelement('r')
    run.append(drawing)
    paragraph = makeelement('p')
    paragraph.append(run)
    return relationshiplist, paragraph


def search(document,search):
    '''Search a document for a regex, return success / fail result'''
    result = False
    searchre = re.compile(search)
    for element in document.iter():
        if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
            if element.text:
                if searchre.search(element.text):
                    result = True
    return result

def findElementByText(document,search):
    '''Search a document for a regex, return the first matching text element'''
    res_element = None
    searchre = re.compile(search)
    for element in document.iter():
        if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
            if element.text:
                if searchre.search(element.text):
                    res_element = element
                    return res_element
    return res_element

def replace(document,search,replace):
    '''Replace all occurences of string with a different string, return updated document'''
    newdocument = document
    searchre = re.compile(search)
    for element in newdocument.iter():
        if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
            if element.text:
                if searchre.search(element.text):
                    element.text = re.sub(search,replace,element.text)
    return newdocument

def clean(document):
    """ Perform misc cleaning operations on documents.
        Returns cleaned document.
    """

    newdocument = document

    # Clean empty text and r tags
    for t in ('t', 'r'):
        rmlist = []
        for element in newdocument.iter():
            if element.tag == '{%s}%s' % (nsprefixes['w'], t):
                if not element.text and not len(element):
                    rmlist.append(element)
        for element in rmlist:
            element.getparent().remove(element)

    return newdocument

def findTypeParent(element, tag):
    """ Finds fist parent of element of the given type
    
    @param object element: etree element
    @param string the tag parent to search for
    
    @return object element: the found parent or None when not found
    """
    
    p = element
    while True:
        p = p.getparent()
        if p.tag == tag:
            return p
    
    # Not found
    return None

def advSearch(document, search, bs=3):
    '''Return set of all regex matches

    This is an advanced version of python-docx.search() that takes into
    account blocks of <bs> elements at a time.

    What it does:
    It searches the entire document body for text blocks.
    Since the text to search could be spawned across multiple text blocks,
    we need to adopt some sort of algorithm to handle this situation.
    The smaller matching group of blocks (up to bs) is then adopted.
    If the matching group has more than one block, blocks other than first
    are cleared and all the replacement text is put on first block.

    Examples:
    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search : 'Hello,'
    output blocks : [ 'Hello,' ]

    original text blocks : [ 'Hel', 'lo', ' __', 'name', '__!' ]
    search : '(__[a-z]+__)'
    output blocks : [ '__name__' ]

    @param instance  document: The original document
    @param str       search: The text to search for (regexp)
                          append, or a list of etree elements
    @param int       bs: See above

    @return set      All occurences of search string

    '''

    # Compile the search regexp
    searchre = re.compile(search)

    matches = []

    # Will match against searchels. Searchels is a list that contains last
    # n text elements found in the document. 1 < n < bs
    searchels = []

    for element in document.iter():
        if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
            if element.text:
                # Add this element to searchels
                searchels.append(element)
                if len(searchels) > bs:
                    # Is searchels is too long, remove first elements
                    searchels.pop(0)

                found = False
                txtsearch = ''
                for el in searchels:
                    txtsearch += el.text

                # Searcs for the text in the whole txtsearch
                match = searchre.search(txtsearch)
                if match:
                    # I've found something :)
                    # The end of match must be in the last element
                    # |----|---s|omet|hing|!---|
                    # Now we need to find the element containing the start of the match
                    # Search all combinations, of searchels, starting from
                    # smaller up to bigger ones
                    # s = search start
                    # e = element IDs to merge
                    els_len = len(searchels)
                    for s in reversed(range(els_len)):
                        txtsearch = ''
                        e = range(s,els_len)
                        for k in e:
                            txtsearch += searchels[k].text
                        match = searchre.search(txtsearch)
                        if match:
                            found = True
                            matches.append(match.group())
                    # Now we have found the start and ending element of the match
                    # |  0 |  1 |  2 |  3 |  4 |
                    # |----|---s|omet|hing|!---|
                    # e = [1,2,3,4]
    return set(matches)


def advFindElementByText(document, search, bs=3):
    '''
    This is an advanced version of python-docx.findElementByText() that takes into
    account blocks of <bs> elements at a time.

    @param instance  document: The original document
    @param str       search: The text to search for (regexp)
                          append, or a list of etree elements
    @param int       bs: See above

    @return list      The element set contains the search when combined
    '''

    # Compile the search regexp
    searchre = re.compile(search)

    # Will match against searchels. Searchels is a list that contains last
    # n text elements found in the document. 1 < n < bs
    searchels = []

    for element in document.iter():
        if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
            if element.text:
                # Add this element to searchels
                searchels.append(element)
                if len(searchels) > bs:
                    # Is searchels is too long, remove first elements
                    searchels.pop(0)

                found = False
                txtsearch = ''
                for el in searchels:
                    txtsearch += el.text

                # Searcs for the text in the whole txtsearch
                match = searchre.search(txtsearch)
                if match:
                    # I've found something :)
                    # The end of match must be in the last element
                    # |----|---s|omet|hing|!---|
                    # Now we need to find the element containing the start of the match
                    # Search all combinations, of searchels, starting from
                    # smaller up to bigger ones
                    # s = search start
                    # e = element IDs to merge
                    els_len = len(searchels)
                    for s in reversed(range(els_len)):
                        txtsearch = ''
                        e = range(s,els_len)
                        for k in e:
                            txtsearch += searchels[k].text
                        match = searchre.search(txtsearch)
                        if match:
                            found = True
                            return searchels[s:]
                    # Now we have found the start and ending element of the match
                    # |  0 |  1 |  2 |  3 |  4 |
                    # |----|---s|omet|hing|!---|
                    # e = [1,2,3,4]
    return None


def advReplace(document,search,replace,bs=3):
    '''Replace all occurences of string with a different string, return updated document

    This is a modified version of python-docx.replace() that takes into
    account blocks of <bs> elements at a time. The replace element can only
    be a string currently. (There is some bug with element replacing)

    What it does:
    It searches the entire document body for text blocks.
    Then scan thos text blocks for replace.
    Since the text to search could be spawned across multiple text blocks,
    we need to adopt some sort of algorithm to handle this situation.
    The smaller matching group of blocks (up to bs) is then adopted.
    If the matching group has more than one block, blocks other than first
    are cleared and all the replacement text is put on first block.

    Examples:
    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hello,' / 'Hi!'
    output blocks : [ 'Hi!', '', ' world!' ]

    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hello, world' / 'Hi!'
    output blocks : [ 'Hi!!', '', '' ]

    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hel' / 'Hal'
    output blocks : [ 'Hal', 'lo,', ' world!' ]

    @param instance  document: The original document
    @param str       search: The text to search for (regexp)
    @param mixed replace: The replacement text or lxml.etree element to
                          append, or a list of etree elements
    @param int       bs: See above

    @return instance The document with replacement applied

    '''
    # Enables debug output
    DEBUG = False

    newdocument = document

    # Compile the search regexp
    searchre = re.compile(search)

    # Will match against searchels. Searchels is a list that contains last
    # n text elements found in the document. 1 < n < bs
    searchels = []

    for element in newdocument.iter():
        if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
            if element.text:
                # Add this element to searchels
                searchels.append(element)
                if len(searchels) > bs:
                    # Is searchels is too long, remove first elements
                    searchels.pop(0)

                found = False
                txtsearch = ''
                for el in searchels:
                    txtsearch += el.text

                # Searcs for the text in the whole txtsearch
                match = searchre.search(txtsearch)
                if match:
                    # I've found something :)
                    # The end of match must be in the last element
                    # |----|---s|omet|hing|!---|
                    # Now we need to find the element containing the start of the match
                    # Search all combinations, of searchels, starting from
                    # smaller up to bigger ones
                    # s = search start
                    # e = element IDs to merge
                    els_len = len(searchels)
                    for s in reversed(range(els_len)):
                        txtsearch = ''
                        e = range(s,els_len)
                        for k in e:
                            txtsearch += searchels[k].text
                        match = searchre.search(txtsearch)
                        if match:
                            found = True
                            break
                    # Now we have found the start and ending element of the match
                    # |  0 |  1 |  2 |  3 |  4 |
                    # |----|---s|omet|hing|!---|
                    # e = [1,2,3,4]
                    if DEBUG:
                        log.debug("Found element!")
                        log.debug("Search regexp: %s", searchre.pattern)
                        log.debug("Requested replacement: %s", replace)
                        log.debug("Matched text: %s", txtsearch)
                        log.debug( "Matched text (splitted): %s", map(lambda i:i.text,searchels))
                        log.debug("Matched at position: %s", match.start())
                        log.debug( "matched in elements: %s", e)
                        if isinstance(replace, etree._Element):
                            log.debug("Will replace with XML CODE")
                        elif isinstance(replace (list, tuple)):
                            log.debug("Will replace with LIST OF ELEMENTS")
                        else:
                            log.debug("Will replace with:", re.sub(search,replace,txtsearch))

                    try:
                        # for text, we want to replace at <w:r> level
                        # because there will be new lines
                        # <w:p>
                        #     <w:pPr></w:pPr>
                        #     <w:r>
                        #         <w:rPr></w:rPr>
                        #         <w:t>[PREFIX]</w:t>
                        #     </w:r>
                        #     <w:r>
                        #         <w:rPr></w:rPr>
                        #         <w:t>[REPLACE_LINE_1]</w:t>
                        #         <w:br/>
                        #     </w:r>
                        #     <w:r>
                        #         <w:rPr></w:rPr>
                        #         <w:t>[REPLACE_LINE_2]</w:t>
                        #     </w:r>
                        #     <w:r>
                        #         <w:rPr></w:rPr>
                        #         <w:t>[SUFIX]</w:t>
                        #     </w:r>
                        # </w:p>
                        txt_prefix = txtsearch[:match.start()]
                        txt_sufix = txtsearch[match.end():]
                        if not isinstance(replace, (list, tuple)):
                            replace = [replace]
                        # get the first matching <w:r> element and copy it
                        # because we want to keep its property <w:rPr> for the replaced content
                        org_first_r = findTypeParent(searchels[e[0]], '{%s}r' % nsprefixes['w'])
                        first_r = deepcopy(org_first_r)
                        # and remove <w:br> if there is any
                        # because these <w:br>s will be added accordingly later
                        el_to_delete = first_r.find('{%s}br' % nsprefixes['w'])
                        if not el_to_delete == None:
                            first_r.remove(el_to_delete)
                        template_r = deepcopy(first_r)
                        # empty the text
                        el_to_delete = template_r.find('{%s}t' % nsprefixes['w'])
                        if not el_to_delete == None:
                            template_r.remove(el_to_delete)
                        p =  org_first_r.getparent()
                        insindex = p.index(org_first_r)

                        # insert the prefix
                        p.insert(insindex, first_r)
                        insindex += 1
                        first_r.find('{%s}t' % nsprefixes['w']).text = txt_prefix
                        # insert the runs to replace
                        for r in replace:
                            if isinstance(r, etree._Element):
                                # TODO: there was some bug in element replacement
                                # delete the function currently
                                # but replacing some text with some picture is really useful
                                # will add element replacing later
                                continue
                            lines = r.split('\n')
                            for line in lines[:-1]:
                                new_r = deepcopy(template_r)
                                new_r.append(makeelement('t',tagtext=line))
                                new_r.append(makeelement('br'))
                                p.insert(insindex, new_r)
                                insindex += 1
                            # the last line of the segment
                            new_r = deepcopy(template_r)
                            new_r.append(makeelement('t',tagtext=lines[-1]))
                            p.insert(insindex, new_r)
                            insindex += 1
                        # copy the last matching <w:r> element
                        # because it may be the same as the first <w:r> element
                        last_r = deepcopy(findTypeParent(searchels[e[-1]], '{%s}r' % nsprefixes['w']))
                        p.insert(insindex, last_r)
                        insindex += 1
                        last_r.find('{%s}t' % nsprefixes['w']).text = txt_sufix

                        # there may be 1 matching element, or multiple elements combined
                        # anyway, remove them all since we have safely copied the prefix
                        # and sufix and inserted replace in between
                        for i in e:
                            # maybe the element cannot be removed
                            # so, anyway, empty them first in case I cannot remove them
                            searchels[i].text = ''
                            try:
                                p.remove(findTypeParent(searchels[i], '{%s}r' % nsprefixes['w']))
                            except Exception, e:
                                log.exception(e)

                        return newdocument
                    except Extension, e:
                        log.exception(e)
    return newdocument

def getdocumenttext(document):
    '''Return the raw text of a document, as a list of paragraphs.'''
    paratextlist=[]
    # Compile a list of all paragraph (p) elements
    paralist = []
    for element in document.iter():
        # Find p (paragraph) elements
        if element.tag == '{'+nsprefixes['w']+'}p':
            paralist.append(element)
    # Since a single sentence might be spread over multiple text elements, iterate through each
    # paragraph, appending all text (t) children to that paragraphs text.
    for para in paralist:
        paratext=u''
        # Loop through each paragraph
        for element in para.iter():
            # Find t (text) elements
            if element.tag == '{'+nsprefixes['w']+'}t':
                if element.text:
                    paratext = paratext+element.text
            elif element.tag == '{'+nsprefixes['w']+'}tab':
                paratext = paratext + '\t'
        # Add our completed paragraph text to the list of paragraph text
        if not len(paratext) == 0:
            paratextlist.append(paratext)
    return paratextlist

def coreproperties(title,subject,creator,keywords,lastmodifiedby=None):
    '''Create core properties (common document properties referred to in the 'Dublin Core' specification).
    See appproperties() for other stuff.'''
    coreprops = makeelement('coreProperties',nsprefix='cp')
    coreprops.append(makeelement('title',tagtext=title,nsprefix='dc'))
    coreprops.append(makeelement('subject',tagtext=subject,nsprefix='dc'))
    coreprops.append(makeelement('creator',tagtext=creator,nsprefix='dc'))
    coreprops.append(makeelement('keywords',tagtext=','.join(keywords),nsprefix='cp'))
    if not lastmodifiedby:
        lastmodifiedby = creator
    coreprops.append(makeelement('lastModifiedBy',tagtext=lastmodifiedby,nsprefix='cp'))
    coreprops.append(makeelement('revision',tagtext='1',nsprefix='cp'))
    coreprops.append(makeelement('category',tagtext='Examples',nsprefix='cp'))
    coreprops.append(makeelement('description',tagtext='Examples',nsprefix='dc'))
    currenttime = time.strftime('%Y-%m-%dT%H:%M:%SZ')
    # Document creation and modify times
    # Prob here: we have an attribute who name uses one namespace, and that
    # attribute's value uses another namespace.
    # We're creating the lement from a string as a workaround...
    for doctime in ['created','modified']:
        coreprops.append(etree.fromstring('''<dcterms:'''+doctime+''' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">'''+currenttime+'''</dcterms:'''+doctime+'''>'''))
        pass
    return coreprops

def appproperties():
    '''Create app-specific properties. See docproperties() for more common document properties.'''
    appprops = makeelement('Properties',nsprefix='ep')
    appprops = etree.fromstring(
    '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>''')
    props = {
            'Template':'Normal.dotm',
            'TotalTime':'6',
            'Pages':'1',
            'Words':'83',
            'Characters':'475',
            'Application':'Microsoft Word 12.0.0',
            'DocSecurity':'0',
            'Lines':'12',
            'Paragraphs':'8',
            'ScaleCrop':'false',
            'LinksUpToDate':'false',
            'CharactersWithSpaces':'583',
            'SharedDoc':'false',
            'HyperlinksChanged':'false',
            'AppVersion':'12.0000',
            }
    for prop in props:
        appprops.append(makeelement(prop,tagtext=props[prop],nsprefix=None))
    return appprops

def websettings():
    '''Generate websettings'''
    web = makeelement('webSettings')
    web.append(makeelement('allowPNG'))
    web.append(makeelement('doNotSaveAsSingleFile'))
    return web

def getDefaultRelationships():
    '''Generate a default Word relationships file'''
    # Default list of relationships
    relationshiplist = [
        ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering','numbering.xml'],
        ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles','styles.xml'],
        ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml'],
        ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings','webSettings.xml'],
        ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable','fontTable.xml'],
        ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme','theme/theme1.xml'],
    ]
    relationships = etree.fromstring(
    '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>'''
    )

    for idx, relationship in enumerate(relationshiplist):
        # Relationship IDs (rId) start at 1.
        element = makeelement('Relationship', attributes={'Id': 'rId' + str(idx + 1),
                              'Type': relationship[0],'Target': relationship[1]}, nsprefix=None)
        if relationship[0] == hlink_relationship:
            element.attrib['TargetMode'] = 'External'
        relationships.append(element)
    
    return relationships

def getRelationships(file=None):
    '''Get a Word relationships file from a given Word document'''
    if file and os.path.isfile(file):
        mydoc = zipfile.ZipFile(file)
        xmlcontent = mydoc.read('word/_rels/document.xml.rels')
        relationships = etree.fromstring(xmlcontent)
        return relationships
    else:
        return getDefaultRelationships()

def savedocx(document, coreprops, appprops, contenttypes, websettings, wordrelationships, output,
             template_dir=None):
    '''Save a modified document'''
    if not template_dir:
        template_dir = TEMP_TEMPLATE_DIR
    assert os.path.isdir(template_dir)
    docxfile = zipfile.ZipFile(output, mode='w', compression=zipfile.ZIP_DEFLATED)

    # save images referred in relationshiplist, adjust relationshiplist
#    _relationshiplist = []
#    for r_type, target in relationshiplist:
#        if r_type == image_relationship:
#            path = 'media/' + basename(target)
#            docxfile.write(target, 'word/' + path)
#            _relationshiplist += [[r_type, path]]
#        else:
#            _relationshiplist += [[r_type, target]]

    # Move to the template data path
    prev_dir = os.path.abspath('.') # save previous working dir
    os.chdir(template_dir)
    
#    _wordrelationships = wordrelationships(_relationshiplist)
    
    # Add page settings
    document = add_page_settings(document)

    # Serialize our trees into out zip file
    treesandfiles = {document: 'word/document.xml',
                     coreprops: 'docProps/core.xml',
                     appprops: 'docProps/app.xml',
                     contenttypes: '[Content_Types].xml',
                     websettings: 'word/webSettings.xml',
                     wordrelationships: 'word/_rels/document.xml.rels'}
    
    for tree in treesandfiles:
        log.info('Saving: ' + treesandfiles[tree])
        treestring = etree.tostring(tree, pretty_print=True)
        docxfile.writestr(treesandfiles[tree], treestring)

    # Add & compress support files
    files_to_ignore = ['.DS_Store'] # nuisance from some os's
    for dirpath, dirnames, filenames in os.walk('.'):
        for filename in filenames:
            if filename in files_to_ignore:
                continue
            templatefile = join(dirpath, filename)
            archivename = templatefile[2:]
            log.info('Saving: %s', archivename)
            docxfile.write(templatefile, archivename)
            
    log.info('Saved new file to: %r', output)
    docxfile.close()
    os.chdir(prev_dir) # restore previous working dir

    return

def djangodocx(document,coreprops,appprops,contenttypes,websettings,wordrelationships,output=None):
    '''Save a modified document'''
    template_dir	= TEMP_TEMPLATE_DIR
    assert os.path.isdir(template_dir)
    
    if not output:
        output = StringIO()
        
    docxfile = zipfile.ZipFile(output,mode='w',compression=zipfile.ZIP_DEFLATED)
    
    # Move to the template data path
    prev_dir = os.path.abspath('.') # save previous working dir
    os.chdir(template_dir)

    # Serialize our trees into out zip file
    treesandfiles = {document:'word/document.xml',
                     coreprops:'docProps/core.xml',
                     appprops:'docProps/app.xml',
                     contenttypes:'[Content_Types].xml',
                     websettings:'word/webSettings.xml',
                     wordrelationships:'word/_rels/document.xml.rels'}
    for tree in treesandfiles:
        log.info('Saving: '+treesandfiles[tree]    )
        treestring = etree.tostring(tree, pretty_print=True)
        docxfile.writestr(treesandfiles[tree],treestring)

    # Add & compress support files
    files_to_ignore = ['.DS_Store'] # nuisance from some os's
    for dirpath,dirnames,filenames in os.walk('.'):
        for filename in filenames:
            if filename in files_to_ignore:
                continue
            templatefile = join(dirpath,filename)
            archivename = templatefile[2:]
            log.info('Saving: %s', archivename)
            docxfile.write(templatefile, archivename)
    docxfile.close()
    docx_zip = output.getvalue()
    output.close()
    os.chdir(prev_dir) # restore previous working dir
    
    return docx_zip


def add_page_settings(document):
    ''' Add custom page settings '''
    sectPr = makeelement('sectPr')
    pgMar = makeelement('pgMar', attributes={'bottom': '720', 'footer': '0', 'gutter': '0', 'header': '0',
                                             'left': '1138', 'right': '1138', 'top': '1138'})
    type = makeelement('type', attributes={'val': 'nextPage'})
    pgSz = makeelement('pgSz', attributes={'h': '15840', 'w': '12240'})
    pgNumType = makeelement('pgNumType', attributes={'fmt': 'decimal'})
    formProt = makeelement('formProt', attributes={'val': 'false'})
    textDirection = makeelement('textDirection', attributes={'val': 'lrTb'})
    docGrid = makeelement('docGrid', attributes={'charSpace': '0', 'linePitch': '240', 'type': 'default'})

    sectPr.append(type)
    sectPr.append(pgSz)
    sectPr.append(pgMar)
    sectPr.append(pgNumType)
    sectPr.append(formProt)
    sectPr.append(textDirection)
    sectPr.append(docGrid)

    docbody = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    docbody.append(sectPr)

    return document
