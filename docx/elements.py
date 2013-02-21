<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
import os

from PIL import Image
from lxml import etree

from utils import make_element


def pagebreak(type='page', orient='portrait'):
=======
# -*- coding: utf-8 -*-
import os
from os.path import join
import shutil
=======
#-*- coding: utf-8 -*-

from os.path import join, basename
from copy import deepcopy
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1

from lxml import etree
try:
    from PIL import Image
except ImportError:
    import Image
<<<<<<< HEAD
    
from metadata import TEMPLATE_DIR, nsprefixes

def Element(tagname, tagtext=None, nsprefix='w', attributes=None,attrnsprefix=None):
    '''Create an element & return it'''
=======

from .metadata import nsprefixes, FORMAT, image_relationship

from .utils import new_id

import logging
log = logging.getLogger(__name__)


def makeelement(tagname, tagtext=None, nsprefix='w', attributes=None,
                                                            attrnsprefix=None):
    '''Create an element & return it

    @param dict attributes:
            first key is the attribut's name,
            second could be the value or
                a dict like {'value': val, 'prefix': prefix}
            to set per attr prefix

    @param string attrnsprefix: Attribut default prefix

    '''
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    # Deal with list of nsprefix by making namespacemap
    namespacemap = None
    if isinstance(nsprefix, list):
        namespacemap = {}
        for prefix in nsprefix:
            namespacemap[prefix] = nsprefixes[prefix]
<<<<<<< HEAD
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
=======
        nsprefix = nsprefix[0]  # FIXME: code below expects a single prefix
    if nsprefix:
        namespace = '{' + nsprefixes[nsprefix] + '}'
    else:
        # For when namespace = None
        namespace = ''

    newelement = etree.Element(namespace + tagname, nsmap=namespacemap)
    # Add attributes with namespaces
    if attributes:
        # If they haven't bothered setting attribute namespace,
        #  use an empty string
        # (equivalent of no namespace)
        if not attrnsprefix:
            # Quick hack: it seems every element that has a 'w' nsprefix
            # for its tag uses the same prefix for it's attributes
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
            if nsprefix == 'w':
                attributenamespace = namespace
            else:
                attributenamespace = ''
        else:
<<<<<<< HEAD
            attributenamespace = '{'+nsprefixes[attrnsprefix]+'}'

        for tagattribute in attributes:
            newelement.set(attributenamespace+tagattribute, attributes[tagattribute])
=======
            attributenamespace = '{' + nsprefixes[attrnsprefix] + '}'

        if isinstance(attributes, dict):
            for attrname, attrvalue in attributes.iteritems():
                if isinstance(attrvalue, dict):
                    ns = '{' + nsprefixes[attrvalue['prefix']] + '}'
                    newelement.set(ns + attrname, attrvalue['value'])
                else:
                    newelement.set(attributenamespace + attrname, attrvalue)
        else:
            newelement.set(attributenamespace + attributes, attributes)

>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    if tagtext:
        newelement.text = tagtext
    return newelement


<<<<<<< HEAD
def pagebreak(breaktype='page', orient='portrait'):
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
    '''Insert a break, default 'page'.
    See http://openxmldeveloper.org/forums/thread/4075.aspx
    Return our page break element.'''
    # Need to enumerate different types of page breaks.
    validtypes = ['page', 'section']
<<<<<<< HEAD
    if type not in validtypes:
        raise ValueError('Page break style "%s" not implemented. Valid styles: %s.' % (type, validtypes))
    pagebreak = make_element('p')
    if type == 'page':
        run = make_element('r')
        br = make_element('br',attributes={'type':type})
        run.append(br)
        pagebreak.append(run)
    elif type == 'section':
        pPr = make_element('pPr')
        sectPr = make_element('sectPr')
        if orient == 'portrait':
            pgSz = make_element('pgSz',attributes={'w':'12240','h':'15840'})
        elif orient == 'landscape':
            pgSz = make_element('pgSz',attributes={'h':'12240','w':'15840', 'orient':'landscape'})
=======
    if breaktype not in validtypes:
        raise ValueError('Page break style "%s" not implemented. Valid styles: %s.' % (breaktype, validtypes))
    pagebreak = Element('p')
    if breaktype == 'page':
        run = Element('r')
        br = Element('br',attributes={'type':breaktype})
        run.append(br)
        pagebreak.append(run)
    elif breaktype == 'section':
        pPr = Element('pPr')
        sectPr = Element('sectPr')
        if orient == 'portrait':
            pgSz = Element('pgSz',attributes={'w':'12240','h':'15840'})
        elif orient == 'landscape':
            pgSz = Element('pgSz',attributes={'h':'12240','w':'15840', 'orient':'landscape'})
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
=======
def pagebreak(breaktype='page', orient='portrait', pageformat='letter'):
    '''Insert a break, default 'page'.
    See http: //openxmldeveloper.org/forums/thread/4075.aspx
    Return our page break element.'''
    # Need to enumerate different types of page breaks.
    validtypes = ['page', 'section']
    if breaktype not in validtypes:
        raise ValueError('Page break style "%s" not implemented.\
                                Valid styles: %s.' % (breaktype, validtypes))
    pagebreak = makeelement('p')
    if breaktype == 'page':
        run = makeelement('r')
        br = makeelement('br', attributes={'type': breaktype})
        run.append(br)
        pagebreak.append(run)
    elif breaktype == 'section':
        pPr = makeelement('pPr')
        sectPr = makeelement('sectPr')

        pageSize = FORMAT[pageformat]
        if orient == 'landscape':
            pageSize['orient'] = 'landscape'

        pgSz = makeelement('pgSz', attributes=pageSize)
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
        sectPr.append(pgSz)
        pPr.append(sectPr)
        pagebreak.append(pPr)
    return pagebreak

<<<<<<< HEAD
<<<<<<< HEAD

def paragraph(paratext,style='BodyText',breakbefore=False,jc='left'):
    '''Make a new paragraph element, containing a run, and some text. 
    Return the paragraph element.
    
=======
def paragraph(paratext,style='BodyText',breakbefore=False,jc='left'):
    '''Make a new paragraph element, containing a run, and some text.
    Return the paragraph element.

>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
    @param string jc: Paragraph alignment, possible values:
                      left, center, right, both (justified), ...
                      see http://www.schemacentral.com/sc/ooxml/t-w_ST_Jc.html
                      for a full list
<<<<<<< HEAD
    
    If paratext is a list, spawn multiple run/text elements.
    Support text styles (paratext must then be a list of lists in the form
    <text> / <style>. Stile is a string containing a combination od 'bui' chars
    
    example
    paratext = [
        ['some bold text', 'b'],
        ['some normal text', ''],
        ['some italic underlined text', 'iu'],
    ]
    
    '''
    # Make our elements
    paragraph = make_element('p')
    
    if type(paratext) == list:
        text = []
        for pt in paratext:
            if type(pt) == list:
                text.append([make_element('t',tagtext=pt[0]), pt[1]])
            else:
                text.append([make_element('t',tagtext=pt), ''])
    else:
        text = [[make_element('t',tagtext=paratext),''],]
    pPr = make_element('pPr')
    pStyle = make_element('pStyle',attributes={'val':style})
    pJc = make_element('jc',attributes={'val':jc})
    pPr.append(pStyle)
    pPr.append(pJc)
                
    # Add the text the run, and the run to the paragraph
    paragraph.append(pPr)
    for t in text:
        run = make_element('r')
        rPr = make_element('rPr')
        # Apply styles
        if t[1].find('b') > -1:
            b = make_element('b')
            rPr.append(b)
        if t[1].find('u') > -1:
            u = make_element('u',attributes={'val':'single'})
            rPr.append(u)
        if t[1].find('i') > -1:
            i = make_element('i')
=======

    If paratext is a list, spawn multiple run/text elements.
    Support text styles (paratext must then be a list of lists in the form
    <text> / <style>. Stile is a string containing a combination od 'bui' chars

    example
    paratext = [
        ('some bold text', 'b'),
        ('some normal text', ''),
        ('some italic underlined text', 'iu'),
=======

def paragraph(paratext, style='BodyText', breakbefore=False, jc='left',
                                        font='Times New Roman', fontsize=12,
                                        tabs=None):
    '''Make a new paragraph element, containing a run, and some text.
    Return the paragraph element.

    @param string jc: Paragraph alignment, possible values:
                      left, center, right, both (justified), ...
                      see http: //www.schemacentral.com/sc/ooxml/t-w_ST_Jc.html
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
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    ]

    '''
    # Make our elements
<<<<<<< HEAD
    paragraph = Element('p')

    if isinstance(paratext, list):
        text = []
        for pt in paratext:
            if isinstance(pt, (list,tuple)):
                text.append([Element('t',tagtext=pt[0]), pt[1]])
            else:
                text.append([Element('t',tagtext=pt), ''])
    else:
        text = [[Element('t',tagtext=paratext),''],]
    pPr = Element('pPr')
    pStyle = Element('pStyle',attributes={'val':style})
    pJc = Element('jc',attributes={'val':jc})
    pPr.append(pStyle)
    pPr.append(pJc)

    # Add the text the run, and the run to the paragraph
    paragraph.append(pPr)
    for t in text:
        run = Element('r')
        rPr = Element('rPr')
        # Apply styles
        if t[1].find('b') > -1:
            b = Element('b')
            rPr.append(b)
        if t[1].find('u') > -1:
            u = Element('u',attributes={'val':'single'})
            rPr.append(u)
        if t[1].find('i') > -1:
            i = Element('i')
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
            rPr.append(i)
=======
    paragraph = makeelement('p')


    def maketext(text):
        out = []
        if '\t' not in text:
            return [makeelement('t', tagtext=text),]
        #with tab node
        for part in text.partition('\t'):
            if part == '\t':
                out.append(makeelement('tab'))
            else:
                out.append(makeelement('t', tagtext=part))
        return out


    if not isinstance(paratext, list):
        paratext = [paratext]
    text = []

    for pt in paratext:
        if not isinstance(pt, (list, tuple)):
            pt = (pt, {})
        lines = pt[0].split('\n')
        for l in lines[:-1]:
            # with line break
            for el in maketext(l)[:-1]:
                text.append((el, pt[1], False))
            text.append((maketext(l)[-1], pt[1], True))
        # the last line, without line break
        for el in maketext(lines[-1]):
            text.append((el, pt[1], False))

    pPr = makeelement('pPr')
    pStyle = makeelement('pStyle', attributes={'val': style})
    pJc = makeelement('jc', attributes={'val': jc})

    if tabs:
        pTabs = makeelement('tabs')
        for tab in tabs:
            pTabs.append(makeelement('tab', attributes=tab))

        pPr.append(pTabs)

    pPr.append(pStyle)
    pPr.append(pJc)

    # Add the text to the run, and the run to the paragraph
    paragraph.append(pPr)
    for t in text:
        run = makeelement('r')
        rPr = makeelement('rPr')
        pFnt = makeelement('rFonts', attributes={'ascii': font, 'cs': font,
                                            'eastAsia': font, 'hAnsi': font})
        sz = makeelement('sz', attributes={'val': str(fontsize * 2)})
        szCs = makeelement('szCs', attributes={'val': str(fontsize * 2)})
        # Apply styles
        if 'style' in t[1]:
            if t[1]['style'].find('b') > -1:
                b = makeelement('b')
                rPr.append(b)
            if t[1]['style'].find('u') > -1:
                u = makeelement('u', attributes={'val': 'single'})
                rPr.append(u)
            if t[1]['style'].find('i') > -1:
                i = makeelement('i')
                rPr.append(i)
        for pr_key in t[1]:
            if not pr_key == 'style':
                pr = makeelement(pr_key, attributes={'val': t[1][pr_key]})
                rPr.append(pr)
        rPr.append(pFnt)
        rPr.append(sz)
        rPr.append(szCs)
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
        run.append(rPr)
        # Insert lastRenderedPageBreak for assistive technologies like
        # document narrators to know when a page break occurred.
        if breakbefore:
<<<<<<< HEAD
<<<<<<< HEAD
            lastRenderedPageBreak = make_element('lastRenderedPageBreak')
=======
            lastRenderedPageBreak = Element('lastRenderedPageBreak')
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
            run.append(lastRenderedPageBreak)
        run.append(t[0])
=======
            lastRenderedPageBreak = makeelement('lastRenderedPageBreak')
            run.append(lastRenderedPageBreak)
        run.append(t[0])
        # Insert line break if there is multiple lines
        if t[2]:
            br = makeelement('br')
            run.append(br)
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
        paragraph.append(run)
    # Return the combined paragraph
    return paragraph

<<<<<<< HEAD
<<<<<<< HEAD

=======
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
def heading(headingtext,headinglevel,lang='en'):
=======

def heading(headingtext, headinglevel, lang='en'):
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    '''Make a new heading, return the heading element'''
    lmap = {
        'en': 'Heading',
        'it': 'Titolo',
<<<<<<< HEAD
    }
    # Make our elements
<<<<<<< HEAD
    paragraph = make_element('p')
    pr = make_element('pPr')
    pStyle = make_element('pStyle',attributes={'val':lmap[lang]+str(headinglevel)})    
    run = make_element('r')
    text = make_element('t',tagtext=headingtext)
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)   
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph


def table(contents, heading=True, colw=None, cwunit='dxa', tblw=0, twunit='auto', borders={}, celstyle=None):
    '''Get a list of lists, return a table
=======
# encoding: utf-8

class Element(object):
    def __init__(tagname, tagtext=None, nsprefix='w', attributes=None, attrnsprefix=None):
        '''Create an element & return it''' 
        # Deal with list of nsprefix by making namespacemap
        namespacemap = None
        if type(nsprefix) == list:
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
              
        self.xml = newelement

    def append(self, element):
        self.xml.append(element)

class PageBreak(Element):
    def __init__(type='page', orientation='portrait'):
        '''Insert a break, default 'page'.
        See http://openxmldeveloper.org/forums/thread/4075.aspx'''
        # Need to enumerate different types of page breaks.
        validtypes = ['page', 'section']
        if type not in validtypes:
            raise ValueError('Page break style "%s" not implemented. Valid styles: %s.' % (type, validtypes))
        pagebreak = Element('p')
        if type == 'page':
            run = Element('r')
            br = Element('br',attributes={'type':type})
            run.append(br)
            pagebreak.append(run)
        elif type == 'section':
            pPr = Element('pPr')
            sectPr = Element('sectPr')
            if orientation == 'portrait':
                pgSz = Element('pgSz',attributes={'w':'12240','h':'15840'})
            elif orientation == 'landscape':
                pgSz = Element('pgSz',attributes={'h':'12240','w':'15840', 'orient':'landscape'})
            sectPr.append(pgSz)
            pPr.append(sectPr)
            pagebreak.append(pPr)
        
        self.type = type
        self.orientation = orientation
        self.xml = pagebreak

class Paragraph(Element):
    def __init__(self, paratext,style='BodyText',breakbefore=False,jc='left'):
        '''Make a new paragraph element, containing a run, and some text. 
        Return the paragraph element.
        
        @param string jc: Paragraph alignment, possible values:
                          left, center, right, both (justified), ...
                          see http://www.schemacentral.com/sc/ooxml/t-w_ST_Jc.html
                          for a full list
        
        If paratext is a list, spawn multiple run/text elements.
        Support text styles (paratext must then be a list of lists in the form
        <text> / <style>. Stile is a string containing a combination od 'bui' chars
        
        example
        paratext = [
            ['some bold text', 'b'],
            ['some normal text', ''],
            ['some italic underlined text', 'iu'],
        ]
        
        '''
        # Make our elements
        paragraph = Element('p')
        
        if type(paratext) == list:
            text = []
            for pt in paratext:
                if type(pt) == list:
                    text.append([Element('t',tagtext=pt[0]), pt[1]])
                else:
                    text.append([Element('t',tagtext=pt), ''])
        else:
            text = [[Element('t',tagtext=paratext),''],]
        pPr = Element('pPr')
        pStyle = Element('pStyle',attributes={'val':style})
        pJc = Element('jc',attributes={'val':jc})
        pPr.append(pStyle)
        pPr.append(pJc)
                    
        # Add the text the run, and the run to the paragraph
        paragraph.append(pPr)
        for t in text:
            run = Element('r')
            rPr = Element('rPr')
            # Apply styles
            if t[1].find('b') > -1:
                b = Element('b')
                rPr.append(b)
            if t[1].find('u') > -1:
                u = Element('u',attributes={'val':'single'})
                rPr.append(u)
            if t[1].find('i') > -1:
                i = Element('i')
                rPr.append(i)
            run.append(rPr)
            # Insert lastRenderedPageBreak for assistive technologies like
            # document narrators to know when a page break occurred.
            if breakbefore:
                lastRenderedPageBreak = Element('lastRenderedPageBreak')
                run.append(lastRenderedPageBreak)
            run.append(t[0])
            paragraph.append(run)
        
        self.xml = paragraph
        self.style = style
        self.jc = jc

def Heading(Element):
    def __init__(headingtext, headinglevel, lang='en'):
        '''Make a new heading, return the heading element'''
        lmap = {
            'en': 'Heading',
            'it': 'Titolo',
        }
        # Make our elements
        paragraph = Element('p')
        pr = Element('pPr')
        pStyle = Element('pStyle',attributes={'val':lmap[lang]+str(headinglevel)})    
        run = Element('r')
        text = Element('t',tagtext=headingtext)
        # Add the text the run, and the run to the paragraph
        pr.append(pStyle)
        run.append(text)
        paragraph.append(pr)   
        paragraph.append(run)    
        
        self.xml = paragraph

def Table(Element):
    def __init__(contents, heading=True, colw=None, cwunit='dxa', tblw=0, twunit='auto', borders={}, celstyle=None):
        '''Get a list of lists, return a table
>>>>>>> b0abe8d0c2ba44c3dd6675c5eb3198925a84431e
    
=======
    paragraph = Element('p')
    pr = Element('pPr')
    pStyle = Element('pStyle',attributes={'val':lmap[lang]+str(headinglevel)})
    run = Element('r')
    text = Element('t',tagtext=headingtext)
=======
        'fr': 'Titre',
    }
    # Make our elements
    paragraph = makeelement('p')
    pr = makeelement('pPr')
    pStyle = makeelement('pStyle', attributes={
                                        'val': lmap[lang] + str(headinglevel)})
    run = makeelement('r')
    text = makeelement('t', tagtext=headingtext)
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)
    paragraph.append(run)
    # Return the combined paragraph
    return paragraph

<<<<<<< HEAD
def table(contents, heading=True, colw=None, cwunit='dxa', tblw=0, twunit='auto', borders={}, celstyle=None):
    '''Get a list of lists, return a table

>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
=======

def table(contents, tblstyle=None, tbllook={'val': '0400'}, heading=True,
            colw=None, cwunit='dxa', tblw=0, twunit='auto', borders={},
            celstyle=None, rowstyle=None, table_props=None):
    '''Get a list of lists, return a table

>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
        @param list contents: A list of lists describing contents
                              Every item in the list can be a string or a valid
                              XML element itself. It can also be a list. In that case
                              all the listed elements will be merged into the cell.
<<<<<<< HEAD
        @param bool heading: Tells whether first line should be threated as heading
                             or not
        @param list colw: A list of interger. The list must have same element
=======
        @param string tblstyle: Specifies name of table style to override default if desired
        @param string tbllook: Specifies which elements of table style to to apply to this table,
                               e.g. {'firstColumn': 'false', 'firstRow': 'true'}, etc.
        @param bool heading: Tells whether first line should be treated as heading
                             or not
        @param list colw: A list of integer. The list must have same element
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
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
<<<<<<< HEAD
                             val: The style of the border, see http://www.schemacentral.com/sc/ooxml/t-w_ST_Border.htm
        @param list celstyle: Specify the style for each colum, list of dicts.
                              supported keys:
                              'align': specify the alignment, see paragraph documentation,
<<<<<<< HEAD
        
        @return lxml.etree: Generated XML etree element
<<<<<<< HEAD
    '''
    table = make_element('tbl')
    columns = len(contents[0])
    # Table properties
    tableprops = make_element('tblPr')
    tablestyle = make_element('tblStyle',attributes={'val':'ColorfulGrid-Accent1'})
    tableprops.append(tablestyle)
    tablewidth = make_element('tblW',attributes={'w':str(tblw),'type':str(twunit)})
    tableprops.append(tablewidth)
    if len(borders.keys()):
        tableborders = make_element('tblBorders')
=======

        @return lxml.etree: Generated XML etree element
    '''
    table = Element('tbl')
    columns = len(contents[0])
    # Table properties
    tableprops = Element('tblPr')
    tablestyle = Element('tblStyle',attributes={'val':'ColorfulGrid-Accent1'})
    tableprops.append(tablestyle)
    tablewidth = Element('tblW',attributes={'w':str(tblw),'type':str(twunit)})
    tableprops.append(tablewidth)
    if len(borders.keys()):
        tableborders = Element('tblBorders')
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            if b in borders.keys() or 'all' in borders.keys():
                k = 'all' if 'all' in borders.keys() else b
                attrs = {}
                for a in borders[k].keys():
                    attrs[a] = unicode(borders[k][a])
<<<<<<< HEAD
                borderelem = make_element(b,attributes=attrs)
                tableborders.append(borderelem)
        tableprops.append(tableborders)
    tablelook = make_element('tblLook',attributes={'val':'0400'})
    tableprops.append(tablelook)
    table.append(tableprops)    
    # Table Grid    
    tablegrid = make_element('tblGrid')
    for i in range(columns):
        tablegrid.append(make_element('gridCol',attributes={'w':str(colw[i]) if colw else '2390'}))
    table.append(tablegrid)     
    # Heading Row    
    row = make_element('tr')
    rowprops = make_element('trPr')
    cnfStyle = make_element('cnfStyle',attributes={'val':'000000100000'})
=======
                borderelem = Element(b,attributes=attrs)
                tableborders.append(borderelem)
        tableprops.append(tableborders)
    tablelook = Element('tblLook',attributes={'val':'0400'})
    tableprops.append(tablelook)
    table.append(tableprops)
    # Table Grid
    tablegrid = Element('tblGrid')
    for i in range(columns):
        tablegrid.append(Element('gridCol',attributes={'w':str(colw[i]) if colw else '2390'}))
    table.append(tablegrid)
    # Heading Row
    row = Element('tr')
    rowprops = Element('trPr')
    cnfStyle = Element('cnfStyle',attributes={'val':'000000100000'})
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
    rowprops.append(cnfStyle)
    row.append(rowprops)
    if heading:
        i = 0
        for heading in contents[0]:
<<<<<<< HEAD
            cell = make_element('tc')  
            # Cell properties  
            cellprops = make_element('tcPr')
=======
            cell = Element('tc')
            # Cell properties
            cellprops = Element('tcPr')
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
            if colw:
                wattr = {'w':str(colw[i]),'type':cwunit}
            else:
                wattr = {'w':'0','type':'auto'}
<<<<<<< HEAD
            cellwidth = make_element('tcW',attributes=wattr)
            cellstyle = make_element('shd',attributes={'val':'clear','color':'auto','fill':'548DD4','themeFill':'text2','themeFillTint':'99'})
            cellprops.append(cellwidth)
            cellprops.append(cellstyle)
            cell.append(cellprops)        
            # Paragraph (Content)
            if not type(heading) == list and not type(heading) == tuple:
=======
            cellwidth = Element('tcW',attributes=wattr)
            cellstyle = Element('shd',attributes={'val':'clear','color':'auto','fill':'548DD4','themeFill':'text2','themeFillTint':'99'})
=======
                             val: The style of the border, see http: //www.schemacentral.com/sc/ooxml/t-w_ST_Border.htm
        @param list celstyle: Specify the style for each colum, list of dicts.
                              supported keys:
                              'align': specify the alignment, see paragraph documentation,

        @return lxml.etree: Generated XML etree element
    '''
    table = makeelement('tbl')
    columns = len(contents[0])
    # Table properties
    tableprops = makeelement('tblPr')
    tablestyle = makeelement('tblStyle', attributes={'val': tblstyle if tblstyle else ''})
    tableprops.append(tablestyle)
    for attr in tableprops.iterchildren():
        if isinstance(attr, etree._Element):
            tableprops.append(attr)
        else:
            raise KeyError('what type of element to make?')
            prop = makeelement(k, attributes=attr)
            tableprops.append(prop)

    tablewidth = makeelement('tblW', attributes={'w': str(tblw), 'type': str(twunit)})
    tableprops.append(tablewidth)
    if len(borders.keys()):
        tableborders = makeelement('tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            if b in borders.keys() or 'all' in borders.keys():
                k = 'all' if 'all' in borders.keys() else b
            attrs = {}
            for a in borders[k].keys():
                attrs[a] = unicode(borders[k][a])
                if not 'val' in attrs:
                    # default border type
                    attrs['val'] = 'single'
            tableborders.append(makeelement(b, attributes=attrs))
        tableprops.append(tableborders)
    tablelook = makeelement('tblLook', attributes=tbllook)
    tableprops.append(tablelook)
    table.append(tableprops)
    # Table Grid
    tablegrid = makeelement('tblGrid')

    # gridCol width must be in dxa so convert pct to dxa if needed and if tblw
    # is defined (in dxa !)
    if colw:
        if tblw is not 0 and twunit is 'dxa' and cwunit is 'pct':

            colw = [tblw * int(size) / 5000 if size is not 'auto' else tblw / len(colw)
                                                            for size in colw]
            cwunit = 'dxa'
        for size in colw:
            if size is not 'auto':
                tablegrid.append(makeelement('gridCol', attributes={'w': str(size)}))
            else:
                tablegrid.append(makeelement('gridCol'))
    else:
        for i in range(columns):
            tablegrid.append(makeelement('gridCol', attributes={'w': '2390'}))

    table.append(tablegrid)
    # Heading Row
    row = makeelement('tr')
    rowprops = makeelement('trPr')
    cnfStyle = makeelement('cnfStyle', attributes={'val': '000000100000'})
    rowprops.append(cnfStyle)
    row.append(rowprops)

    if heading:
        i = 0
        for heading in contents[0]:
            cell = makeelement('tc')
            # Cell properties
            cellprops = makeelement('tcPr')
            if colw:
                wattr = {'w': str(colw[i]), 'type': cwunit}
            else:
                wattr = {'w': '0', 'type': 'auto'}
            cellwidth = makeelement('tcW', attributes=wattr)
            cellstyle = makeelement('shd', attributes={'val': 'clear', 'color': 'auto', 'fill': 'FFFFFF'})  # , 'themeFill': 'text2', 'themeFillTint': '99'
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
            cellprops.append(cellwidth)
            cellprops.append(cellstyle)
            cell.append(cellprops)
            # Paragraph (Content)
            if not isinstance(heading, (list, tuple)):
<<<<<<< HEAD
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
                heading = [heading,]
=======
                heading = [heading, ]
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
            for h in heading:
                if isinstance(h, etree._Element):
                    cell.append(h)
                else:
<<<<<<< HEAD
                    cell.append(paragraph(h,jc='center'))
            row.append(cell)
            i += 1
<<<<<<< HEAD
        table.append(row)          
    # Contents Rows
    for contentrow in contents[1 if heading else 0:]:
        row = make_element('tr')     
        i = 0
        for content in contentrow:   
            cell = make_element('tc')
            # Properties
            cellprops = make_element('tcPr')
=======
        table.append(row)
    # Contents Rows
    for contentrow in contents[1 if heading else 0:]:
        row = Element('tr')
        i = 0
        for content in contentrow:
            cell = Element('tc')
            # Properties
            cellprops = Element('tcPr')
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
            if colw:
                wattr = {'w':str(colw[i]),'type':cwunit}
            else:
                wattr = {'w':'0','type':'auto'}
<<<<<<< HEAD
            cellwidth = make_element('tcW',attributes=wattr)
            cellprops.append(cellwidth)
            cell.append(cellprops)
            # Paragraph (Content)
            if not type(content) == list and not type(content) == tuple:
=======
            cellwidth = Element('tcW',attributes=wattr)
            cellprops.append(cellwidth)
            cell.append(cellprops)
            # Paragraph (Content)
            if not isinstance(content, (list, tuple)):
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
                content = [content,]
            for c in content:
                if isinstance(c, etree._Element):
                    cell.append(c)
                else:
                    if celstyle and 'align' in celstyle[i].keys():
                        align = celstyle[i]['align']
                    else:
                        align = 'left'
                    cell.append(paragraph(c,jc=align))
<<<<<<< HEAD
            row.append(cell)    
            i += 1
        table.append(row)   
    return table


def picture(document, picname, picdescription, pixelwidth=None,
=======
=======
                    cell.append(paragraph(h, jc='center'))
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
                wattr = {'w': str(colw[i]), 'type': cwunit}
            else:
                wattr = {'w': '0', 'type': 'auto'}
            cellwidth = makeelement('tcW', attributes=wattr)
            cellprops.append(cellwidth)
            align = 'left'
            cell_spec_style = {}

            if celstyle:
                cell_spec_style = deepcopy(celstyle[i])

            if isinstance(content_cell, dict):
                # accepted keys : 'content', 'style',
                if 'style' in content_cell:
                    cell_spec_style.update(content_cell['style'])
                if 'content' in content_cell:
                    content_cell = content_cell['content']
                else:
                    content_cell = ''

            # spec. align property
            SPEC_PROPS = ['align', ]
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
                content_cell = [content_cell, ]
            for c in content_cell:
                # cell.append(cellprops)
                if isinstance(c, etree._Element):
                    cell.append(c)
                else:
                    cell.append(paragraph(c, jc=align))
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
            row.append(cell)
            i += 1
        table.append(row)
    return table

<<<<<<< HEAD
def picture(relationshiplist, picname, picdescription, pixelwidth=None,
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
            pixelheight=None, nochangeaspect=True, nochangearrowheads=True):
    '''Take a relationshiplist, picture file name, and return a paragraph containing the image
    and an updated relationshiplist'''
    # http://openxmldeveloper.org/articles/462.aspx
    # Create an image. Size may be specified, otherwise it will based on the
<<<<<<< HEAD
    # pixel size of image. Return a paragraph containing the picture'''  
    document.word_relationships.to_copy.append([picname, os.path.abspath(picname)])
=======
    # pixel size of image. Return a paragraph containing the picture'''
    # Copy the file into the media dir
    media_dir = join(TEMPLATE_DIR,'word','media')
    if not os.path.isdir(media_dir):
        os.mkdir(media_dir)
    shutil.copyfile(picname, join(media_dir,picname))
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f

    # Check if the user has specified a size
    if not pixelwidth or not pixelheight:
        # If not, get info from the picture itself
        pixelwidth,pixelheight = Image.open(picname).size[0:2]

    # OpenXML measures on-screen objects in English Metric Units
<<<<<<< HEAD
    # 1cm = 36000 EMUs            
    emuperpixel = 12667
    width = str(pixelwidth * emuperpixel)
    height = str(pixelheight * emuperpixel)   
    
    # Set relationship ID to the first available  
    picid = '2'
    picrelid = 'rId'+str(len(document.word_relationships.relationshiplist) + 1)
    picid = str(len(document.word_relationships.relationshiplist) + 1)
    document.word_relationships.relationshiplist.append([
        picrelid, 
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'media/'+picname])
    # There are 3 main elements inside a picture
    # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
    blipfill = make_element('blipFill',nsprefix='pic')
    blipfill.append(make_element('blip',nsprefix='a',attrnsprefix='r',attributes={'embed':picrelid}))
    stretch = make_element('stretch',nsprefix='a')
    stretch.append(make_element('fillRect',nsprefix='a'))
    blipfill.append(make_element('srcRect',nsprefix='a'))
    blipfill.append(stretch)
    
    # 2. The non visual picture properties 
    nvpicpr = make_element('nvPicPr',nsprefix='pic')
    cnvpr = make_element('cNvPr',nsprefix='pic',
                        attributes={'id':'0','name':'Picture 1','descr':picname}) 
    nvpicpr.append(cnvpr) 
    cnvpicpr = make_element('cNvPicPr',nsprefix='pic')                           
    cnvpicpr.append(make_element('picLocks', nsprefix='a', 
                    attributes={'noChangeAspect':str(int(nochangeaspect)),
                    'noChangeArrowheads':str(int(nochangearrowheads))}))
    nvpicpr.append(cnvpicpr)
        
    # 3. The Shape properties
    sppr = make_element('spPr',nsprefix='pic',attributes={'bwMode':'auto'})
    xfrm = make_element('xfrm',nsprefix='a')
    xfrm.append(make_element('off',nsprefix='a',attributes={'x':'0','y':'0'}))
    xfrm.append(make_element('ext',nsprefix='a',attributes={'cx':width,'cy':height}))
    prstgeom = make_element('prstGeom',nsprefix='a',attributes={'prst':'rect'})
    prstgeom.append(make_element('avLst',nsprefix='a'))
    sppr.append(xfrm)
    sppr.append(prstgeom)
    
    # Add our 3 parts to the picture element
    pic = make_element('pic',nsprefix='pic')    
    pic.append(nvpicpr)
    pic.append(blipfill)
    pic.append(sppr)
    
    # Now make the supporting elements
    # The following sequence is just: make element, then add its children
    graphicdata = make_element('graphicData',nsprefix='a',
        attributes={'uri':'http://schemas.openxmlformats.org/drawingml/2006/picture'})
    graphicdata.append(pic)
    graphic = make_element('graphic',nsprefix='a')
    graphic.append(graphicdata)

    framelocks = make_element('graphicFrameLocks',nsprefix='a',attributes={'noChangeAspect':'1'})    
    framepr = make_element('cNvGraphicFramePr',nsprefix='wp')
    framepr.append(framelocks)
    docpr = make_element('docPr',nsprefix='wp',
        attributes={'id':picid,'name':'Picture 1','descr':picdescription})
    effectextent = make_element('effectExtent',nsprefix='wp',
        attributes={'l':'25400','t':'0','r':'0','b':'0'})
    extent = make_element('extent',nsprefix='wp',attributes={'cx':width,'cy':height})
    inline = make_element('inline',
=======
=======
def picture(doc, picpath, picdescription, pixelwidth=0,
            pixelheight=0, nochangeaspect=True, nochangearrowheads=True,
            template_dir=None, relationships=None):
    ''' Take a document, a picture path, and
        return a paragraph containing the image
        and an updated relationshiplist
    '''
    # http: //openxmldeveloper.org/articles/462.aspx
    # Create an image. Size may be specified, otherwise it will based on the
    # pixel size of image. Return a paragraph containing the picture

    pic_relpath = join('media', basename(picpath))

    # add the picture to the
    doc.add_file(picpath, 'word/'+pic_relpath)

    # Check if the user has specified a size
    # If not, get info from the picture itself

    basew, baseh = Image.open(picpath).size[0: 2]
    if not pixelheight:
        pixelheight = (baseh * pixelwidth / basew) if pixelwidth else baseh
    if not pixelwidth:
        pixelwidth = (basew * pixelheight / baseh) if pixelheight else basew

    # OpenXML measures on-screen objects in English Metric Units
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    # 1cm = 36000 EMUs
    emuperpixel = 12667
    width = str(pixelwidth * emuperpixel)
    height = str(pixelheight * emuperpixel)

    # Set relationship ID to the first available
<<<<<<< HEAD
    picid = '2'
    picrelid = 'rId'+str(len(relationshiplist)+1)
    relationshiplist.append([
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'media/'+picname])

    # There are 3 main elements inside a picture
    # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
    blipfill = Element('blipFill',nsprefix='pic')
    blipfill.append(Element('blip',nsprefix='a',attrnsprefix='r',attributes={'embed':picrelid}))
    stretch = Element('stretch',nsprefix='a')
    stretch.append(Element('fillRect',nsprefix='a'))
    blipfill.append(Element('srcRect',nsprefix='a'))
    blipfill.append(stretch)

    # 2. The non visual picture properties
    nvpicpr = Element('nvPicPr',nsprefix='pic')
    cnvpr = Element('cNvPr',nsprefix='pic',
                        attributes={'id':'0','name':'Picture 1','descr':picname})
    nvpicpr.append(cnvpr)
    cnvpicpr = Element('cNvPicPr',nsprefix='pic')
    cnvpicpr.append(Element('picLocks', nsprefix='a',
                    attributes={'noChangeAspect':str(int(nochangeaspect)),
                    'noChangeArrowheads':str(int(nochangearrowheads))}))
    nvpicpr.append(cnvpicpr)

    # 3. The Shape properties
    sppr = Element('spPr',nsprefix='pic',attributes={'bwMode':'auto'})
    xfrm = Element('xfrm',nsprefix='a')
    xfrm.append(Element('off',nsprefix='a',attributes={'x':'0','y':'0'}))
    xfrm.append(Element('ext',nsprefix='a',attributes={'cx':width,'cy':height}))
    prstgeom = Element('prstGeom',nsprefix='a',attributes={'prst':'rect'})
    prstgeom.append(Element('avLst',nsprefix='a'))
=======
    if relationships is None:
        # default is the word part of the doc
        relationships = doc.wordrelationships
    picrelid = new_id(relationships)
    relationships.append(makeelement(
                            'Relationship',
                            attributes={
                                'Id': picrelid,
                                'Type': image_relationship,
                                'Target': pic_relpath},
                            nsprefix=None))

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
    xfrm = makeelement('xfrm', nsprefix='a')
    xfrm.append(makeelement('off', nsprefix='a', attributes={'x': '0', 'y': '0'}))
    xfrm.append(makeelement('ext', nsprefix='a', attributes={'cx': width, 'cy': height}))
    prstgeom = makeelement('prstGeom', nsprefix='a', attributes={'prst': 'rect'})
    prstgeom.append(makeelement('avLst', nsprefix='a'))
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    sppr.append(xfrm)
    sppr.append(prstgeom)

    # Add our 3 parts to the picture element
<<<<<<< HEAD
    pic = Element('pic',nsprefix='pic')
=======
    pic = makeelement('pic', nsprefix='pic')
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    pic.append(nvpicpr)
    pic.append(blipfill)
    pic.append(sppr)

    # Now make the supporting elements
    # The following sequence is just: make element, then add its children
<<<<<<< HEAD
    graphicdata = Element('graphicData',nsprefix='a',
        attributes={'uri':'http://schemas.openxmlformats.org/drawingml/2006/picture'})
    graphicdata.append(pic)
    graphic = Element('graphic',nsprefix='a')
    graphic.append(graphicdata)

    framelocks = Element('graphicFrameLocks',nsprefix='a',attributes={'noChangeAspect':'1'})
    framepr = Element('cNvGraphicFramePr',nsprefix='wp')
    framepr.append(framelocks)
    docpr = Element('docPr',nsprefix='wp',
        attributes={'id':picid,'name':'Picture 1','descr':picdescription})
    effectextent = Element('effectExtent',nsprefix='wp',
        attributes={'l':'25400','t':'0','r':'0','b':'0'})
    extent = Element('extent',nsprefix='wp',attributes={'cx':width,'cy':height})
    inline = Element('inline',
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
        attributes={'distT':"0",'distB':"0",'distL':"0",'distR':"0"},nsprefix='wp')
=======
    graphicdata = makeelement('graphicData', nsprefix='a',
                              attributes={'uri': 'http: //schemas.openxmlformats.org/drawingml/2006/picture'})
    graphicdata.append(pic)
    graphic = makeelement('graphic', nsprefix='a')
    graphic.append(graphicdata)

    framelocks = makeelement('graphicFrameLocks', nsprefix='a', attributes={'noChangeAspect': '1'})
    framepr = makeelement('cNvGraphicFramePr', nsprefix='wp')
    framepr.append(framelocks)
    docpr = makeelement('docPr', nsprefix='wp',
                        attributes={'id': picrelid, 'name': 'Picture 1', 'descr': picdescription})
    effectextent = makeelement('effectExtent', nsprefix='wp',
                               attributes={'l': '25400', 't': '0', 'r': '0', 'b': '0'})
    extent = makeelement('extent', nsprefix='wp', attributes={'cx': width, 'cy': height})
    inline = makeelement('inline',
                         attributes={'distT': "0", 'distB': "0", 'distL': "0", 'distR': "0"},
                         nsprefix='wp')
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    inline.append(extent)
    inline.append(effectextent)
    inline.append(docpr)
    inline.append(framepr)
    inline.append(graphic)
<<<<<<< HEAD
<<<<<<< HEAD
    drawing = make_element('drawing')
    drawing.append(inline)
    run = make_element('r')
    run.append(drawing)
    paragraph = make_element('p')
    paragraph.append(run)
    return paragraph



=======
        '''
        table = Element('tbl')
        columns = len(contents[0])
        # Table properties
        tableprops = Element('tblPr')
        tablestyle = Element('tblStyle',attributes={'val':'ColorfulGrid-Accent1'})
        tableprops.append(tablestyle)
        tablewidth = Element('tblW',attributes={'w':str(tblw),'type':str(twunit)})
        tableprops.append(tablewidth)
        if len(borders.keys()):
            tableborders = Element('tblBorders')
            for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                if b in borders.keys() or 'all' in borders.keys():
                    k = 'all' if 'all' in borders.keys() else b
                    attrs = {}
                    for a in borders[k].keys():
                        attrs[a] = unicode(borders[k][a])
                    borderelem = Element(b,attributes=attrs)
                    tableborders.append(borderelem)
            tableprops.append(tableborders)
        tablelook = Element('tblLook',attributes={'val':'0400'})
        tableprops.append(tablelook)
        table.append(tableprops)    
        # Table Grid    
        tablegrid = Element('tblGrid')
        for i in range(columns):
            tablegrid.append(Element('gridCol',attributes={'w':str(colw[i]) if colw else '2390'}))
        table.append(tablegrid)     
        # Heading Row    
        row = Element('tr')
        rowprops = Element('trPr')
        cnfStyle = Element('cnfStyle',attributes={'val':'000000100000'})
        rowprops.append(cnfStyle)
        row.append(rowprops)
        if heading:
            i = 0
            for heading in contents[0]:
                cell = Element('tc')  
                # Cell properties  
                cellprops = Element('tcPr')
                if colw:
                    wattr = {'w':str(colw[i]),'type':cwunit}
                else:
                    wattr = {'w':'0','type':'auto'}
                cellwidth = Element('tcW',attributes=wattr)
                cellstyle = Element('shd',attributes={'val':'clear','color':'auto','fill':'548DD4','themeFill':'text2','themeFillTint':'99'})
                cellprops.append(cellwidth)
                cellprops.append(cellstyle)
                cell.append(cellprops)        
                # Paragraph (Content)
                if not type(heading) == list and not type(heading) == tuple:
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
            row = Element('tr')     
            i = 0
            for content in contentrow:   
                cell = Element('tc')
                # Properties
                cellprops = Element('tcPr')
                if colw:
                    wattr = {'w':str(colw[i]),'type':cwunit}
                else:
                    wattr = {'w':'0','type':'auto'}
                cellwidth = Element('tcW',attributes=wattr)
                cellprops.append(cellwidth)
                cell.append(cellprops)
                # Paragraph (Content)
                if not type(content) == list and not type(content) == tuple:
                    content = [content,]
                for c in content:
                    if isinstance(c, etree._Element):
                        cell.append(c)
                    else:
                        if celstyle and 'align' in celstyle[i].keys():
                            align = celstyle[i]['align']
                        else:
                            align = 'left'
                        cell.append(paragraph(c,jc=align))
                row.append(cell)    
                i += 1
            table.append(row)   
        
        self.xml = table              

class Picture(Element):
    def __init__(relationshiplist, picname, picdescription, pixelwidth=None,
            pixelheight=None, nochangeaspect=True, nochangearrowheads=True):
        '''Take a relationshiplist, picture file name, and return a paragraph containing the image
        and an updated relationshiplist'''
        # http://openxmldeveloper.org/articles/462.aspx
        # Create an image. Size may be specified, otherwise it will based on the
        # pixel size of image. Return a paragraph containing the picture'''  
        # Copy the file into the media dir
        media_dir = join(template_dir,'word','media')
        if not os.path.isdir(media_dir):
            os.mkdir(media_dir)
        shutil.copyfile(picname, join(media_dir,picname))
    
        # Check if the user has specified a size
        if not pixelwidth or not pixelheight:
            # If not, get info from the picture itself
            pixelwidth,pixelheight = Image.open(picname).size[0:2]
    
        # OpenXML measures on-screen objects in English Metric Units
        # 1cm = 36000 EMUs            
        emuperpixel = 12667
        width = str(pixelwidth * emuperpixel)
        height = str(pixelheight * emuperpixel)   
        
        # Set relationship ID to the first available  
        picid = '2'    
        picrelid = 'rId'+str(len(relationshiplist)+1)
        relationshiplist.append([
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            'media/'+picname])
        
        # There are 3 main elements inside a picture
        # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
        blipfill = Element('blipFill',nsprefix='pic')
        blipfill.append(Element('blip',nsprefix='a',attrnsprefix='r',attributes={'embed':picrelid}))
        stretch = Element('stretch',nsprefix='a')
        stretch.append(Element('fillRect',nsprefix='a'))
        blipfill.append(Element('srcRect',nsprefix='a'))
        blipfill.append(stretch)
        
        # 2. The non visual picture properties 
        nvpicpr = Element('nvPicPr',nsprefix='pic')
        cnvpr = Element('cNvPr',nsprefix='pic',
                            attributes={'id':'0','name':'Picture 1','descr':picname}) 
        nvpicpr.append(cnvpr) 
        cnvpicpr = Element('cNvPicPr',nsprefix='pic')                           
        cnvpicpr.append(Element('picLocks', nsprefix='a', 
                        attributes={'noChangeAspect':str(int(nochangeaspect)),
                        'noChangeArrowheads':str(int(nochangearrowheads))}))
        nvpicpr.append(cnvpicpr)
            
        # 3. The Shape properties
        sppr = Element('spPr',nsprefix='pic',attributes={'bwMode':'auto'})
        xfrm = Element('xfrm',nsprefix='a')
        xfrm.append(Element('off',nsprefix='a',attributes={'x':'0','y':'0'}))
        xfrm.append(Element('ext',nsprefix='a',attributes={'cx':width,'cy':height}))
        prstgeom = Element('prstGeom',nsprefix='a',attributes={'prst':'rect'})
        prstgeom.append(Element('avLst',nsprefix='a'))
        sppr.append(xfrm)
        sppr.append(prstgeom)
        
        # Add our 3 parts to the picture element
        pic = Element('pic',nsprefix='pic')    
        pic.append(nvpicpr)
        pic.append(blipfill)
        pic.append(sppr)
        
        # Now make the supporting elements
        # The following sequence is just: make element, then add its children
        graphicdata = Element('graphicData',nsprefix='a',
            attributes={'uri':'http://schemas.openxmlformats.org/drawingml/2006/picture'})
        graphicdata.append(pic)
        graphic = Element('graphic',nsprefix='a')
        graphic.append(graphicdata)
    
        framelocks = Element('graphicFrameLocks',nsprefix='a',attributes={'noChangeAspect':'1'})    
        framepr = Element('cNvGraphicFramePr',nsprefix='wp')
        framepr.append(framelocks)
        docpr = Element('docPr',nsprefix='wp',
            attributes={'id':picid,'name':'Picture 1','descr':picdescription})
        effectextent = Element('effectExtent',nsprefix='wp',
            attributes={'l':'25400','t':'0','r':'0','b':'0'})
        extent = Element('extent',nsprefix='wp',attributes={'cx':width,'cy':height})
        inline = Element('inline',
            attributes={'distT':"0",'distB':"0",'distL':"0",'distR':"0"},nsprefix='wp')
        inline.append(extent)
        inline.append(effectextent)
        inline.append(docpr)
        inline.append(framepr)
        inline.append(graphic)
        drawing = Element('drawing')
        drawing.append(inline)
        run = Element('r')
        run.append(drawing)
        paragraph = Element('p')
        paragraph.append(run)
        
        self.xml = paragraph
        self.relationship_list = relationshiplist
        
    def with_relationships(self):
        return (self.relationship_list, paragraph)
>>>>>>> b0abe8d0c2ba44c3dd6675c5eb3198925a84431e
=======
    drawing = Element('drawing')
    drawing.append(inline)
    run = Element('r')
    run.append(drawing)
    paragraph = Element('p')
    paragraph.append(run)
    return relationshiplist,paragraph

>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
=======
    drawing = makeelement('drawing')
    drawing.append(inline)
    run = makeelement('r')
    run.append(drawing)
    paragraph = makeelement('p')
    paragraph.append(run)
    return  paragraph
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
