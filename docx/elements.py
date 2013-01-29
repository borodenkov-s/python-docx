#-*- coding:utf-8 -*-

import os
from os.path import join, basename
import shutil
from copy import deepcopy

from lxml import etree
try:
    from PIL import Image
except ImportError:
    import Image

from .metadata import nsprefixes, FORMAT, PAGESETTINGS
from .metadata import TEMPLATE_DIR, TMP_TEMPLATE_DIR, image_relationship, header_relationship

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
    # Deal with list of nsprefix by making namespacemap
    namespacemap = None
    if isinstance(nsprefix, list):
        namespacemap = {}
        for prefix in nsprefix:
            namespacemap[prefix] = nsprefixes[prefix]
        nsprefix = nsprefix[0]  # FIXME: code below expects a single prefix
    if nsprefix:
        namespace = '{' + nsprefixes[nsprefix] + '}'
    else:
        # For when namespace = None
        namespace = ''

    newelement = etree.Element(namespace + tagname, nsmap=namespacemap)
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
            attributenamespace = '{' + nsprefixes[attrnsprefix] + '}'

        for attrname, attrvalue in attributes.iteritems():
            if isinstance(attrvalue,dict):
                ns = '{' + nsprefixes[attrvalue['prefix']] + '}'
                newelement.set(ns + attrname, attrvalue['value'])
            else:
                newelement.set(attributenamespace + attrname, attrvalue)
    if tagtext:
        newelement.text = tagtext
    return newelement

def pagebreak(breaktype='page', orient='portrait', pageformat='letter'):
    '''Insert a break, default 'page'.
    See http://openxmldeveloper.org/forums/thread/4075.aspx
    Return our page break element.'''
    # Need to enumerate different types of page breaks.
    validtypes = ['page', 'section']
    if breaktype not in validtypes:
        raise ValueError('Page break style "%s" not implemented. Valid styles: %s.' % (breaktype, validtypes))
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
        sectPr.append(pgSz)
        pPr.append(sectPr)
        pagebreak.append(pPr)
    return pagebreak

def paragraph(paratext, style='BodyText', breakbefore=False, jc='left',
                                        font='Times New Roman', fontsize=12,
                                        tabs = None):
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
        if not isinstance(pt, (list, tuple)):
            pt = (pt, {})
        lines = pt[0].split('\n')
        for l in lines[:-1]:
            text.append((makeelement('t',tagtext=l), pt[1], True))  # with line break
        text.append((makeelement('t',tagtext=lines[-1]), pt[1], False))  # the last line, without line break

    pPr = makeelement('pPr')
    pStyle = makeelement('pStyle',attributes={'val':style})
    pJc = makeelement('jc',attributes={'val':jc})

    if tabs:
        pTabs = makeelement('tabs')
        for tab in tabs:
            pTabs.append(makeelement('tab',  attributes = tab))

        pPr.append(pTabs)

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
        if 'style' in t[1]:
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

def heading(headingtext, headinglevel, lang='en'):
    '''Make a new heading, return the heading element'''
    lmap = {
        'en': 'Heading',
        'it': 'Titolo',
        'fr': 'Titre',
    }
    # Make our elements
    paragraph = makeelement('p')
    pr = makeelement('pPr')
    pStyle = makeelement('pStyle', attributes={'val': lmap[lang] + str(headinglevel)})
    run = makeelement('r')
    text = makeelement('t', tagtext=headingtext)
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)
    paragraph.append(run)
    # Return the combined paragraph
    return paragraph

def table(contents, tblstyle=None, tbllook={'val': '0400'}, heading=True,
            colw=None, cwunit='dxa', tblw=0, twunit='auto', borders={},
            celstyle=None, rowstyle=None, table_props=None):
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
        @param list colw: A list of integer. The list must have same element
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

    # gridCol width must be in dxa so convert pct to dxa if needed and if tblw
    # is defined (in dxa !)
    if colw:
        if tblw is not 0 and twunit is 'dxa' and cwunit is 'pct':
            colw = [ tblw*int(size)/5000 if size is not 'auto' else 5000
                                                            for size in colw ]
        for size in colw:
            if size is not 'auto':
                tablegrid.append(makeelement('gridCol',attributes={'w': str(size)}))
            else:
                tablegrid.append(makeelement('gridCol'))
    else:
        for i in range(columns):
            tablegrid.append(makeelement('gridCol',attributes={'w': '2390'}))


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
                heading = [heading, ]
            for h in heading:
                if isinstance(h, etree._Element):
                    cell.append(h)
                else:
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
        template_dir = TMP_TEMPLATE_DIR
    media_dir = join(TEMPLATE_DIR, 'word', 'media')


    if not os.path.isdir(media_dir):
        os.makedirs(media_dir)

    pic_tmppath = join(media_dir, basename(picpath))
    pic_relpath = join('media', basename(picpath))

    shutil.copyfile(picpath,join(media_dir, basename(picpath)))

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
        relationshiplist.append(makeelement(
                            'Relationship',
                            attributes={
                                'Id': 'rId' + str(len(relationshiplist)+1),
                                'Type': image_relationship,
                                'Target': join('media', basename(picpath))},
                            nsprefix=None))
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
    xfrm = makeelement('xfrm', nsprefix='a')
    xfrm.append(makeelement('off', nsprefix='a', attributes={'x': '0', 'y': '0'}))
    xfrm.append(makeelement('ext', nsprefix='a', attributes={'cx': width, 'cy': height}))
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
    return  paragraph
