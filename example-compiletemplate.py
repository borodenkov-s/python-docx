#!/usr/bin/env python2.6
'''
This file opens 'example-template.docx' and replaces some tags in the form
{{TAGNAME}} with custom attributes, then saves the output in out.docx

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
'''
import docx

import re
import zipfile
from lxml import etree

# -----------------------------------------------------------------------------
# Utility replacement function
# -----------------------------------------------------------------------------

def replaceTag(doc, tag, replace, fmt = {}):
    """ Searches for {{tag}} and replaces it with replace.

        Replace is a list with two indexes: 0=type, 1=The replacement
            Supported values for type:
                'str': <string> Renders a simple text string
                'tab': <list> Renders a table, use fmt to tune look
    """

    if replace[0] == 'str':
        repl = unicode(replace[1])
    elif replace[0] == 'tab':
        # Will make a table
        
        r = []
        for el in replace[1]:
            # Unicodize
            r.append( map(lambda i: unicode(i), el) )
        if not len(r):
            # Empty table
            repl = ''
        else:
            repl = docx.table(
                r,
                heading = fmt['heading'] if 'heading' in fmt.keys() else False,
                colw = fmt['colw'] if 'colw' in fmt.keys() else None,
                cwunit = fmt['cwunit'] if 'cwunit' in fmt.keys() else 'dxa',
                tblw = fmt['tblw'] if 'tblw' in fmt.keys() else 0,
                twunit = fmt['twunit'] if 'twunit' in fmt.keys() else 'auto',
                borders = fmt['borders'] if 'borders' in fmt.keys() else {},
                celstyle = fmt['celstyle'] if 'celstyle' in fmt.keys() else None,
                headstyle = fmt['headstyle'] if 'headstyle' in fmt.keys() else {},
            )
    else:
        raise NotImplementedError, "Unsupported " + replace[0] + " tag type!"

    return docx.advReplace(doc, '\{\{'+re.escape(tag)+'\}\}', repl)

# -----------------------------------------------------------------------------

# Load the original template
template = zipfile.ZipFile('example-template.docx',mode='r')
if template.testzip():
    raise Exception('File is corrupted!')

# List of section to modify
# (<section>, <namespace>)
actlist = (
    ('word/document.xml', '/w:document/w:body'), # Main
    ('word/footer1.xml', '/w:ftr'), # FOOTER
    ('word/header1.xml', '/w:hdr'), # HEADER
)

# Will store modified sections here
outdoc = {}

for curact in actlist:
    xmlcontent = template.read(curact[0])
    outdoc[curact[0]] = etree.fromstring(xmlcontent)

    # Will work on body
    docbody = outdoc[curact[0]].xpath(curact[1], namespaces=docx.nsprefixes)[0]

    # Replace some tags
    docbody = replaceTag(docbody, 'SUBJECT', ('str', 'This is a replaced subject') )
    docbody = replaceTag(docbody, 'FOOTER', ('str', 'This text comes from python') )
    docbody = replaceTag(docbody, 'TABLEITEMS', ('tab', (
        ( 'Header 1', 'Header 2', 'Header 3' ),
        ( 'This is an example', 'table', 'generated' ),
        ( 'using', 'pydocx', '' ),
    )),
    {
        'heading': True,
        'colw': [ 2400, 1000, 400], # 5000 = 100%
        'cwunit': 'pct',
        'tblw': 3800, # 5000 / 50 = 100%
        'twunit': 'pct',
        'borders': {
            'all': {
                'color': 'auto',
                'space': 0,
                'sz': 6,
                'val': 'single',
            },
        },
        'celstyle': [
            {'align': 'center'},
            {'align': 'left'},
            {'align': 'right'},
        ],
        'headstyle': { 'fill':'C6D9F1', 'themeFill':None, 'themeFillTint':None },
    })
    
    # Cleaning
    docbody = docx.clean(docbody)

# ------------------------------
# Save output
# ------------------------------

# Prepare output file
outfile = zipfile.ZipFile('out.docx',mode='w',compression=zipfile.ZIP_DEFLATED)

# Copy unmodified sections
for f in template.namelist():
    if not f in map(lambda i: i[0], actlist):
        fo = template.open(f,'rU')
        data = fo.read()
        outfile.writestr(f,data)
        fo.close()

# The copy modified sections
for sec in outdoc.keys():
    treestring = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "\n"
    treestring += etree.tostring(outdoc[sec], pretty_print=True)
    outfile.writestr(sec,treestring)

# Done. close files.
outfile.close()
template.close()
