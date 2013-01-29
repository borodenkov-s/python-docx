# -*- coding: utf-8 -*-
'''
Open and modify Microsoft Word 2007 docx files (called 'OpenXML'
and 'Office OpenXML' by Microsoft)

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
'''

from copy import deepcopy
from datetime import datetime

from lxml import etree
from dateutil import parser as dateParser

from zipfile import ZipFile
import shutil
import re
import os


from tempfile import NamedTemporaryFile
from .utils import findTypeParent, dir_to_docx

from .metadata import nsprefixes, TEMPLATE_DIR, PAGESETTINGS
from .elements import makeelement

import logging
log = logging.getLogger(__name__)


class Docx(object):
    trees_and_files = {"document": 'word/document.xml',
                     "coreprops": 'docProps/core.xml',
                     "header": 'word/header1.xml',
                     "appprops": 'docProps/app.xml',
                     "contenttypes": '[Content_Types].xml',
                     "websettings": 'word/webSettings.xml',
                     "wordrelationships": 'word/_rels/document.xml.rels',
                     "styles": 'word/styles.xml'
                         }

    def __init__(self, srcfile=None):
        create_new_doc = srcfile is None
        self.settings = {}

        self._orig_docx = srcfile
        self._tmp_file = NamedTemporaryFile()

        if create_new_doc:
            srcfile = self.__generate_empty_docx()

        shutil.copyfile(srcfile, self._tmp_file.name)
        self._docx = ZipFile(self._tmp_file.name, mode='a')

        for tree, relpathfile in self.trees_and_files.items():
            self._load_etree(tree, relpathfile)

        self.docbody = self.document.xpath('/w:document/w:body',
                                                    namespaces=nsprefixes)[0]
        self.headercontent = self.header.xpath('/w:hdr',
                                                    namespaces=nsprefixes)[0]

        if create_new_doc:
            self.created = datetime.utcnow()

    def __new__(cls, *args, **kwargs):
        # Make getters and setter for the core properties
        def set_coreprop_property(prop, to_python=unicode, to_str=unicode):
            getter = lambda self: to_python(self._get_coreprop_val(prop))
            setter = lambda self, val: self._set_coreprop_val(prop, to_str(val))
            setattr(cls, prop, property(getter, setter))

        for prop in ['title', 'subject', 'creator', 'description',
                     'lastModifiedBy', 'revision']:
            set_coreprop_property(prop)

        for datetimeprop in ['created', 'modified']:
            set_coreprop_property(datetimeprop,
                to_python=dateParser.parse,
                to_str=lambda obj: (obj.isoformat()
                                    if hasattr(obj, 'isoformat')
                                    else dateParser.parse(obj).isoformat())
            )
        return super(Docx, cls).__new__(cls, *args, **kwargs)

    def append(self, *args, **kwargs):
        return self.docbody.append(*args, **kwargs)

    def set_page_settings(self, settings=None):
        ''' Add or update page settings
            with settings = None :  add defaults
            else update settings dict
        '''
        def merge(d1, d2):
            for k1, v1 in d1.iteritems():
                if not k1 in d2:
                    d2[k1] = v1
                elif isinstance(v1, dict):
                    merge(v1, d2[k1])
            return d2

        if not settings or not self.settings:
            self.settings = PAGESETTINGS

        if settings:
            self.settings = merge(self.settings, settings)

    def _apply_page_settings(self):
        if not self.settings:
            self.settings = PAGESETTINGS

        sectPr = makeelement('sectPr')
        sectPr.append(makeelement(
                            'headerReference',
                            attributes={
                                'id': {'value': 'rId7', 'prefix': 'r'},
                                'type': 'default',},
                            nsprefix='w',))
        for settingname, settingattrs in self.settings.iteritems():
            #log.info('applying settings: %s = %s' %(settingname,settingattrs))
            sectPr.append(makeelement(settingname, attributes=settingattrs))
        self.docbody.append(sectPr)


    def search(self, search):
        '''Search a document for a regex, return success / fail result'''
        document = self.docbody
        result = False
        searchre = re.compile(search)
        for element in document.iter():
            if element.tag == '{%s}t' % nsprefixes['w']:  # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        result = True
        return result

    def replace(self, search, replace):
        '''Replace all occurences of string with a different string, return updated document'''
        newdocument = self.docbody
        searchre = re.compile(search)
        for element in newdocument.iter():
            if element.tag == '{%s}t' % nsprefixes['w']:  # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        element.text = re.sub(search, replace, element.text)
        return newdocument

    def clean(self):
        """ Perform misc cleaning operations on documents.
            Returns cleaned document.
        """

        newdocument = self.docbody

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

    def load_style(self, stylefile):
        if os.path.isfile(stylefile):
            with open(stylefile, 'r') as f:
                try:
                    self.styles = etree.fromstring(f.read())
                    log.info('new style load from %s' % stylefile)
                except:
                    log.debug('unable to load xml tree from %s' % stylefile)
        else:
            log.debug('%s is not a file'% stylefile)

    def advReplace(self, search, replace, bs=3):
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

        newdocument = self.docbody

        # Compile the search regexp
        searchre = re.compile(search)

        # Will match against searchels. Searchels is a list that contains last
        # n text elements found in the document. 1 < n < bs
        searchels = []

        for element in newdocument.iter():
            if element.tag == '{%s}t' % nsprefixes['w']:  # t (text) elements
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
                            e = range(s, els_len)
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
                            log.debug("Matched text (splitted): %s", map(lambda i: i.text, searchels))
                            log.debug("Matched at position: %s", match.start())
                            log.debug("matched in elements: %s", e)
                            if isinstance(replace, etree._Element):
                                log.debug("Will replace with XML CODE")
                            elif isinstance(replace(list, tuple)):
                                log.debug("Will replace with LIST OF ELEMENTS")
                            else:
                                log.debug("Will replace with:", re.sub(search,
                                                        replace, txtsearch))

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
                            p = org_first_r.getparent()
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
                                    new_r.append(makeelement('t', tagtext=line))
                                    new_r.append(makeelement('br'))
                                    p.insert(insindex, new_r)
                                    insindex += 1
                                # the last line of the segment
                                new_r = deepcopy(template_r)
                                new_r.append(makeelement('t', tagtext=lines[-1]))
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

    # check if used ..
    def _get_etree(self, xmldoc):
        return etree.fromstring(self._docx.read(xmldoc))

    def _load_etree(self, name, xmldoc):
        setattr(self, name, self._get_etree(xmldoc))

    def template(self, cx, max_blocks=5, raw_document=False):
        """
        Accepts a context dictionary (cx) and looks for the dict keys wrapped
        in {{key}}. Replaces occurances with the corresponding value from the
        cx dictionary.

        example:
            with the context...
                cx = {
                    'name': 'James',
                    'lang': 'English'
                }

            ...and a docx file containing:

                Hi! My name is {{name}} and I speak {{lang}}

            Calling `docx.template(cx)` will return a new docx instance (the
            original is not modified) that looks like:

                Hi! My name is James and I speak English

            Note: the template must not have spaces in the curly braces unless
            the dict key does (i.e., `{{ name }}` will not work unless your
            dictionary has `{" name ": ...}`)

        The `raw_document` argument accepts a boolean, which (if True) will
        treat the word/document.xml file as a text template (rather than only
        replacing text that is visible in the document via a word processor)

        If you pass `max_blocks=None` you will cause the template function to
        use `docx.replace()` rather than `docx.advanced_replace()`.

        When `max_blocks` is a number, it is passed to the advReplace
        method as is.
        """
        output = self.copy()

        if raw_document:
            raw_doc = etree.tostring(output.document)

        for key, val in cx.items():
            key = "{{%s}}" % key
            if raw_document:
                raw_doc = raw_doc.replace(key, unicode(val))
            elif max_blocks is None:
                output.replace(key, unicode(val))
            else:
                output.advReplace(key, val, max_blocks=max_blocks)

        if raw_document:
            output.document = etree.fromstring(raw_doc)

        return output

    #need to be rewrite with django/pagesettings see comments
    def save(self, dest=None):
        # pre-save ops
        self.modified = datetime.utcnow()
        self._apply_page_settings()

        # prepare our file
        outf = NamedTemporaryFile()
        out_zip = ZipFile(outf.name, mode='w')

        orig_contents = self._docx.namelist()
        modified_contents = self.trees_and_files.values()

        # Serialize our trees into our zip file
        for tree, dest_file in self.trees_and_files.items():
            log.info('Saving: ' + dest_file)
            out_zip.writestr(dest_file, etree.tostring(getattr(self, tree),
                                                            pretty_print=True))

        for dest_file in set(orig_contents) - set(modified_contents):
            out_zip.writestr(dest_file, self._docx.read(dest_file))


        out_zip.close()

        if dest is not None:
            log.info('Saved new file to: %r', dest)
            shutil.copyfile(outf.name, dest)
            outf.close()
        else:
            log.info('File saved')
            self._docx.close()
            shutil.copyfile(outf.name, self._tmp_file.name)

            # reopen the file so it can continue to be used
            self._docx = ZipFile(self._tmp_file.name, mode='a')

        # for 'djangodoc' style, need to return the file object
        #return outf.name

    # check if used ..
    def copy(self):
        tmp = NamedTemporaryFile()
        self.save(tmp.name)
        docx = self.__class__(tmp.name)
        docx._orig_docx = self._orig_docx
        tmp.close()
        return docx

    # check if used ..
    def __del__(self):
        try:
            self.__empty_docx.close()
        except AttributeError:
            pass
        self._docx.close()
        self._tmp_file.close()

    def _get_coreprop(self, tagname):
        try:
            return self.coreprops.xpath("*[local-name()='%s']" % tagname)[0]
        except IndexError:
            return None

    def _get_coreprop_val(self, tagname):
        return self._get_coreprop(tagname).text

    def _set_coreprop_val(self, tagname, val):
        self._get_coreprop(tagname).text = val

    # check if used ..
    def __generate_empty_docx(self):
        self.__empty_docx = NamedTemporaryFile()
        loc = self.__empty_docx.name
        dir_to_docx(TEMPLATE_DIR, loc)
        return loc


        dir_to_docx(TEMPLATE_DIR, loc)

        return loc

    # check if used ..
    @property
    def text(self):
        '''Return the raw text of a document, as a list of paragraphs.'''
        document = self.docbody
        paratextlist = []
        # Compile a list of all paragraph (p) elements
        paralist = []
        for element in document.iter():
            # Find p (paragraph) elements
            if element.tag == '{' + nsprefixes['w'] + '}p':
                paralist.append(element)
        # Since a single sentence might be spread over multiple text elements, iterate through each
        # paragraph, appending all text (t) children to that paragraphs text.
        for para in paralist:
            paratext = u''
            # Loop through each paragraph
            for element in para.iter():
                # Find t (text) elements
                if element.tag == '{' + nsprefixes['w'] + '}t':
                    if element.text:
                        paratext = paratext + element.text
            # Add our completed paragraph text to the list of paragraph text
            if not len(paratext) == 0:
                paratextlist.append(paratext)
        return paratextlist


""" old version keeped for compatibility """

def opendocx(mydocfile):
    '''Open a docx file, return a document XML tree'''
    doc = Docx(mydocfile)
    return doc.document

def opendocxheader(mydocfile):
    '''Open a docx file, return the header XML tree'''
    doc = Docx(mydocfile)
    return doc.header

def newdocument():
    doc = Docx()
    return doc.document

def createdoc(document, coreprops, appprops, contenttypes, websettings,
                wordrelationships, output, settings=None, template_dir=None):

    '''Save a modified document'''
    if not template_dir:
        template_dir = TEMP_TEMPLATE_DIR
    assert os.path.isdir(template_dir)

    if not output:
        output = StringIO()

    docxfile = zipfile.ZipFile(output, mode='w', compression=zipfile.ZIP_DEFLATED)

    # Move to the template data path
    prev_dir = os.path.abspath('.') # save previous working dir
    os.chdir(template_dir)

    # Add page settings
    document = add_page_settings(document, settings)

    # Serialize our trees into out zip file

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

    log.info('Saved new file to : %r', output or 'Direct Output')
    os.chdir(prev_dir) # restore previous working dir
    docxfile.close()

    return docxfile, output

def savedocx(document, coreprops, appprops, contenttypes, websettings,
                wordrelationships, output, settings=None, template_dir=None):

    docxfile, output = createdoc(document, coreprops, appprops, contenttypes,
                websettings, wordrelationships, output, settings, template_dir)

    return docxfile

def djangodocx(document, coreprops, appprops, contenttypes, websettings,
                wordrelationships, output=None, settings=None, template_dir=None):

    docxfile, output = createdoc(document, coreprops, appprops, contenttypes,
                websettings, wordrelationships, output, settings, template_dir)

    docx_zip = output.getvalue()
    output.close()

    return docx_zip

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

def getRelationships(file=None, prefix="document"):
    '''Get a Word relationships file from a given Word document'''
    if file and os.path.isfile(file):
        mydoc = zipfile.ZipFile(file)
        xmlcontent = mydoc.read('word/_rels/{0}.xml.rels'.format(prefix))
        relationships = etree.fromstring(xmlcontent)
        return relationships
    else:
        return getDefaultRelationships()

def add_page_settings(document, settings):
    ''' Add custom page settings '''
    sectPr = makeelement('sectPr')

    if not settings:
        settings = PAGESETTINGS

    for settingname, settingattrs in settings.iteritems():
        sectPr.append( makeelement(settingname, attributes=settingattrs))

    docbody = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    docbody.append(sectPr)

    return document
