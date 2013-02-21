<<<<<<< HEAD
<<<<<<< HEAD
import collections
import os
import zipfile
import shutil
import re

from docx import FILES_TO_IGNORE, NSPREFIXES
from docx.utils import make_element
from docx.meta import *


class DocxDocument(object):
    def __init__(self, template_file=None, template_dir=None):
        self.template_file = template_file
        self.template_dir = template_dir
        if self.template_file and os.path.isfile(self.template_file):
            self._init_from_file(self.template_file)
        else:
            self.template_file = None
            self.document = make_element("document")
            self.body = make_element("body")
            self.document.append(self.body)
            self.app_properties = AppProperties()
            self.core_properties = None
            self.word_relationships = WordRelationships()
            self.web_settings = WebSettings()
            self.content_types = ContentTypes()

    def _init_from_file(self, template_file):
        self.template_zip = zipfile.ZipFile(self.template_file, 'r', compression=zipfile.ZIP_DEFLATED)
        self.document = etree.fromstring(self.template_zip.read('word/document.xml'))
        self.word_relationships = WordRelationships(xml=self.template_zip.read('word/_rels/document.xml.rels'))
        #self.app_properties = AppProperties() # TODO: make available for manipulation
        #self.core_properties = CoreProperties() # TODO: make available for manipulation
        self.content_types = ContentTypes(xml=self.template_zip.read('[Content_Types].xml'))
        self.body = self.document.xpath('/w:document/w:body', namespaces=NSPREFIXES)[0]
=======
# encoding: utf-8

from namespaces import namespaces, content_types
from os import path

''' Open and modify Microsoft Word 2007 docx files (called 'OpenXML' 
and 'Office OpenXML' by Microsoft) '''

from lxml import etree
import Image
import zipfile
import shutil
import re
import time
import os
from copy import deepcopy
from os import path
from elements import Element

class Document(object):
    @classmethod
    def contenttypes(cls):
        # FIXME - doesn't quite work...read from string as temp hack...
        #types = Element('Types',nsprefix='ct')
        types = etree.fromstring('''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>''')
        for part in content_types:
            types.append(Element('Override',nsprefix=None,attributes={'PartName':part,'ContentType':parts[part]}))
        # Add support for filetypes
        for extension in file_types:
            types.append(Element('Default',nsprefix=None,attributes={'Extension':extension,'ContentType':filetypes[extension]}))
        return types

    def __init__(self, path=None):
        if path:
            self._open(path)
        else:
            self._create()
    
    def _open(self, path):
        '''Open a docx file, return a document XML tree'''
        self.path = path
        mydoc = zipfile.ZipFile(path)
        xmlcontent = mydoc.read('word/document.xml')
        self.xml = etree.fromstring(xmlcontent)
    
    def _create(self):
        xml = Element('document')
        xml.append(Element('body'))
        self.xml = xml

    def save(self, coreprops, appprops, contenttypes, websettings, wordrelationships, path):
        '''Save a modified document'''
        assert os.path.isdir(template_dir)
        docxfile = zipfile.ZipFile(path,mode='w',compression=zipfile.ZIP_DEFLATED)
        
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
            print 'Saving: '+treesandfiles[tree]    
            treestring = etree.tostring(tree, pretty_print=True)
            docxfile.writestr(treesandfiles[tree],treestring)
        
        # Add & compress support files
        files_to_ignore = ['.DS_Store'] # nuisance from OSX
        for dirpath,dirnames,filenames in os.walk('.'):
            for filename in filenames:
                if filename in files_to_ignore:
                    continue
                templatefile = path.join(dirpath, filename)
                archivename = templatefile[2:]
                print 'Saving: '+archivename          
                docxfile.write(templatefile, archivename)
        print 'Saved new file to: '+docxfilename
        os.chdir(prev_dir) # restore previous working dir
        
        self.path = path
>>>>>>> b0abe8d0c2ba44c3dd6675c5eb3198925a84431e

    def search(self, search):
        '''Search a document for a regex, return success / fail result'''
        result = False
        searchre = re.compile(search)
<<<<<<< HEAD
        for element in self.document.iter():
            if element.tag == '{%s}t' % NSPREFIXES['w']: # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        result = element
        return result

    def replace(self, search, replace):
        '''Replace all occurences of string with a different string, return updated document'''
        searchre = re.compile(search)
        for element in self.document.iter():
            if element.tag == '{%s}t' % NSPREFIXES['w']: # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        if isinstance(replace, str) or isinstance(replace, unicode):
                            element.text = re.sub(search,replace,element.text)
                        else:
                            parent = element.getparent()
                            parent.replace(element, replace)
                            #element.addnext(element)

    def add(self, element, position=None):
        if position:
            # TODO: enable adding stuff to at specific points of the text.
            pass
        else:
            self.body.append(element)

    def get_text(self):
=======
        for element in self.xml.iter():
            if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        result = True
        return result

    def clean(self):
        ''' Perform misc cleaning operations on documents.
        Returns cleaned document. '''
                
        newdocument = deepcopy(self)
        
        # Clean empty text and r tags
        for t in ('t', 'r'):
            rmlist = []
            for element in newdocument.xml.iter():
                if element.tag == '{%s}%s' % (nsprefixes['w'], t):
                    if not element.text and not len(element):
                        rmlist.append(element)
            for element in rmlist:
                element.getparent().remove(element)
        
        return newdocument

    def replace(self, search, replace):
        '''Replace all occurences of string with a different string, return updated document'''
        newdocument = deepcopy(self)
        searchre = re.compile(search)
        for element in newdocument.xml.iter():
            if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        element.text = re.sub(search,replace,element.text)
        
        return newdocument
    
    def advanced_replace(self, search, replace, bs=3):
        '''Replace all occurences of string with a different string, return updated document
        
        This is a modified version of python-docx.replace() that takes into
        account blocks of <bs> elements at a time. The replace element can also
        be a string or an xml etree element.
        
=======
#!/usr/bin/env python2.6
# -*- coding: utf-8 -*-
"""
Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and 'Office OpenXML' by Microsoft)

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

import logging
import re
from lxml.etree import fromstring, Element, _Element, tostring
try:
    from PIL import Image
except ImportError:
    import Image
import zipfile
import shutil
import time
from os import listdir, makedirs, mkdir, chdir, walk
from os.path import join, isdir, exists, dirname, abspath
from distutils.dir_util import copy_tree, remove_tree
from shutil import rmtree

log = logging.getLogger(__name__)

# Record template directory is docx/docx-template
BASE_TEMPLATE_DIR = join(dirname(__file__), 'docx-template')

# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily.
NSPREFIXES = {
    'mo': "http://schemas.microsoft.com/office/mac/office/2008/main",
    'o': 'urn:schemas-microsoft-com:office:office',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    # Text Content
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mv': "urn:schemas-microsoft-com:mac:vml",
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v': 'urn:schemas-microsoft-com:vml',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    # Properties (core and extended)
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    # Content Types (we're just making up our own namespaces here to save time)
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships (we're just making up our own namespaces here to save time)
    'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
    # Dublin Core document properties
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'dcterms': 'http://purl.org/dc/terms/',
    # New Document Terms
    'wpc': "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    'mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
    'wp14': "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    'w14': "http://schemas.microsoft.com/office/word/2010/wordml",
    'wpg': "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    'wpi': "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    'wps': "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}


class Document:
    def __init__(self, workFolder):
        if not exists(workFolder):
            makedirs(workFolder)
        assert isdir(workFolder)
        if listdir(workFolder):
            remove_tree(workFolder)
            makedirs(workFolder)
        copy_tree(BASE_TEMPLATE_DIR, workFolder)
        self.workFolder = workFolder
        self.new()

    def advReplace(self, document, search, replace, bs=3):
        """
        Replace all occurences of string with a different string, return updated document

        This is a modified version of python-docx.replace() that takes into
        account blocks of <bs> elements at a time. The replace element can also
        be a string or an xml etree element.

>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
        What it does:
        It searches the entire document body for text blocks.
        Then scan thos text blocks for replace.
        Since the text to search could be spawned across multiple text blocks,
        we need to adopt some sort of algorithm to handle this situation.
        The smaller matching group of blocks (up to bs) is then adopted.
        If the matching group has more than one block, blocks other than first
        are cleared and all the replacement text is put on first block.
<<<<<<< HEAD
        
=======

>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
        Examples:
        original text blocks : [ 'Hel', 'lo,', ' world!' ]
        search / replace: 'Hello,' / 'Hi!'
        output blocks : [ 'Hi!', '', ' world!' ]
<<<<<<< HEAD
        
        original text blocks : [ 'Hel', 'lo,', ' world!' ]
        search / replace: 'Hello, world' / 'Hi!'
        output blocks : [ 'Hi!!', '', '' ]
        
        original text blocks : [ 'Hel', 'lo,', ' world!' ]
        search / replace: 'Hel' / 'Hal'
        output blocks : [ 'Hal', 'lo,', ' world!' ]
        
=======

        original text blocks : [ 'Hel', 'lo,', ' world!' ]
        search / replace: 'Hello, world' / 'Hi!'
        output blocks : [ 'Hi!!', '', '' ]

        original text blocks : [ 'Hel', 'lo,', ' world!' ]
        search / replace: 'Hel' / 'Hal'
        output blocks : [ 'Hal', 'lo,', ' world!' ]

>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
        @param instance  document: The original document
        @param str       search: The text to search for (regexp)
        @param mixed replace: The replacement text or lxml.etree element to
                              append, or a list of etree elements
        @param int       bs: See above
<<<<<<< HEAD
        
        @return instance The document with replacement applied
        
        '''
        # Enables debug output
        DEBUG = False
        
        newdocument = deepcopy(self)
        
        # Compile the search regexp
        searchre = re.compile(search)
        
        # Will match against searchels. Searchels is a list that contains last
        # n text elements found in the document. 1 < n < bs
        searchels = []
        
        for element in newdocument.xml.iter():
            if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
=======

        @return instance The document with replacement applied

        """
        # Enables debug output
        DEBUG = False

        newdocument = document

        # Compile the search regexp
        searchre = re.compile(search)

        # Will match against searchels. Searchels is a list that contains last
        # n text elements found in the document. 1 < n < bs
        searchels = []

        for element in newdocument.iter():
            if element.tag == '{%s}t' % NSPREFIXES['w']:  # t (text) elements
>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
                if element.text:
                    # Add this element to searchels
                    searchels.append(element)
                    if len(searchels) > bs:
                        # Is searchels is too long, remove first elements
                        searchels.pop(0)
<<<<<<< HEAD
                    
=======

>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
                    # Search all combinations, of searchels, starting from
                    # smaller up to bigger ones
                    # l = search lenght
                    # s = search start
                    # e = element IDs to merge
                    found = False
<<<<<<< HEAD
                    for l in range(1,len(searchels)+1):
                        if found:
                            break
                        #print "slen:", l
                        for s in range(len(searchels)):
                            if found:
                                break
                            if s+l <= len(searchels):
                                e = range(s,s+l)
=======
                    for l in range(1, len(searchels) + 1):
                        if found:
                            break
                            #print "slen:", l
                        for s in range(len(searchels)):
                            if found:
                                break
                            if s + l <= len(searchels):
                                e = range(s, s + l)
>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
                                #print "elems:", e
                                txtsearch = ''
                                for k in e:
                                    txtsearch += searchels[k].text
<<<<<<< HEAD
                    
=======

>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
                                # Searcs for the text in the whole txtsearch
                                match = searchre.search(txtsearch)
                                if match:
                                    found = True
<<<<<<< HEAD
                                    
                                    # I've found something :)
                                    if DEBUG:
                                        print "Found element!"
                                        print "Search regexp:", searchre.pattern
                                        print "Requested replacement:", replace
                                        print "Matched text:", txtsearch
                                        print "Matched text (splitted):", map(lambda i:i.text,searchels)
                                        print "Matched at position:", match.start()
                                        print "matched in elements:", e
                                        if isinstance(replace, etree._Element):
                                            print "Will replace with XML CODE"
                                        elif type(replace) == list or type(replace) == tuple:
                                            print "Will replace with LIST OF ELEMENTS"
                                        else:
                                            print "Will replace with:", re.sub(search,replace,txtsearch)
    
=======

                                    # I've found something :)
                                    if DEBUG:
                                        log.debug("Found element!")
                                        log.debug("Search regexp: %s", searchre.pattern)
                                        log.debug("Requested replacement: %s", replace)
                                        log.debug("Matched text: %s", txtsearch)
                                        log.debug("Matched text (splitted): %s", map(lambda i: i.text, searchels))
                                        log.debug("Matched at position: %s", match.start())
                                        log.debug("matched in elements: %s", e)
                                        if isinstance(replace, _Element):
                                            log.debug("Will replace with XML CODE")
                                        elif isinstance(replace, (list, tuple)):
                                            log.debug("Will replace with LIST OF ELEMENTS")
                                        else:
                                            log.debug("Will replace with:", re.sub(search, replace, txtsearch))

>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
                                    curlen = 0
                                    replaced = False
                                    for i in e:
                                        curlen += len(searchels[i].text)
                                        if curlen > match.start() and not replaced:
                                            # The match occurred in THIS element. Puth in the
                                            # whole replaced text
<<<<<<< HEAD
                                            if isinstance(replace, etree._Element):
                                                # If I'm replacing with XML, clear the text in the
                                                # tag and append the element
                                                searchels[i].text = re.sub(search,'',txtsearch)
                                                searchels[i].append(replace)
                                            elif type(replace) == list or type(replace) == tuple:
                                                # I'm replacing with a list of etree elements
                                                searchels[i].text = re.sub(search,'',txtsearch)
                                                for r in replace:
                                                    searchels[i].append(r)
                                            else:
                                                # Replacing with pure text
                                                searchels[i].text = re.sub(search,replace,txtsearch)
                                            replaced = True
                                            if DEBUG:
                                                print "Replacing in element #:", i
=======
                                            if isinstance(replace, _Element):
                                                # Convert to a list and process it later
                                                replace = [replace, ]
                                            if isinstance(replace, (list, tuple)):
                                                # I'm replacing with a list of etree elements
                                                # clear the text in the tag and append the element after the
                                                # parent paragraph
                                                # (because t elements cannot have childs)
                                                p = self.findTypeParent(searchels[i], '{%s}p' % NSPREFIXES['w'])
                                                searchels[i].text = re.sub(search, '', txtsearch)
                                                insindex = p.getparent().index(p) + 1
                                                for r in replace:
                                                    p.getparent().insert(insindex, r)
                                                    insindex += 1
                                            else:
                                                # Replacing with pure text
                                                searchels[i].text = re.sub(search, replace, txtsearch)
                                            replaced = True
                                            log.debug("Replacing in element #: %s", i)
>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
                                        else:
                                            # Clears the other text elements
                                            searchels[i].text = ''
        return newdocument
<<<<<<< HEAD
    
    @property
    def plain(self):
>>>>>>> b0abe8d0c2ba44c3dd6675c5eb3198925a84431e
        '''Return the raw text of a document, as a list of paragraphs.'''
        paratextlist=[]   
        # Compile a list of all paragraph (p) elements
        paralist = []
<<<<<<< HEAD
        for element in self.document.iter():
            # Find p (paragraph) elements
            if element.tag == '{'+NSPREFIXES['w']+'}p':
=======
        for element in self.xml.iter():
            # Find p (paragraph) elements
            if element.tag == '{'+namespaces['w']+'}p':
>>>>>>> b0abe8d0c2ba44c3dd6675c5eb3198925a84431e
                paralist.append(element)    
        # Since a single sentence might be spread over multiple text elements, iterate through each 
        # paragraph, appending all text (t) children to that paragraphs text.     
        for para in paralist:      
<<<<<<< HEAD
            paratext = u''  
            # Loop through each paragraph
            for element in para.iter():
                # Find t (text) elements
                if element.tag == '{'+NSPREFIXES['w']+'}t':
=======
            paratext=u''  
            # Loop through each paragraph
            for element in para.iter():
                # Find t (text) elements
                if element.tag == '{'+namespaces['w']+'}t':
>>>>>>> b0abe8d0c2ba44c3dd6675c5eb3198925a84431e
                    if element.text:
                        paratext = paratext+element.text
            # Add our completed paragraph text to the list of paragraph text    
            if not len(paratext) == 0:
<<<<<<< HEAD
                paratextlist.append(paratext)                    
        return paratextlist

    def append(self, element):
        self.document.append(element)

#    def _clean(self):
#        """ Perform misc cleaning operations on documents.
#            Returns cleaned document.
#        """
#        # Clean empty text and r tags
#        for t in ('t', 'r'):
#            rmlist = []
#            for element in self.document.iter():
#                if element.tag == '{%s}%s' % (NSPREFIXES['w'], t):
#                    if not element.text and not len(element):
#                        rmlist.append(element)
#            for element in rmlist:
#                element.getparent().remove(element)
#        return newdocument

    def _write_xml_files(self):
        # Serialize our trees into out zip file
        files = {
            self.core_properties: 'docProps/core.xml',
            self.app_properties: 'docProps/app.xml',
            self.content_types:'[Content_Types].xml',
            self.web_settings:'word/webSettings.xml',
            self.word_relationships:'word/_rels/document.xml.rels'
        }
        for f in files:
            treestring = etree.tostring(f._xml(), pretty_print=True)
            self.zip_file.writestr(files[f],treestring)

    def _copy_template_dir(self):
        """Copy a template document to our container."""
        for (dirpath, dirnames, filenames) in os.walk(self.template_dir):
            for filename in filenames:
                if filename in FILES_TO_IGNORE:
                    continue
                path = os.path.join(dirpath, filename)
                self.zip_file.write(path, os.path.relpath(path, self.template_dir))

    def _copy_template_file(self):
        """ Copy contents of template docx file into new docx file """
        for filename in self.template_zip.namelist():
            if os.path.basename(filename) in ['document.xml', 'document.xml.rels', '[Content_Types].xml']:
                continue
            self.zip_file.writestr(filename, self.template_zip.read(filename))

    def _copy_media_files(self):
        for name, path in self.word_relationships.to_copy:
            out = 'word/media/' + name # IN DESPERATE NEED OF A FIX
            self.zip_file.write(path, out)

    def save(self, filename):
        '''Save a modified document'''
        self.zip_file = zipfile.ZipFile(filename, mode='w', compression=zipfile.ZIP_DEFLATED)
        # TODO: determine what to do when template_file AND template_dir are specified
        if self.template_dir:
            self._write_xml_files()
            self._copy_template_dir()
        if self.template_file:
            self._copy_template_file()
            self.zip_file.writestr('word/_rels/document.xml.rels',
                                    etree.tostring(self.word_relationships._xml(),
                                    pretty_print=True))
            self.zip_file.writestr('[Content_Types].xml',
                                    etree.tostring(self.content_types._xml(),
                                    pretty_print=True))

        # Copying over any newly added media files.
        self._copy_media_files()
        # Adding the content file.
        self.zip_file.writestr('word/document.xml', etree.tostring(self.document, pretty_print=True))
        
        #self.zip_file.close()
=======
                paratextlist.append(paratext)             
               
        return paratextlist
    
    # collapse fancy structures, and give output like ("heading", "This was a <strong>lovely</strong> day.")
    # (if MS uses different tags for strong, em and links, we can do that transformation in a to_html 
    # method or something)
    def fold(self):
        raise NotImplementedError()
>>>>>>> b0abe8d0c2ba44c3dd6675c5eb3198925a84431e
=======

    def advSearch(self, document, search, bs=3):
        """Return set of all regex matches

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

        @param document: The original document
        @param search: The text to search for (regexp), append, or a list of etree elements
        @param bs: See above
        @return set of All occurences of search string

        """
        # Compile the search regexp
        searchre = re.compile(search)

        matches = []

        # Will match against searchels. Searchels is a list that contains last
        # n text elements found in the document. 1 < n < bs
        searchels = []

        for element in document.iter():
            if element.tag == '{%s}t' % NSPREFIXES['w']:  # t (text) elements
                if element.text:
                    # Add this element to searchels
                    searchels.append(element)
                    if len(searchels) > bs:
                        # Is searchels is too long, remove first elements
                        searchels.pop(0)

                    # Search all combinations, of searchels, starting from
                    # smaller up to bigger ones
                    # l = search lenght
                    # s = search start
                    # e = element IDs to merge
                    found = False
                    for l in range(1, len(searchels) + 1):
                        if found:
                            break
                        for s in range(len(searchels)):
                            if found:
                                break
                            if s + l <= len(searchels):
                                e = range(s, s + l)
                                txtsearch = ''
                                for k in e:
                                    txtsearch += searchels[k].text

                                # Searcs for the text in the whole txtsearch
                                match = searchre.search(txtsearch)
                                if match:
                                    matches.append(match.group())
                                    found = True

        return set(matches)

    def appProperties(self):
        """
        Create app-specific properties. See docproperties() for more common
        document properties.
        """
        # appprops = self.makeElement('Properties', nsprefix='ep')
        appprops = fromstring(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/'
            '2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.or'
            'g/officeDocument/2006/docPropsVTypes"></Properties>')
        props = {
            'Template': 'Normal.dotm',
            'TotalTime': '6',
            'Pages': '1',
            'Words': '83',
            'Characters': '475',
            'Application': 'Microsoft Word 12.0.0',
            'DocSecurity': '0',
            'Lines': '12',
            'Paragraphs': '8',
            'ScaleCrop': 'false',
            'LinksUpToDate': 'false',
            'CharactersWithSpaces': '583',
            'SharedDoc': 'false',
            'HyperlinksChanged': 'false',
            'AppVersion': '12.0000'
        }
        for prop in props:
            appprops.append(self.makeElement(prop, tagtext=props[prop], nsprefix=None))
        return appprops

    def clean(self):
        """ Perform misc cleaning operations on document."""
        newdocument = self.doc
        # Clean empty text and r tags
        for t in ('t', 'r'):
            rmlist = []
            for element in newdocument.iter():
                if element.tag == '{%s}%s' % (NSPREFIXES['w'], t):
                    if not element.text and not len(element):
                        rmlist.append(element)
            for element in rmlist:
                element.getparent().remove(element)
        self.doc = newdocument

    def contentTypes(self):
        # FIXME - doesn't quite work...read from string as temp hack...
        #types = self.makeElement('Types', nsprefix='ct')
        types = fromstring('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>')
        parts = {
            '/word/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml',
            '/word/fontTable.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
            '/docProps/core.xml': 'application/vnd.openxmlformats-package.core-properties+xml',
            '/docProps/app.xml': 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
            '/word/document.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
            '/word/settings.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
            '/word/numbering.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
            '/word/styles.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
            '/word/webSettings.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml'
        }
        for part in parts:
            types.append(self.makeElement('Override', nsprefix=None, attributes={'PartName': part,
                                                                                 'ContentType': parts[part]}))
            # Add support for filetypes
        filetypes = {
            'gif': 'image/gif',
            'jpeg': 'image/jpeg',
            'jpg': 'image/jpeg',
            'png': 'image/png',
            'rels': 'application/vnd.openxmlformats-package.relationships+xml',
            'xml': 'application/xml'
        }
        for extension in filetypes:
            types.append(self.makeElement('Default', nsprefix=None, attributes={'Extension': extension,
                                                                                'ContentType': filetypes[extension]}))
        return types

    def coreProperties(self, title, subject, creator, keywords, lastModifiedBy=None):
        """Create core properties (common document properties referred to in the 'Dublin Core' specification).
        See appproperties() for other stuff."""
        coreprops = self.makeElement('coreProperties', nsprefix='cp')
        coreprops.append(self.makeElement('title', tagtext=title, nsprefix='dc'))
        coreprops.append(self.makeElement('subject', tagtext=subject, nsprefix='dc'))
        coreprops.append(self.makeElement('creator', tagtext=creator, nsprefix='dc'))
        coreprops.append(self.makeElement('keywords', tagtext=','.join(keywords), nsprefix='cp'))
        if not lastModifiedBy:
            lastModifiedBy = creator
        coreprops.append(self.makeElement('lastModifiedBy', tagtext=lastModifiedBy, nsprefix='cp'))
        coreprops.append(self.makeElement('revision', tagtext='1', nsprefix='cp'))
        coreprops.append(self.makeElement('category', tagtext='Examples', nsprefix='cp'))
        coreprops.append(self.makeElement('description', tagtext='Examples', nsprefix='dc'))
        currenttime = time.strftime('%Y-%m-%dT%H:%M:%SZ')
        # Document creation and modify times
        # Prob here: we have an attribute who name uses one namespace, and that
        # attribute's value uses another namespace.
        # We're creating the element from a string as a workaround...
        for doctime in ['created', 'modified']:
            coreprops.append(fromstring(
                """<dcterms:""" + doctime + """ xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" """ +
                """xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">""" + currenttime +
                """</dcterms:""" + doctime + """>"""))
            pass
        self.coreProps = coreprops

    def findTypeParent(self, element, tag):
        """ Finds fist parent of element of the given type

        @param element: etree element
        @param tag: string the tag parent to search for

        @return element: the found parent or None when not found
        """
        p = element
        while True:
            p = p.getparent()
            if p.tag == tag:
                return p

        # Not found
        return None

    def getDocumentText(self):
        """Return the raw text of a document, as a list of paragraphs."""
        paratextlist = []
        # Compile a list of all paragraph (p) elements
        paralist = []
        for element in self.doc.iter():
            # Find p (paragraph) elements
            if element.tag == '{' + NSPREFIXES['w'] + '}p':
                paralist.append(element)
                # Since a single sentence might be spread over multiple text elements, iterate through each
            # paragraph, appending all text (t) children to that paragraphs text.
        for para in paralist:
            paratext = u''
            # Loop through each paragraph
            for element in para.iter():
                # Find t (text) elements
                if element.tag == '{' + NSPREFIXES['w'] + '}t':
                    if element.text:
                        paratext += element.text
                elif element.tag == '{' + NSPREFIXES['w'] + '}tab':
                    paratext += '\t'
                    # Add our completed paragraph text to the list of paragraph text
            if not len(paratext) == 0:
                paratextlist.append(paratext)
        return paratextlist

    def heading(self, headingText, headingLevel, lang='en', jc='left'):
        """Make a new heading, return the heading element"""
        lmap = {'en': 'Heading', 'it': 'Titolo'}
        # Make our elements
        paragraph = self.makeElement('p')
        pr = self.makeElement('pPr')
        pStyle = self.makeElement('pStyle', attributes={'val': lmap[lang] + str(headingLevel)})
        pJc = self.makeElement('jc', attributes={'val': jc})
        run = self.makeElement('r')
        text = self.makeElement('t', tagtext=headingText)
        # Add the text the run, and the run to the paragraph
        pr.append(pStyle)
        pr.append(pJc)
        run.append(text)
        paragraph.append(pr)
        paragraph.append(run)
        # Return the combined paragraph
        self.body.append(paragraph)

    def makeElement(self, tagname, tagtext=None, nsprefix='w', attributes=None, attrnsprefix=None):
        """Create an element & return it"""
        # Deal with list of nsprefix by making namespacemap
        # namespacemap = None
        attrib = None
        if nsprefix == 'w':
            attrib = {'{http://www.w3.org/XML/1998/namespace}space': "preserve"}
        if isinstance(nsprefix, list):
            namespacemap = {}
            for prefix in nsprefix:
                namespacemap[prefix] = NSPREFIXES[prefix]
            nsprefix = nsprefix[0]  # FIXME: rest of code below expects a single prefix
        if nsprefix:
            namespace = '{' + NSPREFIXES[nsprefix] + '}'
        else:
            # For when namespace = None
            namespace = ''
        log.debug(namespace)
        #newelement = Element(namespace + tagname, nsmap=namespacemap)
        newelement = Element(namespace + tagname, attrib=attrib, nsmap=NSPREFIXES)
        # Add attributes with namespaces
        if attributes:
            # If they haven't bothered setting attribute namespace, use an empty string
            # (equivalent of no namespace)
            if not attrnsprefix:
                # Quick hack: seems all elements with 'w' nsprefix for its tag uses the same prefix for it's attributes
                if nsprefix == 'w':
                    attributenamespace = namespace
                else:
                    attributenamespace = ''
            else:
                attributenamespace = '{' + NSPREFIXES[attrnsprefix] + '}'

            for tagattribute in attributes:
                newelement.set(attributenamespace + tagattribute, attributes[tagattribute])
        if tagtext:
            newelement.text = tagtext
        return newelement

    def new(self):
        self.doc = self.makeElement('document')
        self.doc.append(self.makeElement('body'))
        self.body = self.doc.xpath('/w:document/w:body', namespaces=NSPREFIXES)[0]
        self.relationships = self.relationshipList()
        self.appProps = self.appProperties()
        self.cTypes = self.contentTypes()
        self.web = self.webSettings()
        self.coreProperties(title='', subject='', creator='', keywords='')

    def open(self, filename):
        """Open a docx file, return a document XML tree"""
        document = zipfile.ZipFile(filename)
        xmlContent = document.read('word/document.xml')
        self.doc = fromstring(xmlContent)
        self.body = self.doc.xpath('/w:document/w:body', namespaces=NSPREFIXES)[0]

    def pageBreak(self, pType='page', orient='portrait'):
        """Insert a break, default 'page'.
        See http://openxmldeveloper.org/forums/thread/4075.aspx
        Return our page break element."""
        # Need to enumerate different types of page breaks.
        validtypes = ['page', 'section']
        if pType not in validtypes:
            raise ValueError('Page break style "%s" not implemented. Valid styles: %s.' % (pType, validtypes))
        pagebreak = self.makeElement('p')
        if pType == 'page':
            run = self.makeElement('r')
            br = self.makeElement('br', attributes={'type': pType})
            run.append(br)
            pagebreak.append(run)
        elif pType == 'section':
            pPr = self.makeElement('pPr')
            sectPr = self.makeElement('sectPr')
            if orient == 'landscape':
                pgSz = self.makeElement('pgSz', attributes={'h': '12240', 'w': '15840', 'orient': 'landscape'})
            else:  # Assume orient == 'portrait':
                pgSz = self.makeElement('pgSz', attributes={'w': '12240', 'h': '15840'})
            sectPr.append(pgSz)
            pPr.append(sectPr)
            pagebreak.append(pPr)
        self.body.append(pagebreak)

    def paragraph(self, paraText, style='BodyText', breakBefore=False, jc='left'):
        """Make a new paragraph element, containing a run, and some text.
        Return the paragraph element.

        @param string jc: Paragraph alignment, possible values:
                          left, center, right, both (justified), ...
                          see http://www.schemacentral.com/sc/ooxml/t-w_ST_Jc.html
                          for a full list

        If paratext is a list, spawn multiple run/text elements.
        Support text styles (paratext must then be a list of lists in the form
        <text> / <style>. Stile is a string containing a combination od 'bui' chars

        example
        paratext =\
            [ ('some bold text', 'b')
            , ('some normal text', '')
            , ('some italic underlined text', 'iu')
            ]

        """
        # Make our elements
        paragraph = self.makeElement('p')

        if isinstance(paraText, list):
            text = []
            for pt in paraText:
                if isinstance(pt, (list, tuple)):
                    items = self.splitNewlineTab(pt[0], pt[1])
                    for item in items:
                        text.append(item)
                else:
                    items = self.splitNewlineTab(pt, '')
                    for item in items:
                        text.append(item)
        else:
            text = self.splitNewlineTab(paraText, '')
        pPr = self.makeElement('pPr')
        pStyle = self.makeElement('pStyle', attributes={'val': style})
        pJc = self.makeElement('jc', attributes={'val': jc})
        pPr.append(pStyle)
        pPr.append(pJc)

        # Add the text the run, and the run to the paragraph
        paragraph.append(pPr)
        for t in text:
            run = self.makeElement('r')
            rPr = self.makeElement('rPr')
            # Apply styles
            if t[1].find('b') > -1:
                b = self.makeElement('b')
                rPr.append(b)
            if t[1].find('u') > -1:
                u = self.makeElement('u', attributes={'val': 'single'})
                rPr.append(u)
            if t[1].find('i') > -1:
                i = self.makeElement('i')
                rPr.append(i)
            run.append(rPr)
            # Insert lastRenderedPageBreak for assistive technologies like
            # document narrators to know when a page break occurred.
            if breakBefore:
                lastRenderedPageBreak = self.makeElement('lastRenderedPageBreak')
                run.append(lastRenderedPageBreak)

            run.append(t[0])
            paragraph.append(run)
            # Return the combined paragraph
        self.body.append(paragraph)

    def picture(self, picFilename, picDescription, pixelWidth=None, pixelHeight=None, noChangeAspect=True,
                noChangeArrowHeads=True):
        """Take a relationshiplist, picture file name, and return a paragraph containing the image
        and an updated relationshiplist"""
        # http://openxmldeveloper.org/articles/462.aspx
        # Create an image. Size may be specified, otherwise it will based on the
        # pixel size of image. Return a paragraph containing the picture"""
        # Copy the file into the media dir
        media_dir = join(self.workFolder, 'word', 'media')
        if not isdir(media_dir):
            mkdir(media_dir)

        picName = picFilename.split('/')[-1]
        shutil.copyfile(picFilename, join(media_dir, picName))

        # Check if the user has specified a size
        if not pixelWidth or not pixelHeight:
            # If not, get info from the picture itself
            pixelWidth, pixelHeight = Image.open(picFilename).size[0:2]

        # OpenXML measures on-screen objects in English Metric Units
        # 1cm = 36000 EMUs
        emuperpixel = 12667
        width = str(pixelWidth * emuperpixel)
        height = str(pixelHeight * emuperpixel)

        # Set relationship ID to the first available
        picid = '2'
        picrelid = 'rId' + str(len(self.relationships) + 1)
        self.relationships.append([
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            'media/' + picName])

        # There are 3 main elements inside a picture
        # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
        blipfill = self.makeElement('blipFill', nsprefix='pic')
        blipfill.append(self.makeElement('blip', nsprefix='a', attrnsprefix='r', attributes={'embed': picrelid}))
        stretch = self.makeElement('stretch', nsprefix='a')
        stretch.append(self.makeElement('fillRect', nsprefix='a'))
        blipfill.append(self.makeElement('srcRect', nsprefix='a'))
        blipfill.append(stretch)

        # 2. The non visual picture properties
        nvpicpr = self.makeElement('nvPicPr', nsprefix='pic')
        cnvpr = self.makeElement('cNvPr', nsprefix='pic',
                                 attributes={'id': '0', 'name': 'Picture 1', 'descr': picName})
        nvpicpr.append(cnvpr)
        cnvpicpr = self.makeElement('cNvPicPr', nsprefix='pic')
        cnvpicpr.append(self.makeElement('picLocks', nsprefix='a',
                                         attributes={'noChangeAspect': str(int(noChangeAspect)),
                                                     'noChangeArrowheads': str(int(noChangeArrowHeads))}))
        nvpicpr.append(cnvpicpr)

        # 3. The Shape properties
        sppr = self.makeElement('spPr', nsprefix='pic', attributes={'bwMode': 'auto'})
        xfrm = self.makeElement('xfrm', nsprefix='a')
        xfrm.append(self.makeElement('off', nsprefix='a', attributes={'x': '0', 'y': '0'}))
        xfrm.append(self.makeElement('ext', nsprefix='a', attributes={'cx': width, 'cy': height}))
        prstgeom = self.makeElement('prstGeom', nsprefix='a', attributes={'prst': 'rect'})
        prstgeom.append(self.makeElement('avLst', nsprefix='a'))
        sppr.append(xfrm)
        sppr.append(prstgeom)

        # Add our 3 parts to the picture element
        pic = self.makeElement('pic', nsprefix='pic')
        pic.append(nvpicpr)
        pic.append(blipfill)
        pic.append(sppr)

        # Now make the supporting elements
        # The following sequence is just: make element, then add its children
        graphicdata = self.makeElement('graphicData', nsprefix='a',
                                       attributes={'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'})
        graphicdata.append(pic)
        graphic = self.makeElement('graphic', nsprefix='a')
        graphic.append(graphicdata)

        framelocks = self.makeElement('graphicFrameLocks', nsprefix='a', attributes={'noChangeAspect': '1'})
        framepr = self.makeElement('cNvGraphicFramePr', nsprefix='wp')
        framepr.append(framelocks)
        docpr = self.makeElement('docPr', nsprefix='wp',
                                 attributes={'id': picid, 'name': 'Picture 1', 'descr': picDescription})
        effectextent = self.makeElement('effectExtent', nsprefix='wp',
                                        attributes={'l': '25400', 't': '0', 'r': '0', 'b': '0'})
        extent = self.makeElement('extent', nsprefix='wp', attributes={'cx': width, 'cy': height})
        inline = self.makeElement('inline',
                                  attributes={'distT': "0", 'distB': "0", 'distL': "0", 'distR': "0"}, nsprefix='wp')
        inline.append(extent)
        inline.append(effectextent)
        inline.append(docpr)
        inline.append(framepr)
        inline.append(graphic)
        drawing = self.makeElement('drawing')
        drawing.append(inline)
        run = self.makeElement('r')
        run.append(drawing)
        paragraph = self.makeElement('p')
        paragraph.append(run)
        self.body.append(paragraph)

    def relationshipList(self):
        relationshiplist = [
            ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', 'numbering.xml'],
            ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'styles.xml'],
            ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml'],
            ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings', 'webSettings.xml'],
            ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable', 'fontTable.xml'],
            ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'theme/theme1.xml']
        ]
        return relationshiplist

    def replace(self, document, search, replace):
        """Replace all occurences of string with a different string, return updated document"""
        newdocument = document
        searchre = re.compile(search)
        for element in newdocument.iter():
            if element.tag == '{%s}t' % NSPREFIXES['w']:  # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        element.text = re.sub(search, replace, element.text)
        return newdocument

    def save(self, output):
        """Save a modified document"""
        assert isdir(self.workFolder)
        docxfile = zipfile.ZipFile(output, mode='w', compression=zipfile.ZIP_DEFLATED)

        # Move to the template data path
        prev_dir = abspath('.')  # save previous working dir
        chdir(self.workFolder)
        word = self.wordRelationships(self.relationships)
        # Serialize our trees into out zip file
        treesandfiles = {self.doc: 'word/document.xml',
                         self.coreProps: 'docProps/core.xml',
                         self.appProps: 'docProps/app.xml',
                         self.cTypes: '[Content_Types].xml',
                         self.web: 'word/webSettings.xml',
                         word: 'word/_rels/document.xml.rels'}
        for tree in treesandfiles:
            log.info('Saving: ' + treesandfiles[tree])
            treestring = tostring(tree, pretty_print=True)
            #treestring = tostring(tree)
            docxfile.writestr(treesandfiles[tree], treestring)

        # Add & compress support files
        files_to_ignore = ['.DS_Store']  # nuisance from some os's
        for dirpath, dirnames, filenames in walk('.'):
            for filename in filenames:
                if filename in files_to_ignore:
                    continue
                templatefile = join(dirpath, filename)
                archivename = templatefile[2:]
                log.info('Saving: %s', archivename)
                docxfile.write(templatefile, archivename)
        log.info('Saved new file to: %r', output)
        docxfile.close()
        chdir(prev_dir)  # restore previous working dir
        return

    def search(self, search):
        """Search a document for a regex, return success / fail result"""
        result = False
        searchre = re.compile(search)
        for element in self.doc.iter():
            if element.tag == '{%s}t' % NSPREFIXES['w']:  # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        result = True
        return result

    def splitNewlineTab(self, t, opt):
        text = []
        last = 0
        for m in re.finditer(r"[\n\t]", t):
            if last != m.start():
                text.append([self.makeElement('t', tagtext=t[last:m.start()]), opt])
            if m.group(0) == '\n':
                text.append([self.makeElement('br'), opt])
            elif m.group(0) == '\t':
                text.append([self.makeElement('tab'), opt])
            last = m.end()
        if last != len(t):
            text.append([self.makeElement('t', tagtext=t[last:]), opt])
        return text

    def table(self, contents, heading=True, colw=None, cwunit='dxa', tblw=0, twunit='auto', borders=None,
              celstyle=None):
        """Get a list of lists, return a table

            @param list contents: A list of lists describing contents
                                  Every item in the list can be a string or a valid
                                  XML element itself. It can also be a list. In that case
                                  all the listed elements will be merged into the cell.
            @param bool heading: Tells whether first line should be threated as heading
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
                                 val: The style of the border:
                                      see http://www.schemacentral.com/sc/ooxml/t-w_ST_Border.htm
            @param list celstyle: Specify the style for each colum, list of dicts.
                                  supported keys:
                                  'align': specify the alignment, see paragraph documentation.

            @return lxml.etree: Generated XML etree element
        """
        if not borders:
            borders = {}
        table = self.makeElement('tbl')
        columns = len(contents[0])
        # Table properties
        tableprops = self.makeElement('tblPr')
        tablestyle = self.makeElement('tblStyle', attributes={'val': ''})
        tableprops.append(tablestyle)
        tablewidth = self.makeElement('tblW', attributes={'w': str(tblw), 'type': str(twunit)})
        tableprops.append(tablewidth)
        if len(borders.keys()):
            tableborders = self.makeElement('tblBorders')
            for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                if b in borders.keys() or 'all' in borders.keys():
                    k = 'all' if 'all' in borders.keys() else b
                    attrs = {}
                    for a in borders[k].keys():
                        attrs[a] = unicode(borders[k][a])
                    borderelem = self.makeElement(b, attributes=attrs)
                    tableborders.append(borderelem)
            tableprops.append(tableborders)
        tablelook = self.makeElement('tblLook', attributes={'val': '0400'})
        tableprops.append(tablelook)
        table.append(tableprops)
        # Table Grid
        tablegrid = self.makeElement('tblGrid')
        for i in range(columns):
            tablegrid.append(self.makeElement('gridCol', attributes={'w': str(colw[i]) if colw else '2390'}))
        table.append(tablegrid)
        # Heading Row
        row = self.makeElement('tr')
        rowprops = self.makeElement('trPr')
        cnfStyle = self.makeElement('cnfStyle', attributes={'val': '000000100000'})
        rowprops.append(cnfStyle)
        row.append(rowprops)
        if heading:
            i = 0
            for heading in contents[0]:
                cell = self.makeElement('tc')
                # Cell properties
                cellprops = self.makeElement('tcPr')
                if colw:
                    wattr = {'w': str(colw[i]), 'type': cwunit}
                else:
                    wattr = {'w': '0', 'type': 'auto'}
                cellwidth = self.makeElement('tcW', attributes=wattr)
                cellstyle = self.makeElement('shd', attributes={'val': 'clear', 'color': 'auto', 'fill': 'FFFFFF',
                                                                'themeFill': 'text2', 'themeFillTint': '99'})
                cellprops.append(cellwidth)
                cellprops.append(cellstyle)
                cell.append(cellprops)
                # Paragraph (Content)
                if not isinstance(heading, (list, tuple)):
                    heading = [heading]
                for h in heading:
                    if isinstance(h, _Element):
                        cell.append(h)
                    else:
                        cell.append(self.paragraph(h, jc='center'))
                row.append(cell)
                i += 1
            table.append(row)
            # Contents Rows
        for contentrow in contents[1 if heading else 0:]:
            row = self.makeElement('tr')
            i = 0
            for content in contentrow:
                cell = self.makeElement('tc')
                # Properties
                cellprops = self.makeElement('tcPr')
                if colw:
                    wattr = {'w': str(colw[i]), 'type': cwunit}
                else:
                    wattr = {'w': '0', 'type': 'auto'}
                cellwidth = self.makeElement('tcW', attributes=wattr)
                cellprops.append(cellwidth)
                cell.append(cellprops)
                # Paragraph (Content)
                if not isinstance(content, (list, tuple)):
                    content = [content]
                for c in content:
                    if isinstance(c, _Element):
                        cell.append(c)
                    else:
                        if celstyle and 'align' in celstyle[i].keys():
                            align = celstyle[i]['align']
                        else:
                            align = 'left'
                        cell.append(self.paragraph(c, jc=align))
                row.append(cell)
                i += 1
            table.append(row)
        self.body.append(table)

    def webSettings(self):
        """Generate websettings"""
        web = self.makeElement('webSettings')
        web.append(self.makeElement('allowPNG'))
        web.append(self.makeElement('doNotSaveAsSingleFile'))
        return web

    def wordRelationships(self, relationshipList):
        """Generate a Word relationships file"""
        # Default list of relationships
        # FIXME: using string hack instead of making element
        #relationships = self.makeElement('Relationships', nsprefix='pr')
        relationships = fromstring('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/' +
                                   'relationships"></Relationships>')
        count = 0
        for relationship in relationshipList:
            # Relationship IDs (rId) start at 1.
            relationships.append(self.makeElement('Relationship', attributes={
                'Id': 'rId' + str(count + 1),
                'Type': relationship[0], 'Target': relationship[1]
            }, nsprefix=None))
            count += 1
        return relationships
>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
