# encoding: utf-8

from namespaces import namespaces, content_types
from os import path

""" Open and modify Microsoft Word 2007 docx files (called 'OpenXML' 
and 'Office OpenXML' by Microsoft)"""

from lxml import etree
import Image
import zipfile
import shutil
import re
import time
import os
from copy import deepcopy
from os.path import join
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

    def __init__(path=None):
        if filename:
            self._open(path)
        else:
            self._create()
    
    def _open(self, path):
        '''Open a docx file, return a document XML tree'''
        self.path = path
        mydoc = zipfile.ZipFile(file)
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

    def search(self, search):
        '''Search a document for a regex, return success / fail result'''
        result = False
        searchre = re.compile(search)
        for element in self.xml.iter():
            if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        result = True
        return result

    def clean(self):
        """ Perform misc cleaning operations on documents.
            Returns cleaned document.
        """
        
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
        
        newdocument = deepcopy(self)
        
        # Compile the search regexp
        searchre = re.compile(search)
        
        # Will match against searchels. Searchels is a list that contains last
        # n text elements found in the document. 1 < n < bs
        searchels = []
        
        for element in newdocument.xml.iter():
            if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
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
                    for l in range(1,len(searchels)+1):
                        if found:
                            break
                        #print "slen:", l
                        for s in range(len(searchels)):
                            if found:
                                break
                            if s+l <= len(searchels):
                                e = range(s,s+l)
                                #print "elems:", e
                                txtsearch = ''
                                for k in e:
                                    txtsearch += searchels[k].text
                    
                                # Searcs for the text in the whole txtsearch
                                match = searchre.search(txtsearch)
                                if match:
                                    found = True
                                    
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
    
                                    curlen = 0
                                    replaced = False
                                    for i in e:
                                        curlen += len(searchels[i].text)
                                        if curlen > match.start() and not replaced:
                                            # The match occurred in THIS element. Puth in the
                                            # whole replaced text
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
                                        else:
                                            # Clears the other text elements
                                            searchels[i].text = ''
        return newdocument
    
    @property
    def plain(self):
        '''Return the raw text of a document, as a list of paragraphs.'''
        paratextlist=[]   
        # Compile a list of all paragraph (p) elements
        paralist = []
        for element in self.xml.iter():
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
            # Add our completed paragraph text to the list of paragraph text    
            if not len(paratext) == 0:
                paratextlist.append(paratext)             
               
        return paratextlist