# -*- coding: utf-8 -*-
import logging
import os
from os.path import join
import re
import shutil
from tempfile import NamedTemporaryFile
import time
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree
try:
    from PIL import Image
except ImportError:
    import Image
    
from utils import findTypeParent

log = logging.getLogger(__name__)

# Record template directory's location which is just 'template' for a docx
# developer or 'site-packages/docx-template' if you have installed docx
template_dir = join(os.path.dirname(__file__),'docx-template') # installed
if not os.path.isdir(template_dir):
    template_dir = join(os.path.dirname(__file__),'template') # dev

# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily.
nsprefixes = {
    # Text Content
    'mv':'urn:schemas-microsoft-com:mac:vml',
    'mo':'http://schemas.microsoft.com/office/mac/office/2008/main',
    've':'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o':'urn:schemas-microsoft-com:office:office',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v':'urn:schemas-microsoft-com:vml',
    'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10':'urn:schemas-microsoft-com:office:word',
    'wne':'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    # Properties (core and extended)
    'cp':"http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    'dc':"http://purl.org/dc/elements/1.1/",
    'dcterms':"http://purl.org/dc/terms/",
    'dcmitype':"http://purl.org/dc/dcmitype/",
    'xsi':"http://www.w3.org/2001/XMLSchema-instance",
    'ep':'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    # Content Types (we're just making up our own namespaces here to save time)
    'ct':'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships (we're just making up our own namespaces here to save time)
    'pr':'http://schemas.openxmlformats.org/package/2006/relationships'
    }
    
class Docx(object):
    trees_and_files = {
        "document": 'word/document.xml',
        "coreprops":'docProps/core.xml',
        "appprops":'docProps/app.xml',
        "contenttypes":'[Content_Types].xml',
        "websettings":'word/webSettings.xml',
        "wordrelationships":'word/_rels/document.xml.rels'
    }
    def __init__(self, file=None):
        self._orig_docx = file
        self._tmp_file = NamedTemporaryFile()
        
        if file is None:
            file = self.__generate_empty_docx()
        
        shutil.copyfile(file, self._tmp_file.name)
        self._docx = ZipFile(self._tmp_file.name, mode='a')
        
        for tree, file in self.trees_and_files.items():
            self._load_etree(tree, file)
            
    def append(self, *args, **kwargs):
        return self.document.append(*args, **kwargs)
    
    def search(self, search):
        '''Search a document for a regex, return success / fail result'''
        document = self.document
        
        result = False
        searchre = re.compile(search)
        for element in document.iter():
            if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        result = True
        return result
    
    def replace(self, search, replace):
        '''Replace all occurences of string with a different string, return updated document'''
        newdocument = self.document
        searchre = re.compile(search)
        for element in newdocument.iter():
            if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        element.text = re.sub(search,replace,element.text)
        return newdocument
    
    def clean(self):
        """ Perform misc cleaning operations on documents.
            Returns cleaned document.
        """
    
        newdocument = self.document
    
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
    
    def advReplace(self,search,replace,bs=3):
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
    
        newdocument = self.document
    
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
    
                                    curlen = 0
                                    replaced = False
                                    for i in e:
                                        curlen += len(searchels[i].text)
                                        if curlen > match.start() and not replaced:
                                            # The match occurred in THIS element. Puth in the
                                            # whole replaced text
                                            if isinstance(replace, etree._Element):
                                                # Convert to a list and process it later
                                                replace = [ replace, ]
                                            if isinstance(replace, (list,tuple)):
                                                # I'm replacing with a list of etree elements
                                                # clear the text in the tag and append the element after the
                                                # parent paragraph
                                                # (because t elements cannot have childs)
                                                p = findTypeParent(searchels[i], '{%s}p' % nsprefixes['w'])
                                                searchels[i].text = re.sub(search,'',txtsearch)
                                                insindex = p.getparent().index(p) + 1
                                                for r in replace:
                                                    p.getparent().insert(insindex, r)
                                                    insindex += 1
                                            else:
                                                # Replacing with pure text
                                                searchels[i].text = re.sub(search,replace,txtsearch)
                                            replaced = True
                                            log.debug("Replacing in element #: %s", i)
                                        else:
                                            # Clears the other text elements
                                            searchels[i].text = ''
        return newdocument
        
    def _get_etree(self, xmldoc):
        return etree.fromstring(self._docx.read(xmldoc))
        
    def _load_etree(self, name, xmldoc):
        setattr(name, xmldoc)

    def save(self, dest=None):
        docxfile = self._docx
    
        # Serialize our trees into our zip file
        for tree, dest_file in self.trees_and_files.items():
            log.info('Saving: ' + dest_file)
            docxfile.writestr(dest_file, getattr(self, tree))
    
        log.info('Saved new file to: %r', dest)
        docxfile.close()
        
        if dest is not None:
            shutil.copyfile(self._tmp_file, dest)
            
    def copy(self):
        tmp = NamedTemporaryFile()
        self.save(tmp.name)
        docx = self.__class__(tmp.name)
        docx._orig_docx = self._orig_docx
        tmp.close()
        return docx
    
    def __del__(self):
        try: 
            self.__empty_docx.close()
        except AttributeError: 
            pass
        self._docx.close()
        self._tmp_file.close()
        
    def __generate_empty_docx(self):
        self.__empty_docx = NamedTemporaryFile()
        loc = self.__empty_docx.name
        
        docxfile = ZipFile(loc, mode='w', compression=ZIP_DEFLATED)

        # Move to the template data path
        prev_dir = os.path.abspath('.') # save previous working dir
        os.chdir(template_dir)
    
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
        log.info('Saved new file to: %r', loc)
        docxfile.close()
        os.chdir(prev_dir) # restore previous working dir
        
        return loc
    
    @property
    def text(self):
        '''Return the raw text of a document, as a list of paragraphs.'''
        document = self.document
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
            # Add our completed paragraph text to the list of paragraph text
            if not len(paratext) == 0:
                paratextlist.append(paratext)
        return paratextlist

"""
def contenttypes():
    # FIXME - doesn't quite work...read from string as temp hack...
    #types = Element('Types',nsprefix='ct')
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
        types.append(Element('Override',nsprefix=None,attributes={'PartName':part,'ContentType':parts[part]}))
    # Add support for filetypes
    filetypes = {'rels':'application/vnd.openxmlformats-package.relationships+xml','xml':'application/xml','jpeg':'image/jpeg','gif':'image/gif','png':'image/png'}
    for extension in filetypes:
        types.append(Element('Default',nsprefix=None,attributes={'Extension':extension,'ContentType':filetypes[extension]}))
    return types
    
def coreproperties(title,subject,creator,keywords,lastmodifiedby=None):
    '''Create core properties (common document properties referred to in the 'Dublin Core' specification).
    See appproperties() for other stuff.'''
    coreprops = Element('coreProperties',nsprefix='cp')
    coreprops.append(Element('title',tagtext=title,nsprefix='dc'))
    coreprops.append(Element('subject',tagtext=subject,nsprefix='dc'))
    coreprops.append(Element('creator',tagtext=creator,nsprefix='dc'))
    coreprops.append(Element('keywords',tagtext=','.join(keywords),nsprefix='cp'))
    if not lastmodifiedby:
        lastmodifiedby = creator
    coreprops.append(Element('lastModifiedBy',tagtext=lastmodifiedby,nsprefix='cp'))
    coreprops.append(Element('revision',tagtext='1',nsprefix='cp'))
    coreprops.append(Element('category',tagtext='Examples',nsprefix='cp'))
    coreprops.append(Element('description',tagtext='Examples',nsprefix='dc'))
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
    appprops = Element('Properties',nsprefix='ep')
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
        appprops.append(Element(prop,tagtext=props[prop],nsprefix=None))
    return appprops


def websettings():
    '''Generate websettings'''
    web = Element('webSettings')
    web.append(Element('allowPNG'))
    web.append(Element('doNotSaveAsSingleFile'))
    return web

def relationshiplist():
    relationshiplist = [
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering','numbering.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles','styles.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings','webSettings.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable','fontTable.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme','theme/theme1.xml'],
    ]
    return relationshiplist

def wordrelationships(relationshiplist):
    '''Generate a Word relationships file'''
    # Default list of relationships
    # FIXME: using string hack instead of making element
    #relationships = Element('Relationships',nsprefix='pr')
    relationships = etree.fromstring(
    '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>'''
    )
    count = 0
    for relationship in relationshiplist:
        # Relationship IDs (rId) start at 1.
        relationships.append(Element('Relationship',attributes={'Id':'rId'+str(count+1),
        'Type':relationship[0],'Target':relationship[1]},nsprefix=None))
        count += 1
    return relationships

def savedocx(document,coreprops,appprops,contenttypes,websettings,wordrelationships,output):
    '''Save a modified document'''
    assert os.path.isdir(template_dir)
    docxfile = ZipFile(output,mode='w',compression=ZIP_DEFLATED)

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
    log.info('Saved new file to: %r', output)
    docxfile.close()
    os.chdir(prev_dir) # restore previous working dir
    return
"""

