from lxml import etree
from docx import NSPREFIXES


def make_element(tagname, tagtext=None, nsprefix='w', attributes=None, attrnsprefix=None):
    '''Create an element & return it''' 
    # Deal with list of nsprefix by making namespacemap
    namespacemap = None
    if type(nsprefix) == list:
        namespacemap = {}
        for prefix in nsprefix:
            namespacemap[prefix] = NSPREFIXES[prefix]
        nsprefix = nsprefix[0] # FIXME: rest of code below expects a single prefix
    if nsprefix:
        namespace = '{'+NSPREFIXES[nsprefix]+'}'
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
            attributenamespace = '{'+NSPREFIXES[attrnsprefix]+'}'
                    
        for tagattribute in attributes:
            newelement.set(attributenamespace+tagattribute, attributes[tagattribute])
    if tagtext:
        newelement.text = tagtext    
    return newelement

# -*- coding: utf-8 -*-
import os
from os.path import join, abspath
from zipfile import ZipFile, ZIP_DEFLATED

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
    
def dir_to_docx(source_dir, output_loc):
    # make sure the output ends up where we expect it
    output_loc = abspath(output_loc)
    
    # Move into the source_dir
    prev_dir = abspath('.') # save previous working dir
    os.chdir(source_dir)
    
    docxfile = ZipFile(output_loc, mode='w', compression=ZIP_DEFLATED)    

    # Add & compress support files
    files_to_ignore = set(['.DS_Store']) # nuisance from some os's
    for dirpath, dirnames, filenames in os.walk('.'): #@UnusedVariable
        for filename in filenames:
            if filename in files_to_ignore:
                continue
            templatefile = join(dirpath,filename)
            archivename = templatefile[2:]
            docxfile.write(templatefile, archivename)
    docxfile.close()
    
    os.chdir(prev_dir) # restore previous working dir

def cm2dxa(cm):
    """ convetion function to make things easy """
    # "word processes files at 72dpi"
    # cf http://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
    return int(cm / 2.54 * 72 * 20)


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


def dir_to_docx(source_dir, output_loc):
    # make sure the output ends up where we expect it
    output_loc = abspath(output_loc)

    # Move into the source_dir
    prev_dir = abspath('.')  # save previous working dir
    os.chdir(source_dir)

    docxfile = ZipFile(output_loc, mode='w', compression=ZIP_DEFLATED)

    # Add & compress support files
    files_to_ignore = set(['.DS_Store'])  # nuisance from some os's
    for dirpath, dirnames, filenames in os.walk('.'):  #@UnusedVariable
        for filename in filenames:
            if filename in files_to_ignore:
                continue
            templatefile = join(dirpath, filename)
            archivename = templatefile[2:]
            docxfile.write(templatefile, archivename)
    docxfile.close()

    os.chdir(prev_dir)  # restore previous working dir

def new_id(relationshiplist):
    try:
        nextid = max(list(
            int(rel.get('Id').replace('rId',''))
            for idx, rel
            in enumerate(relationshiplist))) + 1
    except:
        nextid = 1
    return 'rId' + str(nextid)

