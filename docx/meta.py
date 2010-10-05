import time
from lxml import etree
from docx.utils import make_element

class CoreProperties(object):
    """Core properties for a document.
    """
    def __init__(self, title, creator, keywords, lastmodifiedby=None):
        self.title=title
        self.creator=creator
        self.keywords=keywords
        self.lastmodifiedby=lastmodifiedby if lastmodifiedby is not None else creator

    def _xml(self):
        coreprops = make_element('coreProperties',nsprefix='cp')    
        coreprops.append(make_element('title', tagtext=self.title, nsprefix='dc'))
        coreprops.append(make_element('subject', tagtext=self.subject, nsprefix='dc'))
        coreprops.append(make_element('creator', tagtext=self.creator, nsprefix='dc'))
        coreprops.append(make_element('keywords', tagtext=','.join(self.keywords), nsprefix='cp'))    
        coreprops.append(make_element('lastModifiedBy', tagtext=self.lastmodifiedby, nsprefix='cp'))
        coreprops.append(make_element('revision', tagtext='1', nsprefix='cp'))
        coreprops.append(make_element('category', tagtext='Examples', nsprefix='cp'))
        coreprops.append(make_element('description', tagtext='Examples', nsprefix='dc'))
        currenttime = time.strftime('%Y-%m-%dT%H:%M:%SZ')
        # Document creation and modify times
        # Prob here: we have an attribute who name uses one namespace, and that 
        # attribute's value uses another namespace.
        # We're creating the lement from a string as a workaround...
        for doctime in ['created','modified']:
            coreprops.append(etree.fromstring('''<dcterms:'''+doctime+''' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">'''+currenttime+'''</dcterms:'''+doctime+'''>'''))
            pass
        return coreprops


class AppProperties(object):
    """Properties describing the application which created the OpenXML file."""

    def __init__(self, application='Microsoft Word 12.0.0', version='12.000'):
        self.application=application
        self.version=version

    def _xml(self):
        appprops = make_element('Properties',nsprefix='ep')
        appprops = etree.fromstring(
        '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>''')
        props = {
                'Template':'Normal.dotm',
                'TotalTime':'6',
                'Pages':'1',  
                'Words':'83',   
                'Characters':'475', 
                'Application':self.application,
                'DocSecurity':'0',
                'Lines':'12', 
                'Paragraphs':'8',
                'ScaleCrop':'false', 
                'LinksUpToDate':'false', 
                'CharactersWithSpaces':'583',  
                'SharedDoc':'false',
                'HyperlinksChanged':'false',
                'AppVersion':self.version,
                }
        for prop in props:
            appprops.append(make_element(prop,tagtext=props[prop],nsprefix=None))
        return appprops

