# encoding: utf-8

from elements import Element

def core_properties(title,subject,creator,keywords,lastmodifiedby=None):
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

def app_properties():
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


def web_settings():
    '''Generate websettings'''
    web = Element('webSettings')
    web.append(Element('allowPNG'))
    web.append(Element('doNotSaveAsSingleFile'))
    return web

def relationship_list():
    relationshiplist = [
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering','numbering.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles','styles.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings','webSettings.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable','fontTable.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme','theme/theme1.xml'],
    ]
    return relationshiplist
    
def word_relationships(relationshiplist):
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