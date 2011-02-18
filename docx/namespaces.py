# encoding: utf-8

namespaces = {
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

content_types = {
    '/word/theme/theme1.xml':'application/vnd.openxmlformats-officedocument.theme+xml',
    '/word/fontTable.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
    '/docProps/core.xml':'application/vnd.openxmlformats-package.core-properties+xml',
    '/docProps/app.xml':'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    '/word/document.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    '/word/settings.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
    '/word/numbering.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
    '/word/styles.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
    '/word/webSettings.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml',
    }

file_types = {'rels':'application/vnd.openxmlformats-package.relationships+xml','xml':'application/xml','jpeg':'image/jpeg','gif':'image/gif','png':'image/png'}