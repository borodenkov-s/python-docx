#-*- coding:utf-8 -*-

from os.path import abspath, dirname, join

PACKAGE_DIR = abspath(dirname(__file__))

TEMPLATE_DIR = join(PACKAGE_DIR, 'template')

image_relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
hlink_relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'


# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily.
nsprefixes = {
    # Text Content
    'mv': 'urn:schemas-microsoft-com:mac:vml',
    'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o': 'urn:schemas-microsoft-com:office:office',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    # Properties (core and extended)
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/',
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    # Content Types (we're just making up our own namespaces here to save time)
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships (we're just making up our own namespaces here to save time)
    'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
    }

FORMAT = {
        "letter": {"w": '12240', "h": '15840'},
        "a4": {"w": '11906', "h": '16838'},
        "a5": {"w": '8391', "h": '11906'},
        }

PAGESETTINGS = {
    'pgMar': {'bottom': '720', 'footer': '0', 'gutter': '0', 'header': '0',
                        'left': '1138', 'right': '1138', 'top': '1138'},
    'type': {'val': 'nextPage'},
    'pgSz': FORMAT['a4'],
    'pgNumType': {'fmt': 'decimal'},
    'formProt': {'val': 'false'},
    'textDirection': {'val': 'lrTb'},
    'docGrid': {'charSpace': '0', 'linePitch': '240', 'type': 'default'}
    }
