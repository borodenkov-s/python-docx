<<<<<<< HEAD
from contextlib import contextmanager
from core import Docx
=======
#-*- coding:utf-8 -*-

from contextlib import contextmanager
from .core import Docx
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
import elements

meta = {
    "title": "",
    "subject": "",
    "creator": "",
    "keywords": [],
}

doc = None
<<<<<<< HEAD
def start_doc(**kwargs):
    global doc, meta
    
    if kwargs.get("meta", None) is not None: 
        meta = kwargs['meta']
        
    doc = Docx()
    
=======

def start_doc(**kwargs):
    global doc, meta

    if kwargs.get("meta", None) is not None:
        meta = kwargs['meta']

    doc = Docx()

def def_settings(settings):
    doc.set_page_settings(settings)

>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
### DSL

h_ = lambda level, txt: doc.append(elements.heading(txt, level))
h1 = lambda txt: h_(1, txt)
h2 = lambda txt: h_(2, txt)
h3 = lambda txt: h_(3, txt)
h4 = lambda txt: h_(4, txt)

p = lambda txt: doc.append(elements.paragraph(txt))

br = lambda **kwargs: doc.append(elements.pagebreak(**kwargs))

<<<<<<< HEAD
def img(src, alt=""):
    raise NotImplementedError
    relationships = None
    relationships, picpara = doc.picture(relationships, src, alt)
    doc.append(picpara)
=======
def img(src, alt="", **kwargs):
    doc.append(elements.picture(doc, doc.wordrelationships, src, alt, **kwargs))
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1

@contextmanager
def ul():
    yield lambda txt: doc.append(elements.paragraph(txt, style='ListBullet'))
<<<<<<< HEAD
        
=======

>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
@contextmanager
def ol():
    yield lambda txt: doc.append(elements.paragraph(txt, style='ListNumber'))

@contextmanager
def table():
    t = []
<<<<<<< HEAD
    
    @contextmanager    
=======

    @contextmanager
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    def tr():
        r = []
        t.append(r)
        yield lambda txt: r.append(txt)
<<<<<<< HEAD
                
    yield tr
    doc.append(elements.table(t, 
        heading=False, 
        borders={"all": {'sz': 2, 'color': 'cccccc'}},
        celstyle=[{'fill': "ffffff"}] * len(t)
    ))
        
## utility functions...
def write_docx(f):
=======

    yield tr
    doc.append(elements.table(t,
        heading=False,
        borders={"all": {'sz': 2, 'color': 'cccccc'}},
        celstyle=[{'fill': "ffffff"}] * len(t)
    ))

def getdoc():
    return doc

## utility functions...
def write_docx(f=None):
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
    doc.title = str(meta.get('title', ''))
    doc.subject = str(meta.get('subject', ''))
    doc.creator = str(meta.get('creator', ''))
    doc.keywords = list(meta.get('keywords', []))
<<<<<<< HEAD
       
    doc.save(f)
=======
    return doc.save(f)
>>>>>>> 5146df06ca63ecd197f762b25934423455e747a1
