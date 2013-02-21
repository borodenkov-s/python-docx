#!/usr/bin/env python
from setuptools import setup, find_packages
from glob import glob

# Make data go into site-packages (http://tinyurl.com/site-pkg)
from distutils.command.install import INSTALL_SCHEMES
for scheme in INSTALL_SCHEMES.values():
    scheme['data'] = scheme['purelib']
    
long_description = """
The docx module creates, reads and writes Microsoft Office Word 2007 docx files.  Contains the following features:

Making documents:

- Paragraphs
- Bullets
- Numbered lists
- Document properties (author, company, etc)
- Multiple levels of headings
- Tables
- Section and page breaks
- Images


Editing documents:

- Search and replace
- Extract plain text of document
- Add and delete items anywhere within the document
- Change document properties
- Run xpath queries against particular locations in the document - usefull to get data from user-completed templates.

"""
setup(
    name='docx',
    version='0.1.3',
    requires=(
        'lxml',
        'python_dateutil',
    ),
    description='The docx module creates, reads and writes Microsoft Office Word 2007 docx files',
    author='Mike MacCana',
    author_email='python.docx@librelist.com',
    url='http://github.com/jiaaro/python-docx',
    packages=find_packages(),
    package_dir={'docx': 'docx'},
    py_modules=['docx'],
    data_files=[
          ('docx/template/', glob('docx/template/*.*')),
          ('docx/template/_rels', glob('docx/template/_rels/.*')),
          ('docx/template/docProps', glob('docx/template/docProps/*.*')),
          ('docx/template/word', glob('docx/template/word/*.xml')),
          ('docx/template/word/_rels', glob('docx/template/word/_rels/*.*')),
          ('docx/template/word/theme', glob('docx/template/word/theme/*.*')),
          ],
)
