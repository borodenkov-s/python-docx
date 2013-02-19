#!/usr/bin/env python

from distutils.core import setup
#from glob import glob
#from distutils.command.build_py import build_py

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

setup(name='docx',
      version='0.0.3',
      requires=['lxml', 'Image', 'PIL'],
      description='The docx module creates, reads and writes Microsoft Office Word 2007 docx files',
      long_description=long_description,
      author='Mike MacCana',
      author_email='python.docx@librelist.com',
      license="MIT License",
      url='http://github.com/mikemaccana/python-docx',
      download_url='http://github.com/mikemaccana/python-docx/tarball/master',
      classifiers=[
          "Development Status :: 3 - Alpha",
          "Intended Audience :: Developers",
          "License :: OSI Approved :: MIT License",
          "Programming Language :: Python :: 2",
          "Programming Language :: Python :: 3",
          "Operating System :: OS Independent",
          "Topic :: Software Development :: Libraries :: Python Modules"
      ],
      packages=["docx"],
      package_data={'docx': ['docx-template/_rels/.*', 'docx-template/docProps/*.*', 'docx-template/word/*.xml',
                             'docx-template/word/theme/*.*']},
      )
