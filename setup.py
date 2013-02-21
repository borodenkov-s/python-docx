#!/usr/bin/env python
from setuptools import setup

<<<<<<< HEAD
<<<<<<< HEAD
from setuptools import setup
from glob import glob
=======
from distutils.core import setup
#from glob import glob
#from distutils.command.build_py import build_py
>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c

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
<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
      version='0.0.2',
=======
      version='0.2',
>>>>>>> b0abe8d0c2ba44c3dd6675c5eb3198925a84431e
=======
      version='0.0.3',
>>>>>>> e38d8395280f6fcf03413fc754b73650aaf9c4ea
      requires=['lxml'],
=======
      version='0.0.4',
      requires=['lxml', 'PIL'],
>>>>>>> 521eccd1766262672d82224524e3cfb279bb9e4f
=======
      version='0.0.3',
      requires=['lxml', 'Image', 'PIL'],
>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
      description='The docx module creates, reads and writes Microsoft Office Word 2007 docx files',
      long_description=long_description,
      author='Mike MacCana',
      author_email='python.docx@librelist.com',
      license="MIT License",
      url='http://github.com/mikemaccana/python-docx',
<<<<<<< HEAD
      py_modules=['docx'],
      data_files=[
<<<<<<< HEAD
          ('docx-template/_rels', glob('docx-template/_rels/.*')),
          ('docx-template/docProps', glob('docx-template/docProps/*.*')),
          ('docx-template/word', glob('docx-template/word/*.xml')),
          ('docx-template/word/theme', glob('docx-template/word/theme/*.*')),
=======
          ('docx-template/_rels', glob('examples/template/_rels/.*')),
          ('docx-template/docProps', glob('examples/template/docProps/*.*')),
          ('docx-template/word', glob('examples/template/word/*.xml')),
          ('docx-template/word/theme', glob('examples/template/word/theme/*.*')),
>>>>>>> b0abe8d0c2ba44c3dd6675c5eb3198925a84431e
          ],
      zip_safe=False,
=======
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
>>>>>>> 7848a674630e83b6ef9c840e0e0ea94fb7a4120c
      )
=======
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
    packages=['docx'],
    package_data={
        'docx': ['docx/template/*']
    },
)
>>>>>>> 7093a83298ef89a905b9be84dac18bfbdeee394f
