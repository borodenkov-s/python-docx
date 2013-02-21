#!/usr/bin/env python

from setuptools import setup
from glob import glob

# Make data go into site-packages (http://tinyurl.com/site-pkg)
from distutils.command.install import INSTALL_SCHEMES
for scheme in INSTALL_SCHEMES.values():
    scheme['data'] = scheme['purelib']

setup(name='docx',
<<<<<<< HEAD
      version='0.0.2',
=======
      version='0.2',
>>>>>>> b0abe8d0c2ba44c3dd6675c5eb3198925a84431e
      requires=['lxml'],
      description='The docx module creates, reads and writes Microsoft Office Word 2007 docx files',
      author='Mike MacCana',
      author_email='python.docx@librelist.com',
      url='http://github.com/mikemaccana/python-docx',
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
      )
