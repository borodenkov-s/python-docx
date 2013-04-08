#!/usr/bin/env python

from setuptools import setup
from glob import glob

# Make data go into site-packages (http://tinyurl.com/site-pkg)
from distutils.command.install import INSTALL_SCHEMES
for scheme in INSTALL_SCHEMES.values():
    scheme['data'] = scheme['purelib']

setup(name='docx',
      version='0.0.2',
      requires=['lxml'],
      description='The docx module creates, reads and writes Microsoft Office Word 2007 docx files',
      author='Mike MacCana',
      author_email='python.docx@librelist.com',
      url='http://github.com/mikemaccana/python-docx',
      py_modules=['docx'],
      data_files=[
          ('docx-template/_rels', glob('docx-template/_rels/.*')),
          ('docx-template/docProps', glob('docx-template/docProps/*.*')),
          ('docx-template/word', glob('docx-template/word/*.xml')),
          ('docx-template/word/theme', glob('docx-template/word/theme/*.*')),
          ],
      zip_safe=False,
      )
