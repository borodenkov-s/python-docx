#!/usr/bin/env python

from distutils.core import setup
from glob import glob

# Make data go into site-packages (http://tinyurl.com/site-pkg)
from distutils.command.install import INSTALL_SCHEMES
for scheme in INSTALL_SCHEMES.values():
    scheme['data'] = scheme['purelib']

setup(name='docx',
      version='0.1.',
      requires=['lxml','python_dateutil'], # dateutil for just ONE line of code !!!
      description='The docx module creates, reads and writes Microsoft Office Word 2007 docx files',
      author='Mike MacCana',
      author_email='python.docx@librelist.com',
      url='http://github.com/mikemaccana/python-docx',
      py_modules=['docx'],
      package_data={
        'docx': ['docx/template/*']}
      )
