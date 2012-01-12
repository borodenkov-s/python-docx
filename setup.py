#!/usr/bin/env python
from setuptools import setup

setup(
    name='docx',
    version='0.1.1',
    requires=(
        'lxml', 
        'python_dateutil',
    ),
    description='The docx module creates, reads and writes Microsoft Office Word 2007 docx files',
    author='Mike MacCana',
    author_email='python.docx@librelist.com',
    url='http://github.com/mikemaccana/python-docx',
    py_modules=['docx'],
    package_data={
        'docx': ['docx/template/*']
    },
)
