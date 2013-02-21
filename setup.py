#!/usr/bin/env python
from setuptools import setup

<<<<<<< HEAD
from setuptools import setup
from glob import glob

# Make data go into site-packages (http://tinyurl.com/site-pkg)
from distutils.command.install import INSTALL_SCHEMES
for scheme in INSTALL_SCHEMES.values():
    scheme['data'] = scheme['purelib']

setup(name='docx',
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
