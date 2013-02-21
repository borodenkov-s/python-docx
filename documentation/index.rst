=======================================
Welcome to python-docx's documentation!
=======================================

.. toctree::
   :glob:
   
   *

The docx module creates, reads and writes Microsoft Office Word 2007 docx files.

These are referred to as 'WordML', 'Office Open XML' and 'Open XML' by Microsoft.

These documents can be opened in Microsoft Office 2007 / 2010, Microsoft Mac Office 2008, Google Docs, OpenOffice.org 3, and Apple iWork 08.

They also `validate as well formed XML <http://validator.w3.org/check>`_

The module was created when I was looking for a Python support for MS Word .doc files, but could only find various hacks involving COM automation, calling .net or Java, or automating OpenOffice or MS Office.

Features
========

The docx module has the following features:

Making documents
----------------

Features for making documents include:

- Paragraphs
- Bullets
- Numbered lists
- Document properties (author, company, etc)
- Multiple levels of headings
- Tables
- Section and page breaks
- Images

<div style="float: right"><img src="http://github.com/mikemaccana/python-docx/raw/master/screenshot.png"></div>

Editing documents
-----------------

Thanks to the awesomeness of the lxml module, we can:

- Search and replace
- Extract plain text of document
- Add and delete items anywhere within the document
- Change document properties
- Run xpath queries against particular locations in the document - useful for retrieving data from user-completed templates.

Getting started
===============

Making and Modifying Documents
------------------------------

- Just `download python docx<http://github.com/mikemaccana/python-docx/tarball/master>`_
- Use **pip** or **easy_install** to fetch the **lxml** and **PIL** modules. 
- Then run: ``example-makedocument.py``

Congratulations, you just made and then modified a Word document!

Extracting Text from a Document
-------------------------------

If you just want to extract the text from a Word file, run: 

    example-extracttext.py 'Some word file.docx' 'new file.txt' 

Tips
====

If Word complains about files
-----------------------------

First, determine whether Word can recover the files:

- If Word cannot recover the file, you most likely have a problem with your zip file
- If Word can recover the file, you most likely have a problem with your XML

Common Zipfile issues
---------------------

- Ensure the same file isn't included twice in your zip archive. Zip supports this, Word doesn't.
- Ensure that all media files have an entry for their file type in [Content_Types].xml
- Ensure that files in zip file file have leading '/'s removed. 

Common XML issues
-----------------

- Ensure the _rels, docProps, word, etc directories are in the top level of your zip file.
- Check your namespaces - on both the tags, and the attributes
- Check capitalization of tag names
- Ensure you're not missing any attributes
- If images or other embedded content is shown with a large red X, your relationships file is missing data.

One common debugging technique we've used before
------------------------------------------------

- Re-save the document in Word will produced a fixed version of the file
- Unzip and grabbing the serialized XML out of the fixed file
- Use etree.fromstring() to turn it into an element, and include that in your code.
- Check that a correct file is generated
- Remove an element from your string-created etree (including both opening and closing tags)
- Use element.append(makelement()) to add that element to your tree
- Open the doc in Word and see if it still works
- Repeat the last three steps until you discover which element is causing the prob

Want to talk? Need help?
========================

Email `mailto:python.docx@librelist.com`_