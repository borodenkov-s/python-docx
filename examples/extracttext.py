#!/usr/bin/env python2.6
"""
This file opens a docx (Office 2007) file and dumps the text.

If you need to extract text from documents, use this file as a basis for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""
from docx.document import *
import sys
if __name__ == '__main__':        
    try:
        document = Docx(sys.argv[1])
        newfile = open(sys.argv[2],'w')        
    except:
        print('Please supply an input and output file. For example:')    
        print("""  example-extracttext.py 'My Office 2007 document.docx' 'outputfile.txt' """)    
        exit()
    ## Fetch all the text out of the document we just created        
    paratextlist = document.text    

    # Make explicit unicode version    
    newparatextlist = []
    for paratext in paratextlist:
        newparatextlist.append(paratext.encode("utf-8"))                  
    
    ## Print our documnts test with two newlines under each paragraph
    newfile.write('\n\n'.join(newparatextlist))
    #print '\n\n'.join(newparatextlist)