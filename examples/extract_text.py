"""
Extracting the text of a simple docx file. 
"""
import os
import sys

# adding the parent directory to PATH
path = os.path.abspath(os.path.join(os.path.dirname(__file__),".."))
sys.path.append(path)

from docx.document import DocxDocument


if __name__ == '__main__':
    try:
        doc = DocxDocument(sys.argv[1])
        newfile = open(sys.argv[2],'w')
    except:
        print('Please supply an input and output file. For example:')
        print('''  extract_text.py 'My Office 2007 extract.docx' 'outputfile.txt' ''')
        exit()
    ## Fetch all the text out of the document we just created
    paragraphs = doc.get_text()
    # Make explicit unicode version
    paragraphs_encoded = []
    for p in paragraphs:
        paragraphs_encoded.append(p.encode("utf-8"))
    ## Print our documnts test with two newlines under each paragraph
    newfile.write('\n\n'.join(paragraphs_encoded))
