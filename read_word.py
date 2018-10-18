#!/usr/local/bin python3
import docx

def read_doc(docname):
    doc = docx.Document(docname)
    fulltext = []
    for para in doc.paragraphs:
        fulltext.append(para.text)
    return('\n'.join(fulltext))

