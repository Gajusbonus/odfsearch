#!/usr/bin/env python
# use odfsearch.py searchword ./place-to-start/
# todo: convert all .doc to .odt (sudo apt-get install antiword)
import os
import sys
import re
import zipfile
from zipfile import ZipFile
import xml.dom.minidom
from lxml import etree
from subprocess import Popen, PIPE

if len(sys.argv) <= 2:
    print('To few arguments, please use: odfsearch.py searchword ./place-to-start/')
    sys.exit(0)
rootdir = sys.argv[2]
matches = []

with open('output.txt','w') as fout:
    for root, subFolders, files in os.walk(rootdir):
        for filename in files:
            if filename.endswith(('.odt', '.odp', '.ods')):
#		print filename
                zipname = os.path.join(root, filename)
                try:
                    zf = ZipFile(zipname, 'r')
                    xml_string = zf.read('content.xml')
                except:
                    print 'error',os.path.join(root, filename), "document is not a zipfile"
                if re.search(sys.argv[1], xml_string):
                    print 'found "',(sys.argv[1]), '" in:' , os.path.join(root, filename)
                    matches.append(os.path.join(root, filename))
                zf.close()
            if filename.endswith(('.docx')):
                zipname = os.path.join(root, filename)
                try:
                    zf = ZipFile(zipname, 'r')
                    xml_string = zf.read('word/document.xml')
                    if re.search(sys.argv[1], xml_string):
                        print 'found', (sys.argv[1]), 'in:' , os.path.join(root, filename)
                        matches.append(os.path.join(root, filename))
                    zf.close()
                except:
                    print 'error',os.path.join(root, filename), " is not a zipfile"
            if filename.endswith(('.doc')):
                word_doc=os.path.join(root, filename)
                cmd = ['antiword',word_doc]
                try:
                    p = Popen(cmd, stdout=PIPE)
                    doc_text = p.stdout.read()
                    if re.search(sys.argv[1], doc_text):
                        print "found" , (sys.argv[1]), 'in:' , os.path.join(root, filename)
                except:
                    print "error"

                    #print '\n '.join(matches)
