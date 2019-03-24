#!/usr/bin/env python
import os
import sys
import re
import zipfile
from zipfile import ZipFile
import xml.dom.minidom
from lxml import etree

rootdir = sys.argv[2]
matches = []

with open('output.txt','w') as fout:
    for root, subFolders, files in os.walk(rootdir):
        for filename in files:
            if filename.endswith(('.odt', '.odp', '.ods')):
                zipname = os.path.join(root, filename)
                zf = ZipFile(zipname, 'r')
                xml_string = zf.read('content.xml')
                if re.search(sys.argv[1], xml_string):
                    print 'found', (sys.argv[1]), 'in:' , os.path.join(root, filename)
                    matches.append(os.path.join(root, filename))
                zf.close()
#print '\n '.join(matches)

