#
# Finds old Excel-files (.xls) and converts them into XML-format (.xlsx)
#
import os
import subprocess
import re
import Settings
from pprint import pprint

def enumNonXMLFiles(dir):
    nonXML = []

    for path, subdirs, files in os.walk(dir):
        for name in files:
            if re.match('.*\.xls$', name):
                nonXML.append(os.path.abspath(os.path.join(path, name)))

    return nonXML

def convertXLSToXLSX(file):
    print "Converting file: " + file

    cmd = "\"" + Settings.PathSettings['ExcelCnv'] + "\"" + ' ' + '-oice ' +  file + ' ' + file+'x'
    ret = subprocess.call( cmd, shell=True )
    
    if ret != 0:
        print "File conversion from xls to xlsx failed: " + file + "!"


# TEST
if __name__ == "__main__":
    files = enumNonXMLFiles(os.getcwd())

    for file in files:
        convertXLSToXLSX(file)

    
