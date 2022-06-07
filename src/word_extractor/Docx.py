import os
import zipfile
from lxml import etree as ET


WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
CELL = WORD_NAMESPACE + 'tc'
SECTION = WORD_NAMESPACE + 'sectPr'


class Docx:

    def __init__(self,filename):

        self.filename = filename
    
    def unzip(self):
    
        with zipfile.ZipFile(self.filename, 'r') as zip_ref:
            zip_ref.extractall("tmp")

    def zip(self):
        with zipfile.ZipFile(self.filename, 'w') as zipObj:
            for foldername, subfolders, filenames in os.walk("tmp"):
                for filename in filenames:
                    filepath = os.path.join(foldername, filename)
                    filepath_zip = filePath.replace('tmp\\','')
                    zipObj.write(filepath,filepath_zip)
