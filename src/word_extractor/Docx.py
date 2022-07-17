import os
import zipfile
from lxml import etree as ET
from datetime import datetime


WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
DCTERMS = "http://purl.org/dc/terms/"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
CELL = WORD_NAMESPACE + 'tc'
SECTION = WORD_NAMESPACE + 'sectPr'
LANG = WORD_NAMESPACE + 'lang'
BODY = WORD_NAMESPACE + 'body'

NS = {
    "w": WORD_NAMESPACE,
    "dcterms": DCTERMS,
    "cp": CP,
    "dc": DC
}

class Docx:

    def __init__(self,filename):
        """Docx Object

        Args:
            filename (string): Filename of the Word Document to open

        """        

        self.filename = filename
        self.creator = None
        self.last_modified_by = None
        self.modified_datetime = None
        self.created_datetime = None
    
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

    def get_document_info(self):
        
        tree = ET.parse("tmp/docProps/core.xml")
        root = tree.getroot()

        creator = root.find('.//dc:creator',NS)
        self.creator = creator.text
        dcterms_created = root.find('.//dcterms:created',NS)
        self.created_datetime = datetime.strptime(dcterms_created.text,"%Y-%m-%dT%H:%M:%SZ")
