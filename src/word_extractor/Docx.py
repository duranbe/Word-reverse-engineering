import os
import zipfile
from lxml import etree as ET
from datetime import datetime
from Namespace import NS


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
