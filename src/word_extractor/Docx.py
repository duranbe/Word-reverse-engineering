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
        self.title = None
    
    def unzip(self):
        """"
        Unzip a docx file and extract the xml files into a tmpfolder

        """
    
        with zipfile.ZipFile(self.filename, 'r') as zip_ref:
            zip_ref.extractall("tmp")

    def zip(self,filename,directory="tmp"):
        """
        Rezip a folder of xml files into a Word file
        
        """

        with zipfile.ZipFile(self.filename, 'w') as zip_object:
            for foldername, subfolders, filenames in os.walk(directory):
                for filename in filenames:
                    filepath = os.path.join(foldername, filename)
                    filepath_zip = filepath.replace(directory+"\\",'')
                    zip_object.write(filepath,filepath_zip)

    def get_document_info(self):
        
        tree = ET.parse("tmp/docProps/core.xml")
        root = tree.getroot()

        creator = root.find('.//dc:creator',NS)
        self.creator = creator.text
        dcterms_created = root.find('.//dcterms:created',NS)
        self.created_datetime = datetime.strptime(dcterms_created.text,"%Y-%m-%dT%H:%M:%SZ")
        title = root.find('.//dc:title',NS)
        self.title = title.text
        last_modified_by = root.find('.//cp:lastModifiedBy',NS)
        self.last_modified_by = last_modified_by 

