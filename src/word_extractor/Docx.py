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

    def zip(self,output_filename="output.docx",directory="tmp"):
        """
        Rezip a folder of xml files into a Word file

        Args:
            output_filename (string): Filename of the output

            directory (string): Directory where the docx file has been extracted
        
        """
        

        with zipfile.ZipFile(output_filename, 'w') as zip_object:
            for foldername, subfolders, filenames in os.walk(directory):
                for filename in filenames:
                    filepath = os.path.join(foldername, filename)
                    filepath_zip = filepath.replace(directory+"\\",'')
                    zip_object.write(filepath,filepath_zip)

    def get_document_info(self):
        """

        Get document information such as creator, creation date and time, title etc
        
        """
        
        tree = ET.parse("tmp/docProps/core.xml")
        root = tree.getroot()

        self.creator = root.find('.//dc:creator',NS).text
        dcterms_created = root.find('.//dcterms:created',NS)
        self.created_datetime = datetime.strptime(dcterms_created.text,"%Y-%m-%dT%H:%M:%SZ")
        self.title = root.find('.//dc:title',NS).text
        self.last_modified_by = root.find('.//cp:lastModifiedBy',NS).text
        

