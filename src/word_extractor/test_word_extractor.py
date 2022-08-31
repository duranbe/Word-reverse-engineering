import unittest
from Docx import Docx

TEST_FILE = "test_document.docx"

class TestWordExtractor(unittest.TestCase):
    def setUp(self):
        self.docx = Docx(TEST_FILE)
        self.docx.unzip()
        self.docx.get_document_info()
        

    def test_user(self):
        self.assertEqual(self.docx.creator,"test_user")

if __name__ == '__main__':
    unittest.main()