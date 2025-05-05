import fitz  # PyMuPDF
from docx import Document
import os

class DocumentParser:
    def parse(self, file_path):
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.pdf':
            return self._parse_pdf(file_path)
        elif file_ext == '.docx':
            return self._parse_docx(file_path)
        elif file_ext == '.txt':
            return self._parse_txt(file_path)
        else:
            raise ValueError("Unsupported file format")
    
    def _parse_pdf(self, file_path):
        text = ""
        with fitz.open(file_path) as doc:
            for page in doc:
                text += page.get_text()
        return text
    
    def _parse_docx(self, file_path):
        doc = Document(file_path)
        return "\n".join([paragraph.text for paragraph in doc.paragraphs])
    
    def _parse_txt(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()