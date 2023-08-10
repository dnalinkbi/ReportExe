from docx import Document
from docx.shared import Inches, Mm, Cm, Pt, RGBColor, Length
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.text.font import Font
from docx.text.parfmt import ParagraphFormat

class ReportExporter:
    def __init__(self, file_name, ReportInfos=None):
        self.doc = Document()
        self.file_name = file_name
        
        # Report Setting values
        self.section = self.doc.sections[0]
        self.section.left_margin = Mm(0.0)
        self.section.right_margin = Mm(0.0)
        self.section.top_margin = Mm(0.0)
        self.section.bottom_margin = Mm(0.0)


    def add_variable(self, name, value):
        self.doc.add_paragraph(f"{name}: {value}")

    def save_document(self):
        self.doc.save(self.file_name)