from docx import Document
from docx.shared import Inches
from docx.shared import Mm
from docx.shared import Cm
from docx.shared import Pt
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsdecls
from docx.text.font import Font
from docx.text.parfmt import ParagraphFormat
from docx.shared import Length
from docx.oxml import parse_xml
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.table import WD_TABLE_ALIGNMENT
import pandas as pd
from pandas import DataFrame
import subprocess
from pandas import read_csv
import sys
import numpy as np

# Not updated packages
# from HL.Add_HL import add_hyperlink as HL
#from matplotlib import pyplot as plt


