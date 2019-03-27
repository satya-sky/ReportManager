from openpyxl.styles import Alignment, Font, Color
from openpyxl.styles.borders import Border, Side
# from openpyxl.utils import coordinate_from_string, column_index_from_string
# from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
# import PIL import Image, ImageOps
import io
import urllib3
import image
import openpyxl
import pandas as pd
import pdb
import sys
import time as t
# import win32com.client
import xlsxwriter
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

TIMESTAMP = t.strftime("%Y%m%d_%H%M%S")
