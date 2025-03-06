import os
import zipfile
import shutil
import pdfplumber
from docx import Document
import pytesseract
from pdf2image import convert_from_path
from PIL import Image


# ✅ Automatically find the ZIP file in "input_files"
input_folder = "input_files"
print("1")
