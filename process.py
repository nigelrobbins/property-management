import os
import zipfile
import pdfplumber
import re
from docx import Document
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

# Define directories
CONFIG_DIR = "config"
LOCAL_SEARCH_DIR = os.path.join(CONFIG_DIR, "local-search")
PATTERNS_FILE = os.path.join(LOCAL_SEARCH_DIR, "patterns.txt")
MESSAGE_IF_EXISTS_FILE = os.path.join(LOCAL_SEARCH_DIR, "message_if_exists.txt")
MESSAGE_IF_NOT_EXISTS_FILE = os.path.join(LOCAL_SEARCH_DIR, "message_if_not_exists.txt")

# Ensure required directories exist
os.makedirs(LOCAL_SEARCH_DIR, exist_ok=True)
os.makedirs(CONFIG_DIR, exist_ok=True)

# Ensure message files exist with default messages if missing
if not os.path.exists(MESSAGE_IF_EXISTS_FILE):
    with open(MESSAGE_IF_EXISTS_FILE, "w") as f:
        f.write("This document contains relevant information.")

if not os.path.exists(MESSAGE_IF_NOT_EXISTS_FILE):
    with open(MESSAGE_IF_NOT_EXISTS_FILE, "w") as f:
        f.write("No relevant information found in this document.")

# Ensure patterns file exists with default patterns
if not os.path.exists(PATTERNS_FILE):
    with open(PATTERNS_FILE, "w") as f:
        f.write("REPLIES TO STANDARD ENQUIRIES.*?(?=\n[A-Z ]+\n|\Z)\n")

def load_patterns():
    """Load regex patterns from config file."""
    if os.path.exists(PATTERNS_FILE):
        with open(PATTERNS_FILE, "r") as f:
            return [line.strip() for line in f.readlines() if line.strip()]
    return []

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF, using OCR if needed."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
            else:
                images = convert_from_path(pdf_path)
                for img in images:
                    text += pytesseract.image_to_string(img) + "\n"
    return text.strip()

def extract_text_from_docx(docx_path):
    """Extract text from a Word document."""
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_matching_sections(text, patterns):
    """Extract relevant sections based on multiple regex patterns."""
    matched_sections = []
    for pattern in patterns:
        matches = re.findall(pattern, text, re.DOTALL)  # Find all matching sections
        matched_sections.extend(matches)
    
    return matched_sections

def process_zip(zip_path, output_docx):
    """Extract and process only relevant sections from documents that contain 'REPLIES TO STANDARD ENQUIRIES'."""
    output_folder = "unzipped_files"

    os.makedirs(output_folder, exist_ok=True)

    if not os.path.exists(zip_path):
        print(f"❌ ERROR: ZIP file does not exist: {zip_path}")
        return

    print(f"📂 Unzipping: {zip_path}")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)

    doc = Document()
    doc.add_paragraph(f"ZIP File: {os.path.basename(zip_path)}", style="Heading 1")
    found_relevant_doc = False
    patterns = load_patterns()

    for file_name in sorted(os.listdir(output_folder)):
        file_path = os.path.join(output_folder, file_name)

        if file_name.endswith(".pdf"):
            extracted_text = extract_text_from_pdf(file_path)
            file_type = "PDF"
        elif file_name.endswith(".docx"):
            extracted_text = extract_text_from_docx(file_path)
            file_type = "Word Document"
        else:
            continue

        # Check if the document contains "REPLIES TO STANDARD ENQUIRIES"
        if "REPLIES TO STANDARD ENQUIRIES" in extracted_text:
            matched_sections = extract_matching_sections(extracted_text, patterns)

            if matched_sections:
                found_relevant_doc = True
                doc.add_paragraph(f"Source ({file_type}): {file_name}", style="Heading 2")
                for section in matched_sections:
                    doc.add_paragraph(section)
                    doc.add_page_break()

    # Load appropriate message from file
    message_file = MESSAGE_IF_EXISTS_FILE if found_relevant_doc else MESSAGE_IF_NOT_EXISTS_FILE
    with open(message_file, "r") as f:
        extra_message = f.read().strip()
        print(f"✅ extra_message: {extra_message}")
        paragraph = doc.add_paragraph(extra_message)
        paragraph.runs[0].italic = True

    os.makedirs(os.path.dirname(output_docx), exist_ok=True)
    doc.save(output_docx)
    print(f"✅ Word document saved: {os.path.abspath(output_docx)}")

# Automatically find ZIP file and process it
input_folder = "input_files"
zip_file_path = None

for file in os.listdir(input_folder):
    if file.endswith(".zip"):
        zip_file_path = os.path.join(input_folder, file)
        break

output_file = "output_files/processed_doc.docx"
if zip_file_path:
    print(f"📂 Found ZIP file: {zip_file_path}")
    process_zip(zip_file_path, output_file)
else:
    print("❌ No ZIP file found in 'input_files' folder.")
