import os
import zipfile
import pdfplumber
from docx import Document
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

# Define directories
CONFIG_DIR = "config"
LOCAL_SEARCH_DIR = os.path.join(CONFIG_DIR, "local-search")
MESSAGE_IF_EXISTS_FILE = os.path.join(CONFIG_DIR, "message_if_exists.txt")
MESSAGE_IF_NOT_EXISTS_FILE = os.path.join(CONFIG_DIR, "message_if_not_exists.txt")

# Ensure required directories exist
os.makedirs(LOCAL_SEARCH_DIR, exist_ok=True)
os.makedirs(CONFIG_DIR, exist_ok=True)

# Ensure message files exist with default messages if missing
if not os.path.exists(MESSAGE_IF_EXISTS_FILE):
    with open(MESSAGE_IF_EXISTS_FILE, "w") as f:
        f.write("This document contains REPLIES TO STANDARD ENQUIRIES.")

if not os.path.exists(MESSAGE_IF_NOT_EXISTS_FILE):
    with open(MESSAGE_IF_NOT_EXISTS_FILE, "w") as f:
        f.write("No relevant replies found in this document.")

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

def find_zip_file(directory):
    """Find the first ZIP file in the directory."""
    print(f"📂 Checking directory: {directory}")

    if not os.path.exists(directory):
        print(f"❌ ERROR: Directory does not exist: {directory}")
        return None

    try:
        files = os.listdir(directory)
        for file in files:
            if file.endswith(".zip"):
                return os.path.join(directory, file)
    except Exception as e:
        print(f"❌ ERROR while listing files: {e}")
        return None

    return None  # No ZIP file found

def process_zip(zip_path, output_docx):
    """Extract text from PDFs and Word docs, filter by keyword, and save results."""
    output_folder = "unzipped_files"
    processed_folder = "processed_files"
    keyword = "REPLIES TO STANDARD ENQUIRIES"

    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(processed_folder, exist_ok=True)

    if not os.path.exists(zip_path):
        print(f"❌ ERROR: ZIP file does not exist: {zip_path}")
        return

    print(f"📂 Unzipping: {zip_path}")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)

    doc = Document()
    doc.add_paragraph(f"ZIP File: {os.path.basename(zip_path)}", style="Heading 1")
    found_relevant_doc = False

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

        if keyword in extracted_text:
            found_relevant_doc = True
            doc.add_paragraph(f"Source ({file_type}): {file_name}", style="Heading 2")
            doc.add_paragraph(extracted_text)
            doc.add_page_break()

    # Load appropriate message from file
    message_file = MESSAGE_IF_EXISTS_FILE if found_relevant_doc else MESSAGE_IF_NOT_EXISTS_FILE
    with open(message_file, "r") as f:
        extra_message = f.read().strip()
        doc.add_paragraph(extra_message, style="Italic")

    os.makedirs(os.path.dirname(output_docx), exist_ok=True)
    doc.save(output_docx)
    print(f"✅ Word document saved: {os.path.abspath(output_docx)}")

# Automatically find ZIP file and process it
input_folder = "input_files"
zip_file_path = find_zip_file(input_folder)

output_file = "output_files/processed_doc.docx"
if zip_file_path:
    print(f"📂 Found ZIP file: {zip_file_path}")
    process_zip(zip_file_path, output_file)
else:
    print("❌ No ZIP file found in 'input_files' folder.")
