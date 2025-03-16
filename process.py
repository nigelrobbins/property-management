import os
import zipfile
import pdfplumber
import re
import time
from docx import Document
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import subprocess
import yaml

def timed_function(func):
    """Decorator to measure function execution time."""
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        print(f"⏱ {func.__name__} took {end_time - start_time:.4f} seconds")
        return result
    return wrapper

@timed_function
def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF, using pdftotext first, then pdfplumber, then OCR if needed."""

    # Try using pdftotext first
    result = subprocess.run(['pdftotext', pdf_path, '-'], capture_output=True, text=True)
    text = result.stdout.strip()

    if text:
        print(f"✅ Extracted text using pdftotext: {text}...")
        return text  # If pdftotext works, return immediately

    print("⚠️ pdftotext failed, trying pdfplumber...")

    # Fallback to pdfplumber
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"

    if text:
        print(f"✅ Extracted text using pdfplumber: {text}...")
        return text.strip()  # If pdfplumber works, return immediately

    print("⚠️ pdfplumber failed, performing OCR...")

    # Final fallback: Use OCR (slow)
    images = convert_from_path(pdf_path)
    for img in images:
        text += pytesseract.image_to_string(img) + "\n"

    print(f"✅ Extracted text using OCR: {text}...")
    return text.strip()

def extract_text_from_docx(docx_path):
    """Extract text from a Word document."""
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

# Load YAML configuration
def load_yaml(yaml_path):
    with open(yaml_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)["groups"]

# Identify question group based on document content
def identify_group(text, groups):
    for group in groups:
        if group["identifier"] in text:
            return group
    return None  # No matching group found

def extract_matching_text(text, pattern, message_template):
    """Extracts matching text based on the given pattern and formats the message."""
    # Find the matching text based on the pattern
    matches = re.findall(pattern, text, re.IGNORECASE | re.MULTILINE)
    
    if matches:
        # Log the matches for debugging
        print(f"✅ Matches found: {matches}")
        
        extracted_text_1 = matches[0]  # First part of the extracted text
        extracted_text_2 = matches[1] if len(matches) > 1 else ''  # Second part of the extracted text (optional)

        # Log the extracted content for debugging
        print(f"✅ Extracted text: {extracted_text_1}, {extracted_text_2}")
        
        # Format the message with the extracted text
        formatted_message = message_template.format(extracted_text_1=extracted_text_1, extracted_text_2=extracted_text_2)
        
        # Log the formatted message for debugging
        print(f"✅ Formatted message: {formatted_message}")
        
        return formatted_message
    else:
        print("⚠️ No matches found for the pattern.")
        return None

def process_zip(zip_path, output_docx, yaml_path):
    """Extract and process only relevant sections from documents that contain filter text."""
    output_folder = "unzipped_files"
    os.makedirs(output_folder, exist_ok=True)
    groups = load_yaml(yaml_path)
    doc = Document()

    if not os.path.exists(zip_path):
        print(f"❌ ERROR: ZIP file does not exist: {zip_path}")
        return

    print(f"📂 Unzipping: {zip_path}")
    unzip_start = time.time()
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)
    unzip_end = time.time()
    print(f"⏱ Unzipping took {unzip_end - unzip_start:.4f} seconds")

    doc.add_paragraph(f"ZIP File: {os.path.basename(zip_path)}", style="Heading 1")

    for file_name in sorted(os.listdir(output_folder)):
        file_path = os.path.join(output_folder, file_name)

        print(f"📄 Processing {file_name}...")
        process_start = time.time()

        if file_name.endswith(".pdf"):
            extracted_text = extract_text_from_pdf(file_path)
            file_type = "PDF"
        elif file_name.endswith(".docx"):
            extracted_text = extract_text_from_docx(file_path)
            file_type = "Word Document"
        else:
            continue
        group = identify_group(extracted_text, groups)

        if group:
            doc.add_paragraph(group["message_if_identifier_found"], style="Heading 2")
            print(group["message_if_identifier_found"])
        else:
            print("⚠️ No matching group found. Skipping.")
            continue  # Skip this file if no match
        
        doc.add_paragraph(f"📂 Processing: {file_name}", style="Heading 1")
        doc.add_paragraph(f"📄 Document identified as: {group['name']}", style="Heading 2")

        for question in group["questions"]:
            doc.add_paragraph(f"🔍 Checking section: {question['section']}", style="Heading 3")

            if question["search_pattern"] in extracted_text:
                doc.add_paragraph(question["message_found"], style="Normal")
                if question["extract_text"]:
                    extracted_section = extract_matching_text(extracted_text, question["extract_pattern"], question["message_found"])
                    if extracted_section:
                        print(f"✅ Extracted content: {extracted_section[:50]}...")  # Log first 50 characters
                        paragraph = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
                        paragraph.insert_paragraph_before(extracted_section)
                        paragraph.runs[0].italic = True
                    else:
                        doc.add_paragraph("⚠️ No matching content found.", style="Normal")
            else:
                doc.add_paragraph(question["message_not_found"], style="Normal")

            doc.add_paragraph("")  # Spacing between questions

        doc.add_page_break()

    # Save output Word document
    os.makedirs(os.path.dirname(output_docx), exist_ok=True)
    doc.save(output_docx)

# Automatically find ZIP file and process it
input_folder = "input_files"
zip_file_path = None
yaml_config = "config.yaml"

for file in os.listdir(input_folder):
    if file.endswith(".zip"):
        zip_file_path = os.path.join(input_folder, file)
        break

output_file = "output_files/processed_doc.docx"
if zip_file_path:
    print(f"📂 Found ZIP file: {zip_file_path}")
    process_zip(zip_file_path, output_file, yaml_config)
else:
    print("❌ No ZIP file found in 'input_files' folder.")

# Example extracted text (replace this with a real sample from your document)
extracted_text = """
2(a) Some road description goes here
(a) Main road with other details.
"""

# Regex pattern from the script
pattern = r"^2\(a\)\s*(.*?)(?:\n|$).*?\(a\)\s*(.*?)\n"

matches = re.findall(pattern, extracted_text, re.IGNORECASE | re.MULTILINE)
print(f"Matches found: {matches}")
