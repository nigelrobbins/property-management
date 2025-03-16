import os
import zipfile
import pdfplumber
import re
import time
import subprocess
import yaml
import unicodedata
import pytesseract
from docx import Document
from pdf2image import convert_from_path
from PIL import Image

def timed_function(func):
    """Decorator to measure function execution time."""
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        print(f"⏱ {func.__name__} took {end_time - start_time:.4f} seconds")
        return result
    return wrapper

def clean_text(text):
    """Remove odd characters from OCR output."""
    # Normalize Unicode characters
    text = unicodedata.normalize("NFKC", text)
    
    # Remove non-printable characters and excessive spaces
    text = re.sub(r'[^\x20-\x7E]', '', text)  # Keep only standard ASCII (printable)
    text = re.sub(r'\s+', ' ', text).strip()  # Replace multiple spaces with a single space
    
    return text

@timed_function
def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF, using pdftotext first, then pdfplumber, then OCR if needed."""

    # Try using pdftotext first
    result = subprocess.run(['pdftotext', pdf_path, '-'], capture_output=True, text=True)
    text = result.stdout.strip()

    if text:
        print(f"✅ Extracted text using pdftotext: {text[:100]}...")  # Show first 100 characters
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
        print(f"✅ Extracted text using pdfplumber: {text[:100]}...")  # Show first 100 characters
        return text.strip()  # If pdfplumber works, return immediately

    print("⚠️ pdfplumber failed, performing OCR...")

    # Final fallback: Use OCR (slow)
    text = ""
    images = convert_from_path(pdf_path)
    for img in images:
        ocr_text = pytesseract.image_to_string(img)
        cleaned_text = clean_text(ocr_text)  # Apply cleaning function
        text += cleaned_text + "\n"

    print(f"✅ Extracted text using OCR (cleaned): {text[:100]}...")  # Show first 100 characters
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

import re

def extract_matching_text(text, pattern, message_template):
    """Extracts matching text based on the given pattern and formats the message."""
    # Log the text and pattern for debugging
    print(f"🔍 Extracting with pattern: {pattern}")
    print(f"🔍 Text to search: {text}")

    # Find the matching text based on the pattern
    matches = re.findall(pattern, text, re.IGNORECASE | re.DOTALL)
    
    if matches:
        # Log the matches for debugging
        print(f"✅ Matches found: {matches}")
        
        extracted_text_1 = matches[0][0]  # First part of the extracted text
        extracted_text_2 = matches[0][1] if len(matches[0]) > 1 else ''  # Second part of the extracted text (optional)

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
                    extracted_section = extract_matching_text(extracted_text, question["extract_pattern"], question["message_template"])
                    if extracted_section:
                        print(f"✅ Extracted content: {extracted_section[:50]}...")  # Log first 50 characters
                        paragraph = doc.add_paragraph(extracted_section)
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
Council

REPLIES TO STANDARD ENQUIRIES
OF LOCAL AUTHORITY (2007 Edition)

Applicant: Mr Kevin Bird
1200, Delta Business Park
Swindon
SN5 7XZ

Search Reference: 1415 03258

Enfield Council |Civic Centre | Silver Street | Enfield |EN1 3XA | DX 90615 Enfield
landcharges@enfield.gov.uk | 020 8379 1000

Page 2 of 9

41.2.

Unless otherwise indicated, matters will be disclosed only if they apply directly to the property described in
Box B.

‘Area’ means any area in which the property is located.

References to ‘the Council’ include any predecessor Council and also any Council Committee, sub-
committee or other body or person exercising powers delegated by the Council and their ‘approval’ includes
their decision to proceed. The replies given to certain enquiries cover knowledge and actions of both the
District Council and the County Council.

References to the provisions of particular Acts of Parliament or Regulations include any provisions which
they have replaced and also include existing or future amendments or re-enactments.

The replies will be given in the belief that they are in accordance with information presently available to the
officers of the replying Council, but none of the Councils or their officers accept legal responsibility for an
incorrect reply, except for negligence. Any liability for negligence will extend to the person who raised the
enquiries and the person on whose behalf they were raised. It will also extend to any other person who has
knowledge (personally or through an agent) of the replies before the time when he purchases, takes
tenancy of, or lends money on the security of the property or (if earlier) the time when he becomes
contractually bound to do so.

INFORMATION REGARDING LOCAL PLANS WILL FOLLOW
Planning Designations and Proposals

1.2 What designations of land use for the property or the area, and what specific proposals for
the

property, are contained in any exisiting or proposed development plan?

None

The Entield Plan - Core Strategy was submitted to the Secretary of State on the 16th March 2010 and the
Council adopted the Core Strategy on the 10th November 2010. The Development Plan for the Local
Authority now comprises of (i) The Enfield Plan Core Strategy adopted November 2070 (ii) the saved
policies of the 1994 London Borough of Enfield Unitary Development Plan as updated November 2010 (iii)
The London Plan, including alterations, 2008.

The Council is continuing to prepare more planning documents as part of the Local Development
Framework. Further details of the document to be prepared, as part of the LDF, are set out in the Local
Development Scheme also available on the Council website www.enfield.gov.uk

if you wish to obtain further details on this matter, please contact the Planning Policy Team on 020 8379
1000 or via email to planningpolicy@enfield.gov.uk

2. ROADS

Which of the roads, footways and footpaths named in the application for this search (via boxes B
and C) are:

2(a) Highways maintainable at public expense;

(a) Gordon Road is publicly maintained.

2(b) Subject to adoption and, supported by a bond or bond waiver;

(b) Not applicable

2(c) To be made up by the local authority who will reclaim the cost from the frontagers; or

(c) Not applicable
"""

# Regex pattern from the script
pattern = r"2\(a\)\s*(.*?)(?:\n|$).*?\(a\)\s*(.*?)\n"

message_template = "{extracted_text_1}. The main road ({extracted_text_2}) is a highway maintainable at public expense. A highway maintainable at public expense is a local highway. The local authority is responsible for maintaining the road, including repairs, resurfacing, and other works. It will be maintained according to the standards of the local authority and you will have access to it."

formatted_message = extract_matching_text(extracted_text, pattern, message_template)
print(f"Formatted message: {formatted_message}")
