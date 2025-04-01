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
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def load_yaml(yaml_path):
    with open(yaml_path, 'r') as file:
        return yaml.safe_load(file)

def add_formatted_paragraph(doc, text, style=None, bold=False, italic=False):
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    return p

def process_sections(doc, sections, level=2):
    for section in sections:
        # Add section heading
        add_formatted_paragraph(doc, section['section'], style=f'Heading {level}', bold=True)
        
        # Process extracted content if available
        if 'extracted_text' in section and section['extracted_text']:
            # Apply message template if available
            if 'message_template' in section:
                try:
                    message = section['message_template'].format(**section['extracted_text'])
                    add_formatted_paragraph(doc, message, italic=True)
                except KeyError as e:
                    add_formatted_paragraph(doc, f"Error in message template: missing key {e}", style='Intense Quote')
            else:
                add_formatted_paragraph(doc, section['extracted_text'], style='Intense Quote')
        else:
            add_formatted_paragraph(doc, "No information found for this section.", style='Intense Quote')
        
        # Process nested sections if they exist
        if 'sections' in section:
            process_sections(doc, section['sections'], level=level+1)

def generate_document(yaml_data, output_path):
    doc = Document()
    
    # Add title
    doc.add_heading(yaml_data['general']['title'], level=0)
    
    # Add scope section
    scope = yaml_data['general']['scope'][0]
    doc.add_heading(scope['heading'], level=1)
    doc.add_paragraph(scope['body'])
    
    # Process each document section
    for doc_section in yaml_data['docs']:
        doc.add_heading(doc_section['heading'], level=1)
        
        # Process address section separately if it exists
        if 'questions' in doc_section:
            for question in doc_section['questions']:
                if 'address' in question:
                    # Address section
                    add_formatted_paragraph(doc, question['address'], style='Heading 2', bold=True)
                    if 'extracted_text' in question and question['extracted_text']:
                        try:
                            message = question['message_template'].format(**question['extracted_text'])
                            add_formatted_paragraph(doc, message, italic=True)
                        except KeyError as e:
                            add_formatted_paragraph(doc, f"Error in address template: missing key {e}", style='Intense Quote')
                    else:
                        add_formatted_paragraph(doc, "No address information found.", style='Intense Quote')
                
                # Process sections
                if 'sections' in question:
                    process_sections(doc, question['sections'])
    
    # Save the document
    doc.save(output_path)

def parse_and_generate(yaml_path, output_docx):
    yaml_data = load_yaml(yaml_path)
    generate_document(yaml_data, output_docx)

# Deepseek above

def timed_function(func):
    """Decorator to measure function execution time and log only if it exceeds 2 seconds."""
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        elapsed_time = end_time - start_time

        if elapsed_time > 2:  # Log only if execution time exceeds 2 seconds
            print(f"⏱ {func.__name__} took {elapsed_time:.4f} seconds")

        return result
    return wrapper

@timed_function
def clean_text(text):
    """Cleans text, removes odd characters but keeps blank lines intact."""
    
    # Split text into lines to keep blank lines intact
    lines = text.splitlines()

    cleaned_lines = []
    for line in lines:
        # Remove unwanted characters, but allow spaces and alphanumeric characters
        cleaned_line = re.sub(r'[^a-zA-Z0-9\s\n*()\-,.:;?!\'"]', '', line)
        cleaned_lines.append(cleaned_line)

    # Join cleaned lines back together, preserving blank lines
    return "\n".join(cleaned_lines)

@timed_function
def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF, using pdftotext first, then pdfplumber, then OCR if needed."""
    
    # Ensure the work_files directory exists
    output_dir = "work_files"
    os.makedirs(output_dir, exist_ok=True)

    # Construct the output file path
    output_file_path = os.path.join(output_dir, os.path.basename(pdf_path) + ".txt")

    # Try using pdftotext first
    result = subprocess.run(['pdftotext', pdf_path, '-'], capture_output=True, text=True)
    text = result.stdout.strip()

    if text:
        print(f"✅ Extracted text using pdftotext: {text[:100]}...")
        with open(output_file_path, "w", encoding="utf-8") as f:
            f.write(text)
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
        print(f"✅ Extracted text using pdfplumber: {text[:100]}...")
        text = text.strip()
        with open(output_file_path, "w", encoding="utf-8") as f:
            f.write(text)
        return text  # If pdfplumber works, return immediately

    print("⚠️ pdfplumber failed, performing OCR...")

    # Final fallback: Use OCR (slow)
    text = ""
    images = convert_from_path(pdf_path)
    for img in images:
        ocr_text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
        cleaned_text = ocr_text.strip()
        text += cleaned_text + "\n"

    text = text.strip()
    print(f"✅ Extracted text using OCR (cleaned): {text[:100]}...")

    # Write the final extracted text to the file
    with open(output_file_path, "w", encoding="utf-8") as f:
        f.write(text)

    return text

@timed_function
def extract_text_from_docx(docx_path):
    """Extract text from a Word document."""
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

# Load YAML configuration
@timed_function
def load_yaml(yaml_path):
    """Load questions and settings from a YAML file."""
    with open(yaml_path, "r", encoding="utf-8") as file:
        yaml_data = yaml.safe_load(file)

    general = yaml_data.get("general", {})
    title = general.get("title", "")
    scope = general.get("scope", [])

    # Extract first scope item (if available)
    heading, body = None, None
    if scope and isinstance(scope, list):
        first_scope = scope[0]  # Assuming you need only the first scope entry
        heading = first_scope.get("heading", "")
        body = first_scope.get("body", "")

    docs = yaml_data.get("docs", [])
    none_subsections = yaml_data.get("none", {}).get("none_subsections", [])
    all_none_message = yaml_data.get("none", {}).get("all_none_message", None)
    all_none_section = yaml_data.get("none", {}).get("all_none_section", None)

    return title, heading, body, docs, none_subsections, all_none_message, all_none_section

# Identify question group based on document content
@timed_function
def identify_group(text, docs):
    for group in docs:
        if group["identifier"] in text:
            return group
    return None  # No matching group found

@timed_function
def extract_matching_text(text, pattern, message_template):
    """Extracts matching text dynamically based on the given pattern and formats the message."""

    print(f"🔍 Extracting with pattern: {pattern}")

    matches = re.search(pattern, text, re.IGNORECASE | re.DOTALL)

    if matches:
        extracted_texts = [matches.group(i) for i in range(1, matches.lastindex + 1)]
        
        # Log extracted values
        print(f"✅ Extracted text values: {extracted_texts}")

        # Dynamically format message template with extracted values
        formatted_message = message_template.format(**{f"extracted_text_{i+1}": extracted_texts[i] for i in range(len(extracted_texts))})

        print(f"✅ Formatted message: {formatted_message}")
        return formatted_message
    else:
        print("⚠️ No matches found.")
        return None

@timed_function
def find_subsection_message_not_found(question):
    """Find the 'message_not_found' from a relevant subsection if it exists."""
    if "subsections" in question:
        for subsection in question["subsections"]:
            if "message_not_found" in subsection:
                return subsection["message_not_found"]
    return "No relevant information found."  # Default fallback message

@timed_function
def process_questions(doc, extracted_text, questions, message_if_identifier_found, none_subsections, all_none_message, all_none_section, section_name=""):
    """Recursively process questions and their subsections."""
    extracted_text_2_values = {}  # Store extracted_text_2 for specified subsections
    section_logged = False  # Ensure "Other Matters" is added only once

    for question in questions:
        if section_name != question.get("section", section_name):
            section_name = question.get("section", section_name) 
            doc.add_paragraph(section_name, style="Heading 2")

        if question["search_pattern"] in extracted_text:
            if question["extract_text"]:
                extracted_section = extract_matching_text(
                    extracted_text, question["extract_pattern"], question["message_template"]
                )
                if extracted_section:
                    if "subsection" in question:
                        doc.add_paragraph(question["subsection"], style="Heading 3")
                    print(f"✅ Extracted content: {extracted_section[:50]}...")  # Debugging
                    paragraph = doc.add_paragraph(extracted_section)
                    paragraph.runs[0].italic = True
                    if message_if_identifier_found not in ["", None]:
                        doc.add_paragraph(message_if_identifier_found, style="Normal")
                        message_if_identifier_found = ""

                    # Check if the subsection is listed in the YAML
                    if "subsection" in question:
                        if question["subsection"] in none_subsections:
                            matches = re.search(question["extract_pattern"], extracted_text, re.IGNORECASE | re.DOTALL)
                            extracted_text_2 = matches[0][1] if matches and len(matches[0]) > 1 else None
                            extracted_text_2_values[question["subsection"]] = extracted_text_2
                else:
                    doc.add_paragraph("⚠️ No matching content found.", style="Normal")
        else:
            if "subsection" in question:
                doc.add_paragraph(f"No {question['subsection']} information found.", style="Normal")
        
        # ✅ Ensure "Other Matters" is only added once before logging `all_none_message`
        if section_name == all_none_section and not section_logged:
            section_logged = True
            # ✅ Dynamically check if we are in the correct section from YAML before logging the message
            if all(extracted_text_2_values.get(sub) is None for sub in none_subsections):
                doc.add_paragraph(all_none_message, style="Normal")

        if "subsections" in question and question["subsections"]:
            process_questions(doc, extracted_text, question["subsections"], message_if_identifier_found, none_subsections, all_none_message, all_none_section, section_name)

@timed_function
def process_zip(zip_path, output_docx, yaml_path):
    """Extract and process only relevant sections from documents that contain filter text."""
    output_folder = "output_files/unzipped_files"
    os.makedirs(output_folder, exist_ok=True)
    title, heading, body, docs, none_subsections, all_none_message, all_none_section = load_yaml(yaml_path)
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

    extracted_text_files = []

    for file_name in sorted(os.listdir(output_folder)):
        file_path = os.path.join(output_folder, file_name)

        print(f"📄 Processing {file_name}...")
        process_start = time.time()

        if file_name.endswith(".pdf"):
            extracted_text = extract_text_from_pdf(file_path)
        elif file_name.endswith(".docx"):
            extracted_text = extract_text_from_docx(file_path)
        else:
            continue

        # Save extracted text to a file
        extracted_text_file = f"{file_path}.txt"
        with open(extracted_text_file, "w", encoding="utf-8") as f:
            f.write(extracted_text)
        extracted_text_files.append(extracted_text_file)

        group = identify_group(extracted_text, docs)
        if group is None:
            print("⚠️ No matching group found. Skipping this document.")
            continue  # Skip processing this file

# use DeepSeek here
        parse_and_generate(yaml_path, output_docx)

    # Create ZIP file including extracted text files
    final_zip_path = "output_files/processed_files.zip"
    with zipfile.ZipFile(final_zip_path, 'w') as zipf:
        for txt_file in extracted_text_files:
            print(f"📄 Adding extracted text file: {txt_file}")
            zipf.write(txt_file, os.path.basename(txt_file))

    print(f"✅ Final ZIP created: {final_zip_path}")

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
