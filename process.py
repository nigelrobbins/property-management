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
from datetime import datetime  # Import datetime module

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

    groups = yaml_data.get("groups", [])
    check_none_subsections = yaml_data.get("config", {}).get("check_none_subsections", [])  # Correct path
    all_none_message = yaml_data.get("config", {}).get("all_none_message", None)  # Correct path
    log_message_section = yaml_data.get("config", {}).get("log_message_section", None)  # Correct path

    return groups, check_none_subsections, all_none_message, log_message_section

# Identify question group based on document content
@timed_function
def identify_group(text, groups):
    for group in groups:
        if group["identifier"] in text:
            return group
    return None  # No matching group found

@timed_function
def extract_matching_text(text, pattern, message_template):
    """Extracts matching text based on the given pattern and formats the message."""
    # Log the text and pattern for debugging
    print(f"🔍 Extracting with pattern: {pattern}")
    print(f"🔍 Text to search: {text}")

    # Find the matching text based on the pattern
    matches = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    
    if matches:
        # Log the matches for debugging
        print(f"✅ Matches found: {matches}")
        print("Full match:", matches.group(0))  # Debugging
        print("First group:", matches.group(1)
        print("Second group:", matches.group(2))  # This should be "Wanted"    
        
        extracted_text_1 = matches.group(1))  # First part of the extracted text
        extracted_text_2 = matches.group(2) if len(matches[0]) > 1 else ''  # Second part of the extracted text (optional)

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

@timed_function
def find_subsection_message_not_found(question):
    """Find the 'message_not_found' from a relevant subsection if it exists."""
    if "subsections" in question:
        for subsection in question["subsections"]:
            if "message_not_found" in subsection:
                return subsection["message_not_found"]
    return "No relevant information found."  # Default fallback message

@timed_function
def process_questions(doc, extracted_text, questions, check_none_subsections, all_none_message, log_message_section, section_name=""):
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
                    doc.add_paragraph(question["subsection"], style="Heading 3")
                    print(f"✅ Extracted content: {extracted_section[:50]}...")  # Debugging
                    paragraph = doc.add_paragraph(extracted_section)
                    paragraph.runs[0].italic = True

                    # Check if the subsection is listed in the YAML
                    if question["subsection"] in check_none_subsections:
                        matches = re.search(question["extract_pattern"], extracted_text, re.IGNORECASE | re.DOTALL)
                        extracted_text_2 = matches[0][1] if matches and len(matches[0]) > 1 else None
                        extracted_text_2_values[question["subsection"]] = extracted_text_2
                else:
                    doc.add_paragraph("⚠️ No matching content found.", style="Normal")
        else:
            doc.add_paragraph(f"No {question['subsection']} information found.", style="Normal")
        
        # ✅ Ensure "Other Matters" is only added once before logging `all_none_message`
        if section_name == log_message_section and not section_logged:
            section_logged = True
            # ✅ Dynamically check if we are in the correct section from YAML before logging the message
            if all(extracted_text_2_values.get(sub) is None for sub in check_none_subsections):
                doc.add_paragraph(all_none_message, style="Normal")

        if "subsections" in question and question["subsections"]:
            process_questions(doc, extracted_text, question["subsections"], check_none_subsections, all_none_message, log_message_section, section_name)

@timed_function
def process_zip(zip_path, output_docx, yaml_path):
    """Extract and process only relevant sections from documents that contain filter text."""
    output_folder = "output_files/unzipped_files"
    os.makedirs(output_folder, exist_ok=True)
    groups, check_none_subsections, all_none_message, log_message_section = load_yaml(yaml_path)  # Load YAML data
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

        group = identify_group(extracted_text, groups)
        if group is None:
            print("⚠️ No matching group found. Skipping this document.")
            continue  # Skip processing this file

        doc.add_paragraph(group["heading"], style="Heading 1")
        # Add today's date in a readable format
        today_date = datetime.today().strftime("%d %B %Y")  # Example: "18 March 2025"
        doc.add_paragraph(f"Date: {today_date}", style="Normal")
        if group:
            doc.add_paragraph(group["message_if_identifier_found"], style="Normal")
            print(group["message_if_identifier_found"])
        else:
            doc.add_paragraph(group["message_if_identifier_not_found"], style="Normal")
            print("⚠️ No matching group found. Skipping.")
            continue  # Skip this file if no match

        for question in group["questions"]:
            doc.add_paragraph(question.get("message_found", ""), style="Normal")

        # 🔹 **Use the recursive function here**
        process_questions(doc, extracted_text, group["questions"], check_none_subsections, all_none_message, log_message_section, section_name="")
        doc.add_page_break()

    # Save Word document
    os.makedirs(os.path.dirname(output_docx), exist_ok=True)
    doc.save(output_docx)

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
