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

def timed_function(func):
    """Decorator to measure function execution time."""
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        elapsed_time = end_time - start_time
        if elapsed_time > 2:
            print(f"⏱ {func.__name__} took {elapsed_time:.4f} seconds")
        return result
    return wrapper

@timed_function
def load_yaml(yaml_path):
    """Load YAML configuration and return structured data."""
    with open(yaml_path, "r", encoding="utf-8") as file:
        yaml_data = yaml.safe_load(file)
    return yaml_data

@timed_function
def clean_text(text):
    """Clean text while preserving line breaks."""
    lines = text.splitlines()
    cleaned_lines = []
    for line in lines:
        cleaned_line = re.sub(r'[^a-zA-Z0-9\s\n*()\-,.:;?!\'"]', '', line)
        cleaned_lines.append(cleaned_line)
    return "\n".join(cleaned_lines)

@timed_function
def extract_text_from_pdf(pdf_path):
    """Extract text from PDF using multiple methods."""
    # [Previous implementation remains the same]
    # ...

@timed_function
def extract_text_from_docx(docx_path):
    """Extract text from Word document."""
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

@timed_function
def add_formatted_paragraph(doc, text, style=None, bold=False, italic=False):
    """Add formatted paragraph to document."""
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    return p

@timed_function
def process_sections(doc, sections, level=2, extracted_text=""):
    """Recursively process document sections."""
    for section in sections:
        if 'section' in section:  # Only process if it's a section
            add_formatted_paragraph(doc, section['section'], style=f'Heading {level}', bold=True)
            
            if section['search_pattern'] in extracted_text and section['extract_text']:
                match = extract_matching_text(extracted_text, section['extract_pattern'], section['message_template'])
                if match:
                    add_formatted_paragraph(doc, match, italic=True)
                else:
                    add_formatted_paragraph(doc, "No matching content found.", style='Intense Quote')
            else:
                add_formatted_paragraph(doc, f"No {section['section']} information found.", style='Intense Quote')
            
            if 'sections' in section:
                process_sections(doc, section['sections'], level+1, extracted_text)

@timed_function
def extract_matching_text(text, pattern, message_template):
    """Extract and format text using regex pattern."""
    matches = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if matches:
        extracted = {f"extracted_text_{i+1}": matches.group(i+1) for i in range(matches.lastindex)}
        return message_template.format(**extracted)
    return None

@timed_function
def generate_report(doc, yaml_data, extracted_text):
    """Generate report document from YAML structure."""
    # Add title and scope
    doc.add_heading(yaml_data['general']['title'], level=0)
    scope = yaml_data['general']['scope'][0]
    doc.add_heading(scope['heading'], level=1)
    doc.add_paragraph(scope['body'])
    
    # Process each document section
    for doc_section in yaml_data['docs']:
        doc.add_heading(doc_section['heading'], level=1)
        
        # Check if identifier exists in text
        if doc_section['identifier'] in extracted_text:
            doc.add_paragraph(doc_section['message_if_identifier_found'])
        else:
            doc.add_paragraph(doc_section['message_if_identifier_not_found'])
        
        # Process questions
        for question in doc_section['questions']:
            if 'address' in question:
                # Process address
                add_formatted_paragraph(doc, question['address'], style='Heading 2')
                if question['search_pattern'] in extracted_text and question['extract_text']:
                    address = extract_matching_text(extracted_text, question['extract_pattern'], question['message_template'])
                    if address:
                        add_formatted_paragraph(doc, address, italic=True)
            
            if 'sections' in question:
                process_sections(doc, question['sections'], extracted_text=extracted_text)

@timed_function
def process_zip(zip_path, output_docx, yaml_path):
    """Process ZIP file containing documents."""
    output_folder = "output_files/unzipped_files"
    os.makedirs(output_folder, exist_ok=True)
    
    # Load YAML config
    yaml_data = load_yaml(yaml_path)
    doc = Document()
    
    # Extract ZIP
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)
    
    # Process each file
    for file_name in os.listdir(output_folder):
        file_path = os.path.join(output_folder, file_name)
        
        if file_name.endswith(".pdf"):
            extracted_text = extract_text_from_pdf(file_path)
        elif file_name.endswith(".docx"):
            extracted_text = extract_text_from_docx(file_path)
        else:
            continue
        
        # Generate report
        generate_report(doc, yaml_data, extracted_text)
        doc.add_page_break()
    
    # Save final document
    os.makedirs(os.path.dirname(output_docx), exist_ok=True)
    doc.save(output_docx)

# Main execution
if __name__ == "__main__":
    input_folder = "input_files"
    yaml_config = "config.yaml"
    output_file = "output_files/processed_doc.docx"
    
    # Find ZIP file
    zip_file_path = next(
        (os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith(".zip")),
        None
    )
    
    if zip_file_path:
        print(f"📂 Found ZIP file: {zip_file_path}")
        process_zip(zip_file_path, output_file, yaml_config)
    else:
        print("❌ No ZIP file found in 'input_files' folder.")
