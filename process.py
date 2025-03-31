import os
import zipfile
import pdfplumber
import re
import time
import yaml
from docx import Document

def load_yaml(yaml_path):
    """Load and parse the YAML file, storing relevant details in a structured format."""
    with open(yaml_path, "r", encoding="utf-8") as file:
        yaml_data = yaml.safe_load(file)

    parsed_data = {
        "title": yaml_data.get("general", {}).get("title", ""),
        "scope": yaml_data.get("general", {}).get("scope", []),
        "docs": yaml_data.get("docs", []),
        "none_subsections": yaml_data.get("none", {}).get("none_subsections", []),
        "all_none_message": yaml_data.get("none", {}).get("all_none_message", ""),
        "all_none_section": yaml_data.get("none", {}).get("all_none_section", ""),
    }

    return parsed_data

def extract_text_from_pdf(pdf_path):
    """Extract text from PDF using pdfplumber."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text.strip()

def extract_text_from_docx(docx_path):
    """Extract text from a Word document."""
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

def identify_document_type(text, docs):
    """Determine the type of document based on identifiers from YAML."""
    for doc_type in docs:
        if doc_type["identifier"] in text:
            return doc_type
    return None

def extract_matching_text(text, pattern, template):
    """Extract text using regex and format with a template."""
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if match:
        extracted_values = [match.group(i) for i in range(1, match.lastindex + 1)]
        return template.format(**{f"extracted_text_{i+1}": extracted_values[i] for i in range(len(extracted_values))})
    return None

def generate_report(doc, parsed_data, extracted_text):
    """Generate the Word document using parsed YAML data and extracted text."""
    doc.add_paragraph(parsed_data["title"], style="Heading 1")

    if parsed_data["scope"]:
        first_scope = parsed_data["scope"][0]  # Assuming only one scope is needed
        doc.add_paragraph(first_scope.get("heading", ""), style="Heading 2")
        doc.add_paragraph(first_scope.get("body", ""), style="Normal")

    doc_type = identify_document_type(extracted_text, parsed_data["docs"])
    if not doc_type:
        doc.add_paragraph("No relevant document type identified.", style="Normal")
        return

    doc.add_paragraph(doc_type["heading"], style="Heading 1")
    doc.add_paragraph(doc_type["message_if_identifier_found"], style="Normal")

    for question in doc_type["questions"]:
        section_name = question.get("section", "")
        doc.add_paragraph(section_name, style="Heading 2")

        if question["search_pattern"] in extracted_text:
            if question["extract_text"]:
                extracted_content = extract_matching_text(
                    extracted_text, question["extract_pattern"], question["message_template"]
                )
                if extracted_content:
                    doc.add_paragraph(extracted_content, style="Normal")
                else:
                    doc.add_paragraph("No relevant content found.", style="Normal")
        else:
            doc.add_paragraph(f"No {question.get('subsection', 'information')} found.", style="Normal")

def process_zip(zip_path, output_docx, yaml_path):
    """Extract and process documents inside a ZIP file."""
    output_folder = "output_files/unzipped_files"
    os.makedirs(output_folder, exist_ok=True)

    parsed_data = load_yaml(yaml_path)
    doc = Document()

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)

    for file_name in sorted(os.listdir(output_folder)):
        file_path = os.path.join(output_folder, file_name)

        if file_name.endswith(".pdf"):
            extracted_text = extract_text_from_pdf(file_path)
        elif file_name.endswith(".docx"):
            extracted_text = extract_text_from_docx(file_path)
        else:
            continue

        generate_report(doc, parsed_data, extracted_text)
        doc.add_page_break()

    doc.save(output_docx)

# Run the script
input_folder = "input_files"
yaml_config = "config.yaml"
output_file = "output_files/processed_doc.docx"

for file in os.listdir(input_folder):
    if file.endswith(".zip"):
        zip_file_path = os.path.join(input_folder, file)
        process_zip(zip_file_path, output_file, yaml_config)
        break
