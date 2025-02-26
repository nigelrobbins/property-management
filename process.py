import os
import zipfile
import pdfplumber
from docx import Document
from docx2txt import process as docx_extract  # Extracts text from .docx files

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text.strip()

def extract_text_from_docx(docx_path):
    """Extract text from a Word (.docx) file."""
    return docx_extract(docx_path).strip()

def find_zip_file(directory):
    """Find the first ZIP file in the given directory."""
    for file in os.listdir(directory):
        if file.endswith(".zip"):
            return os.path.join(directory, file)
    return None  # No ZIP file found

def process_zip(zip_path, output_docx):
    """Unzip, extract text from PDFs and Word docs, and save to a new Word document."""
    output_folder = "unzipped_files"
    os.makedirs(output_folder, exist_ok=True)

    # Unzip the files
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)

    # Create a new Word document
    doc = Document()

    # Process each file
    for file_name in sorted(os.listdir(output_folder)):  # Sort for consistent order
        file_path = os.path.join(output_folder, file_name)

        if file_name.endswith(".pdf"):
            extracted_text = extract_text_from_pdf(file_path)
            file_type = "PDF"
        elif file_name.endswith(".docx"):
            extracted_text = extract_text_from_docx(file_path)
            file_type = "Word Document"
        else:
            continue  # Skip other file types

        if extracted_text:
            doc.add_paragraph(f"Source ({file_type}): {file_name}", style="Heading 2")
            doc.add_paragraph(extracted_text)
            doc.add_page_break()  # Add a page break after each file

    # Ensure the output directory exists
    os.makedirs(os.path.dirname(output_docx), exist_ok=True)

    # Save the final Word document
    doc.save(output_docx)
    print(f"Word document saved: {os.path.abspath(output_docx)}")

# ✅ Automatically find the ZIP file in the "input_files" folder
input_folder = "input_files"
zip_file_path = find_zip_file(input_folder)

# Define output file path
output_file = "output_files/processed_doc.docx"

if zip_file_path:
    print(f"Found ZIP file: {zip_file_path}")
    process_zip(zip_file_path, output_file)
else:
    print("Error: No ZIP file found in 'input_files' folder.")
