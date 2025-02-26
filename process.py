import os
import zipfile
import pdfplumber
from docx import Document

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

def find_zip_file(directory):
    """Find the first ZIP file in the given directory."""
    for file in os.listdir(directory):
        if file.endswith(".zip"):
            return os.path.join(directory, file)
    return None  # No ZIP file found

def process_zip(zip_path, output_docx):
    """Unzip, extract text from PDFs, and save to a Word document."""
    output_folder = "unzipped_pdfs"
    os.makedirs(output_folder, exist_ok=True)

    # Unzip the files
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)

    # Create Word document
    doc = Document()

    # Process each PDF
    for file_name in sorted(os.listdir(output_folder)):  # Sort for consistent order
        if file_name.endswith(".pdf"):
            pdf_path = os.path.join(output_folder, file_name)
            extracted_text = extract_text_from_pdf(pdf_path)

            # Add PDF file name as a heading
            doc.add_paragraph(f"Source: {file_name}", style="Heading 2")
            doc.add_paragraph(extracted_text)
            doc.add_page_break()  # Add a page break after each PDF

    # Save the final Word document
    print(f"Saving Word document to: {os.path.abspath(output_file)}")
    doc.save(output_file)

# ✅ Automatically find the ZIP file in the "input_files" folder
input_folder = "input_files"
zip_file_path = find_zip_file(input_folder)

# Define paths
output_file = "output_files/processed_doc.docx"

# Ensure output directory exists
os.makedirs("output_files", exist_ok=True)

if zip_file_path:
    print(f"Found ZIP file: {zip_file_path}")
    process_zip(zip_file_path, output_file)
else:
    print("Error: No ZIP file found in 'input_files' folder.")
