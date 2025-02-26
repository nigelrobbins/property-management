import os
import zipfile
import shutil
import pdfplumber
from docx import Document

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
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

def find_zip_file(directory):
    """Find the first ZIP file in the given directory."""
    for file in os.listdir(directory):
        if file.endswith(".zip"):
            return os.path.join(directory, file)
    return None  # No ZIP file found

def process_zip(zip_path, output_docx):
    """Unzip, extract text from PDFs and Word docs, save to a Word file, and move ZIP file."""
    output_folder = "unzipped_files"
    processed_folder = "processed_files"  # ✅ Destination for processed ZIPs

    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(processed_folder, exist_ok=True)  # ✅ Ensure processed folder exists

    if not os.path.exists(zip_path):
        print(f"❌ ERROR: ZIP file does not exist: {zip_path}")
        return

    print(f"📂 Unzipping: {zip_path}")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)

    doc = Document()

    for file_name in sorted(os.listdir(output_folder)):
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
            doc.add_page_break()

    os.makedirs(os.path.dirname(output_docx), exist_ok=True)
    doc.save(output_docx)
    print(f"✅ Word document saved: {os.path.abspath(output_docx)}")

    # ✅ Move ZIP file after processing
    processed_zip_path = os.path.join(processed_folder, os.path.basename(zip_path))
    
    print(f"🔄 Moving ZIP file to: {processed_zip_path}")
    shutil.move(zip_path, processed_zip_path)

    if os.path.exists(processed_zip_path):
        print(f"✅ Successfully moved ZIP file to: {processed_zip_path}")
    else:
        print(f"❌ ERROR: Failed to move ZIP file.")

# ✅ Automatically find the ZIP file in "input_files"
input_folder = "input_files"
zip_file_path = find_zip_file(input_folder)

output_file = "output_files/processed_doc.docx"

if zip_file_path:
    print(f"📂 Found ZIP file: {zip_file_path}")
    process_zip(zip_file_path, output_file)
else:
    print("❌ No ZIP file found in 'input_files' folder.")
