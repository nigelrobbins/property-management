import os
import zipfile
import shutil
import pdfplumber
from docx import Document
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

SEARCH_PHRASE = "REPLIES TO STANDARD ENQUIRIES"

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file, including scanned PDFs using OCR."""
    text = ""
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:  
                text += page_text + "\n"
            else:  
                # Perform OCR if no text is detected
                images = convert_from_path(pdf_path)  # Convert PDF pages to images
                for img in images:
                    text += pytesseract.image_to_string(img) + "\n"

    return text.strip()

def extract_text_from_docx(docx_path):
    """Extract text from a Word (.docx) file."""
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

def find_zip_file(directory):
    """Find the first ZIP file in the given directory."""
    print(f"📂 Checking directory: {directory}")

    if not os.path.exists(directory):  
        print(f"❌ ERROR: Directory does not exist: {directory}")
        return None

    try:
        files = os.listdir(directory)  
        print(f"📄 Files in {directory}: {files}")

        for file in files:
            print(f"🔍 Checking file: {file}")  
            if file.endswith(".zip"):
                print(f"✅ Found ZIP file: {file}")
                return os.path.join(directory, file)

    except Exception as e:
        print(f"❌ ERROR while listing files: {e}")  
        return None

    print("❌ No ZIP file found in the directory.")
    return None  

def process_zip(zip_path, output_docx):
    """Unzip, extract text from PDFs and Word docs, save to a Word file, and move ZIP file."""
    output_folder = "unzipped_files"
    processed_folder = "processed_files"

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
    
    processed_any = False  # Track if any document met the condition

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

        # Check if the required phrase is in the extracted text
        if SEARCH_PHRASE in extracted_text:
            print(f"✅ Processing {file_name} (contains required phrase)")
            doc.add_paragraph(f"Source ({file_type}): {file_name}", style="Heading 2")
            doc.add_paragraph(extracted_text)
            doc.add_page_break()
            processed_any = True
        else:
            print(f"⏭ Skipping {file_name} (does not contain required phrase)")

    if processed_any:
        os.makedirs(os.path.dirname(output_docx), exist_ok=True)
        doc.save(output_docx)
        print(f"✅ Word document saved: {os.path.abspath(output_docx)}")
    else:
        print("⚠️ No files contained the required phrase. No document was created.")

# ✅ Automatically find the ZIP file in "input_files"
input_folder = "input_files"
print("start")
zip_file_path = find_zip_file(input_folder)

output_file = "output_files/processed_doc.docx"
if zip_file_path:
    print(f"📂 Found ZIP file: {zip_file_path}")
    process_zip(zip_file_path, output_file)
else:
    print("❌ No ZIP file found in 'input_files' folder.")
