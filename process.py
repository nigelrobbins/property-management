import time
import pytesseract
from pdf2image import convert_from_path
import pdfplumber

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file, including scanned PDFs using OCR."""
    text = ""

    print(f"📄 Processing PDF: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            print(f"🔍 Extracting text from page {i+1} using pdfplumber...")
            page_text = page.extract_text()
            
            if page_text:  
                text += page_text + "\n"
            else:  
                # Perform OCR if no text is detected
                print(f"🖼️ No text found on page {i+1}, using OCR...")
                images = convert_from_path(pdf_path, first_page=i+1, last_page=i+1)  # Convert only one page at a time
                for img in images:
                    ocr_text = pytesseract.image_to_string(img)
                    text += ocr_text + "\n"
                    
                    print(f"✅ OCR done for page {i+1}, extracted {len(ocr_text)} characters.")

            # Stop processing if text contains the keyword
            if "REPLIES TO STANDARD ENQUIRIES" in text:
                print("🎯 Found keyword! Stopping early.")
                break

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
    """Unzip, extract text from PDFs and Word docs that contain the keyword, then save to a Word file."""
    print(f"📂 Unzipping {zip_path}...")

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall("unzipped_files")

    print("✅ Unzipping complete. Processing files...")

    for file_name in os.listdir("unzipped_files"):
        print(f"📄 Found file: {file_name}")

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

        # ✅ Only include documents that contain the keyword
        if extracted_text and KEYWORD in extracted_text:
            found_matching_docs = True
            doc.add_paragraph(f"Source ({file_type}): {file_name}", style="Heading 2")
            doc.add_paragraph(extracted_text)
            doc.add_page_break()

    if found_matching_docs:
        os.makedirs(os.path.dirname(output_docx), exist_ok=True)
        doc.save(output_docx)
        print(f"✅ Word document saved: {os.path.abspath(output_docx)}")
    else:
        print("⚠️ No documents containing the required text were found. No file was created.")

# ✅ Automatically find the ZIP file in "input_files"
input_folder = "input_files"
zip_file_path = find_zip_file(input_folder)

output_file = "output_files/processed_doc.docx"

if zip_file_path:
    print(f"📂 Found ZIP file: {zip_file_path}")
    process_zip(zip_file_path, output_file)
else:
    print("❌ No ZIP file found in 'input_files' folder.")
