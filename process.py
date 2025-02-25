import os
import pdfplumber
from docx import Document

# Define paths
unzipped_folder = "input_files/unzipped/"
output_file = "output_files/processed_doc.docx"

# Ensure output directory exists
os.makedirs("output_files", exist_ok=True)

# Create a new Word document
doc = Document()

# Process each PDF file
for file in os.listdir(unzipped_folder):
    if file.endswith(".pdf"):
        pdf_path = os.path.join(unzipped_folder, file)
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    doc.add_paragraph(text)

# Save the Word document
doc.save(output_file)
print(f"✅ Word document saved: {output_file}")
