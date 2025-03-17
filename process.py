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

def timed_function(func):
    """Decorator to measure function execution time."""
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        print(f"⏱ {func.__name__} took {end_time - start_time:.4f} seconds")
        return result
    return wrapper

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

    # Try using pdftotext first
    result = subprocess.run(['pdftotext', pdf_path, '-'], capture_output=True, text=True)
    text = result.stdout.strip()

    if text:
        print(f"✅ Extracted text using pdftotext: {text[:100]}...")  # Show first 100 characters
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
        print(f"✅ Extracted text using pdfplumber: {text[:100]}...")  # Show first 100 characters
        return text.strip()  # If pdfplumber works, return immediately

    print("⚠️ pdfplumber failed, performing OCR...")

    # Final fallback: Use OCR (slow)
    text = ""
    images = convert_from_path(pdf_path)
    for img in images:
        ocr_text = pytesseract.image_to_string(img, lang='eng', config='--oem 3 --psm 6')
        ocr_text = ocr_text.encode("utf-8").decode("utf-8")  # Ensure UTF-8 encoding
        cleaned_text = clean_text(ocr_text)  # Apply cleaning function
        text += cleaned_text + "\n"

    print(f"✅ Extracted text using OCR (cleaned): {text[:100]}...")  # Show first 100 characters
    return text.strip()

def extract_text_from_docx(docx_path):
    """Extract text from a Word document."""
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

# Load YAML configuration
def load_yaml(yaml_path):
    with open(yaml_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)["groups"]

# Identify question group based on document content
def identify_group(text, groups):
    for group in groups:
        if group["identifier"] in text:
            return group
    return None  # No matching group found

import re

def extract_matching_text(text, pattern, message_template):
    """Extracts matching text based on the given pattern and formats the message."""
    # Log the text and pattern for debugging
    print(f"🔍 Extracting with pattern: {pattern}")
    print(f"🔍 Text to search: {text}")

    # Find the matching text based on the pattern
    matches = re.findall(pattern, text, re.IGNORECASE | re.DOTALL)
    
    if matches:
        # Log the matches for debugging
        print(f"✅ Matches found: {matches}")
        
        extracted_text_1 = matches[0][0]  # First part of the extracted text
        extracted_text_2 = matches[0][1] if len(matches[0]) > 1 else ''  # Second part of the extracted text (optional)

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

def process_zip(zip_path, output_docx, yaml_path):
    """Extract and process only relevant sections from documents that contain filter text."""
    output_folder = "unzipped_files"
    output_folder = "output_files/unzipped_files"

    os.makedirs(output_folder, exist_ok=True)
    groups = load_yaml(yaml_path)
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

    doc.add_paragraph(f"ZIP File: {os.path.basename(zip_path)}", style="Heading 1")
    extracted_text_files = []  # Store paths to extracted text files

    for file_name in sorted(os.listdir(output_folder)):
        file_path = os.path.join(output_folder, file_name)

        print(f"📄 Processing {file_name}...")
        process_start = time.time()

        if file_name.endswith(".pdf"):
            extracted_text = extract_text_from_pdf(file_path)
            file_type = "PDF"
        elif file_name.endswith(".docx"):
            extracted_text = extract_text_from_docx(file_path)
            file_type = "Word Document"
        else:
            continue

        # Save extracted text to a .txt file
        extracted_text_file = f"{file_path}.txt"
        with open(extracted_text_file, "w", encoding="utf-8") as f:
            f.write(extracted_text)
        extracted_text_files.append(extracted_text_file)

        group = identify_group(extracted_text, groups)

        if group:
            doc.add_paragraph(group["message_if_identifier_found"], style="Heading 2")
            print(group["message_if_identifier_found"])
        else:
            print("⚠️ No matching group found. Skipping.")
            continue  # Skip this file if no match
        
        doc.add_paragraph(f"📂 Processing: {file_name}", style="Heading 1")
        doc.add_paragraph(f"📄 Document identified as: {group['name']}", style="Heading 2")

        for question in group["questions"]:
            doc.add_paragraph(f"🔍 Checking section: {question['section']}", style="Heading 3")

            if question["search_pattern"] in extracted_text:
                doc.add_paragraph(question["message_found"], style="Normal")
                if question["extract_text"]:
                    extracted_section = extract_matching_text(extracted_text, question["extract_pattern"], question["message_template"])
                    if extracted_section:
                        print(f"✅ Extracted content: {extracted_section[:50]}...")  # Log first 50 characters
                        paragraph = doc.add_paragraph(extracted_section)
                        paragraph.runs[0].italic = True
                    else:
                        doc.add_paragraph("⚠️ No matching content found.", style="Normal")
            else:
                doc.add_paragraph(question["message_not_found"], style="Normal")

            doc.add_paragraph("")  # Spacing between questions

        doc.add_page_break()


    text_file_path = os.path.join(output_folder, f"{file_name}.txt")
    with open(text_file_path, "w", encoding="utf-8") as text_file:
        text_file.write(extracted_text)
    # Save output Word document
    os.makedirs(os.path.dirname(output_docx), exist_ok=True)
    doc.save(output_docx)

    # Create ZIP file including extracted text files
    final_zip_path = "output_files/processed_files.zip"
    with zipfile.ZipFile(final_zip_path, 'w') as zipf:
        #zipf.write(output_docx, os.path.basename(output_docx))  # Add processed docx
        for txt_file in extracted_text_files:
            print(f"📄 Adding extracted text file: {txt_file}")
            zipf.write(txt_file, os.path.basename(txt_file))  # Add extracted text files

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

# Example extracted text (replace this with a real sample from your document)
extracted_text = """enrjes Ae
Council

REPLIES TO STANDARD ENQUIRIES
OF LOCAL AUTHORITY 2007 Edition

Applicant Mr Kevin Bird
1200 Delta Business Park
Swindon
SN5 7XZ

Search Reference 1415 03258
NLIS Reference

Date 27Feb2015
Property 40 Gordon Road
Edmonton
London
Enfield
N9 OLU
Other Roads
ete
Additional 40 Gordon Road Edmonton N9 OLU
Properties 40A Gordon Road Edmonton N9 OLU

 refer to your Standard Enquiries relating to the above property These replies relate
to that property as shown on the location plan where supplied The replies are given
subject to the Notes to the Standard Enquiries

All correspondence relating to these answers should quote the official Search
Reference

Standard Enquiries of Local Authority

Search Reference 141503258

Enfield Council Civic Centre  Silver Street  Enfield JEN1 3XA  DX 90615 Enfield
landchargesenfieldgovuk  020 8379 1000

Page 1 of 9
1 PLANNING AND BUILDING REGULATIONS
11 Planning and Building Regulation Decisions and Pending Applications

11a A Planning Permission

None

11b A Listed Building consent

b None

11c A Conservation Area consent

c None

11d A Certificate of Lawfulness of existing use or development
d Reference 1500443CEU

Use of premises as 2 selfcontained flats

No Decision to date

11e A Certificate of Lawfulness of proposed use or development

e None

Copies of any of the planning documents listed above from 2006 onwards are available free of charge via
wwwenfieldgovuk For copies of any of the planning application documents listed above prior to 2006 are
available for a fee from licensingenfieldgovuk The reply shown in 11 ae does not cover other properties
in the vicinity of the property To obtain information regarding developments which may affect the property
please access the Planning Portal this can be found on the Enfield Council Website wwwenfieldgovuk

11f Building Regulation Approval
f None
11g A Building Regulation Completion Certificate and

g None

The Councils computerised records relating to building regulation information do not extend back before 1st
January 1999 and this reply only covers the period since that date

11h Competent Persons Scheme
any building regulations certificate or notice issued in respect of work carried out under a
competent person selfcertification scheme

h None

As from 1st April 2002 the installation of a replacement window rooflight or roof window or specified type of
glazed door must either have building regulation approval or be carried out and certified by a person who is
registered under the Fenestration SelfAssessment Scheme by the Glass and Glazing federation

The replies supplied in answer to questions 31  313 on form CON29R relate only to matters which are not
entered on the Local Land Charges Register Notices that have been withdrawn or quashed are also not
referred to

Search Reference 1415 03258

Enfield Council Civic Centre  Silver Street  Enfield EN1 3XA  DX 90615 Enfield
landchargesenfieldgovuk  020 8379 1000

Page 2 of 9
412

Unless otherwise indicated matters will be disclosed only if they apply directly to the property described in
Box B

Area means any area in which the property is located

References to the Council include any predecessor Council and also any Council Committee sub
committee or other body or person exercising powers delegated by the Council and their approval includes
their decision to proceed The replies given to certain enquiries cover knowledge and actions of both the
District Council and the County Council

References to the provisions of particular Acts of Parliament or Regulations include any provisions which
they have replaced and also include existing or future amendments or reenactments

The replies will be given in the belief that they are in accordance with information presently available to the
officers of the replying Council but none of the Councils or their officers accept legal responsibility for an
incorrect reply except for negligence Any liability for negligence will extend to the person who raised the
enquiries and the person on whose behalf they were raised It will also extend to any other person who has
knowledge personally or through an agent of the replies before the time when he purchases takes
tenancy of or lends money on the security of the property or if earlier the time when he becomes
contractually bound to do so

INFORMATION REGARDING LOCAL PLANS WILL FOLLOW
Planning Designations and Proposals

12 What designations of land use for the property or the area and what specific proposals for
the

property are contained in any exisiting or proposed development plan

None

The Entield Plan  Core Strategy was submitted to the Secretary of State on the 16th March 2010 and the
Council adopted the Core Strategy on the 10th November 2010 The Development Plan for the Local
Authority now comprises of i The Enfield Plan Core Strategy adopted November 2070 ii the saved
policies of the 1994 London Borough of Enfield Unitary Development Plan as updated November 2010 iii
The London Plan including alterations 2008

The Council is continuing to prepare more planning documents as part of the Local Development
Framework Further details of the document to be prepared as part of the LDF are set out in the Local
Development Scheme also available on the Council website wwwenfieldgovuk

if you wish to obtain further details on this matter please contact the Planning Policy Team on 020 8379
1000 or via email to planningpolicyenfieldgovuk

2 ROADS

Which of the roads footways and footpaths named in the application for this search via boxes B
and C) are

2(a) Highways maintainable at public expense

(a) Gordon Road is publicly maintained

Page 3 of 9
2d To be adopted by a local authority without reclaiming the cost from the frontagers

d Not applicable

If a road footpath or footway is not a highway there may be no right to use it The Council cannot express
an opinion without seeing the title plan of the property and carrying out an inspection whether or not any
existing or proposed highway directly abuts the boundary of the property

An affirmative answer does not imply that the public highway directly abuts the boundary of the property If a
road footpath or footway is not a highway there may be no right to use it

3 OTHER MATTERS

THE REPLIES TO ENQUIRIES 31 TO 313 RELATE ONLY TO MATTERS WHICH ARE NOT
ENTERED ON THE LOCAL LAND CHARGES REGISTER

31 Land required for Public Purposes
31 Land required for Public Purposes
None
32 Land to be acquired for Road Works
Is the property included in land to be acquired for road works
None
Relevant documents can be obtained from trafficntransportsupportenfieldgovuk
if a road footpath or footway is not a highway there may be no right to use it The Council cannot express an opinion
without seeing the title plan of the property and carrying out an inspection whether or not any existing or proposed
highway directly abuts the boundary of the property
33 Drainage Agreements and Consents
Do either of the following exist in relation to the property
33a An agreement to drain buildings in combination into an existing sewer by means of a
private sewer or
a No
33b An agreement or consent for i a building or ii extention to a building on the property to
be built over or in the vicinity of a drain sewer or disposal main
b No
Copy Combined Drainage Orders can be obtained for a fee from landchargesenfieldgovuk

Page 4 of 9
35

36

34b The centre line of a proposed alteration or improvement to an existing road involving
construction of a subway underpass flyover footbridge elevated road or dual carriageway

b None

34c The outer limits of construction works fora proposed alteration or improvement to an
existing road involving i construction of a roundabout other than a mini roundabout or ii
widening by construction of one or more addtional traffic lanes

c None

34d The outer limits of i construction of a new road to be built by a local authority ii an
approved alteration or improvement to an existing road involving construction of a subway
underpass flyover footbridge elevated road or dual carriageway or iii construction of a
roundabout other than a mini roundabout or widening by construction of one or more additional
traffic lanes

d None

34e The centre line of the proposed route of a new road under proposals published for public
consultation or

e None

34f The outer limits of i construction of a proposed alteration or improvement to an existing
road involving construction of a subway underpass flyover footbridge elevated road or dual
carriageway or ji construction of a roundabout other than a mini roundabout or iii widening
by construction of one or more additional traffic lanes under proposals published for public
consultation

f None
Relevant documents can be obtained from trafficntransportsupportenfieldgovuk

Page 5 of 9
37

36c One way driving

c None

36d Prohibition of driving

d None

36e Pedestrianisation

e None

36f Vehicle width or weight restriction

f None

36g Trafic calming works including road humps
g None

36h Residents parking controls

h None

37e highways or

e None

37f Public health

f None

Relevant documents can be obtained from envirocrimeenfieldgovuk

Infringement of Building Regulations

Has a local authority authorised in relation to the property any proceedings for the contravention
of any provision contained in Building Regulations

None
Relevant documents can be obtained from buildingcontrolenfieldgovuk

Notices Orders Directions and Proceedings under Planning Acts

Do any of the following subsist in relation to the property or has a local authority decided to
issue serve make or commence any of the following

39a An enforcement notice

a None

39b A stop notice

b None

39c A listed building enforcement notice
c None

39d A breach of condition notice

d None

39e A planning contravention notice

e None

39f Another notice relating to breach of planning control
f None

39g A listed building repairs notice

Search Reference 141503258

Enfield Council Civic Centre  Silver Street  Enfield EN1 3XA  DX 90615 Enfield
landchargesenfieldgovuk  020 8379 1000

39i A building preservation notice

i None

39j A direction restricting permitted development

j None

39k An order revoking or modifying planning permission

k None

391 An order requiring discontinuance of use or alteration or removal of building or works
I None

39m A tree preservation order or

m None

39n Proceedings to enforce a planning agreement or planning contribution

n None
Relevant documents can be obtained from envirocrimeenfieldgovuk

310 Conservation Area
Do the following apply in relation to the property
310a The making of the area a Conservation Area before 31 August 1974 or
a None
310b An unimplemented resolution to designate the area a Conservation Area

b None
Relevant documents can be obtained from planning policyenfieldgovuk
31



 Compulsory Purchase

Has any enforceable order or decision been made to compulsorily purchase or acquire the
property

Page 8 of 9
pollution of controlled waters might be caused on the property
312a A contaminated land notice
a None

312b In relation to a register maintained under section 78R of the Environmental Protection Act
1990 i a decision to make an entry or ii an entry or

b None

312c consultation with the owner or occupier of the property conducted under section 78G3
of the Environmental Protection Act 1990 before the service of a remediation notice
c None

A negative reply does not imply that the property or any adjoining or adjacent land is free from contamination or the risk
of it and the reply may not disclose steps taken by another Council in whose area adjacent or adjoinging land is situated

313 Radon Gas

Do records indicate that the property is in a Radon Affected Area as identified by the Health
Protection Agency

None

The replies will be given after the appropriate enquiries and in the belief that they are in
accordance with the information at present available to the Officers of the replying Councils but
on the distinct understanding that none of the Councils nor any Council Officer is legally
responsible for them except for negligence Any liability for negligence shall extend for the
benefit of not only the person by or for whom these enquiries are made but also a person being
a purchaser for the purpose of S103 of whom the Local Land Charges Act 1975 who or whose
agent had knowledge before the relevant time as defined in that section of the replies to those
enquiries

Kate Robertson

Assistant Director Customer Services and Information Finance Resources and Customer Services Department

Search Reference 141503258

Page 9 of 9
1401 eGeq

ZEsebSdd 49H OTL

Bally SUOZ ONUOD eyOWS

9S6L JV JI UBSID

uorsiaoid A10nes ayerudosdde
0 29Uesaj301 Bulpnjou aBseyo jo uoHdOSeq

PlayUy Jang JaAIIg aNUED
DIAID OUND pjeyuy sebseyD pue e007

aq Aew SUSWNDOP JUBAGIAl B1OYM BIE

young pjayug

Awouny uyeubuo

sabieyd snoauelaosiw y Hed

younog

a713I4NI

YdIBIS JO SJEIIJIPIOD CIO 0 ajnpeyog
sabieyd pue e907 Jo 19sIBay

SLZ022 918Q 8SZE0 SLpL aouaiajey Yeas 1071
ENFIELD

Councit

REGISTER OF LOCAL LAND CHARGES
OFFICIAL CERTIFICATE OF SEARCH

Search Reference 141503258
NLIS Reference
Date 27Feb2015
Applicant
Mr Kevin Bird
1200 Delta Business Park
Swindon
SN5 7XZ

Official Search required in all parts of the Register of Local Land Charges for subsisting registrations
against the land described and the plan submitted

Land
40 Gordon Road
Edmonton
London
Enfield
N9 OLU

It is hereby certified that the search requested above reveals the 1 registration described in the
Schedules hereto up to and including the date of this certificate

Enfield Council Civic Centre  Silver Street  Enfield IEN1 3XA  DX 90615 Enfield
landchargesenfieldgovuk  020 8379 1000

Page 1 of 1

Associated Notes Search Reference 141503258

ADDITIONAL INFORMATION

We would like to draw your attention to the following

Note For copies of Smoke Control Orders Combined Drainage Orders Tree Preservation Orders Section 106s and
Deeds of Dedication please note there is a 10 charge per document

This can be paid either in writing with an accompanying cheque payable to London Borough of Enfield  on your
covering letter quote the search address and search reference number or alternatively over the phone by debit  credit
card

For other documents please contact the relevant departments
Note Reference NO573758

Page 1 of 1
Search Reference 1415 03258

Enfield Council

Civic Centre
Silver Street
Enfield ENFIELD
Property Address 40 Gordon Road EN1 3XA Council
Edmonton
London DX 90615 Enfield
Enfield
N9 OLU landchargesenfieldgovuk
Date 26Feb2015 Scale 1 1250

Reproduced from the Ordnance Survey mapping with the permission of the Controller of Her Majestys Stationery Office  Crown copyright Unauthorised
reproduction infringes Crown copyright and may lead to prosecution or civil proceedings Enfield Council Ordinance Survey License 100019820
"""

# Regex pattern from the script
pattern = r"2\(a\)\s*(.*?)(?:\n|$).*?\(a\)\s*(.*?)\n"

message_template = "{extracted_text_1}. The main road ({extracted_text_2}) is a highway maintainable at public expense. A highway maintainable at public expense is a local highway. The local authority is responsible for maintaining the road, including repairs, resurfacing, and other works. It will be maintained according to the standards of the local authority and you will have access to it."

formatted_message = extract_matching_text(extracted_text, pattern, message_template)
print(f"Formatted message: {formatted_message}")
