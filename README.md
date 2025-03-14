# Property Purchase Report

To generate a Property Purchase Report, click on the `input_files` folder and then the `Add file` dropdown, select `Upload files` and drag in a zip of the legal docs. After the file has finishd uploading, scroll down and click `Commit changes`.

The workflow will then automatically unzip the legal files (pdfs, etc), process the files to extract the required information and generate a word document report that can be downloaded.

To download the report, click on `Actions` then on the top icon under `workflow runs`, scroll down and click on the `processed_word_document_and_zip` link under `Artifacts`. The uploaded zip will be included in the artifact.

## Local Authority Search

To check if a local authority search has been conducted, the code looks for the text `REPLIES TO STANDARD ENQUIRIES` in all the documents in the uploaded zip.

If it doesn't find the text it writes an appropriate message in the report and exits.

If the text is found it creates a different message in the report and processes the search document as follows.

The code processes questions that just require extracting text from the document, such as `Are there any existing Planning Permissions?`, which are added to the report if they exist.
It then processess questions that will always require text to be added to the report if the answer to the question is yes or no (such as `Are there Local Land charges?`).

