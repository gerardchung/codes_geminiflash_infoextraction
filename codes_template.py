
## pip install these in terminal
# pip install google-generativeai
# pip install python-docx
# pip install python-dotenv


import os
import csv
import google.generativeai as genai
from docx import Document
from tqdm import tqdm

# You need an API key from Google -> get it for free!
# Copy the key in by replace "YOUR_API_KEY_HERE"
api_key = "YOUR_API_KEY_HERE"
genai.configure(api_key=api_key)




def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    return " ".join([paragraph.text for paragraph in doc.paragraphs])


def generate_response(prompt, document_text):
    model = genai.GenerativeModel('gemini-1.5-flash')
    full_prompt = f"Document content: {document_text}\n\nPrompt: {prompt}\n\nResponse:"
    response = model.generate_content(full_prompt)
    return response.text


def process_documents(folder_path, prompt, output_csv):
    results = []
    doc_files = [f for f in os.listdir(folder_path) if f.endswith('.docx')]

    for doc_file in tqdm(doc_files, desc="Processing documents"):
        doc_path = os.path.join(folder_path, doc_file)
        document_text = extract_text_from_docx(doc_path)
        response = generate_response(prompt, document_text)
        results.append({"Document": doc_file, "Response": response})

    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Document', 'Response']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in results:
            writer.writerow(row)


# Usage
folder_path = "documents" # create a folder call documents and put your words docs inside
prompt = ''' Your prompt here. It can be multi-line as long you keep your prompt between these triple quotes '''

output_csv = "output_results.csv"

process_documents(folder_path, prompt, output_csv)
