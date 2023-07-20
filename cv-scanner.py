import PyPDF2
import os
import re
import magic
from docx import Document

def extract_text_from_pdf(file):
    with open(file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)

        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()

    return text

def extract_text_from_docx(file):
    doc = Document(file)

    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text

    return text

def search_words(text, words):
    results = {}
    for word in words:
        occurrences = re.findall(word, text, re.IGNORECASE)
        if occurrences:
            results[word] = len(occurrences)
    return results

if __name__ == "__main__":
    words_to_search = ["CELTA", "DELTA", "TEFL"]

    # Get all files in current directory
    files = os.listdir()

    # Filter for only PDF and Word files
    target_files = [file for file in files if os.path.isfile(file) and magic.from_file(file, mime=True) in ['application/pdf', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document']]

    for file in target_files:
        mime_type = magic.from_file(file, mime=True)
        if mime_type == 'application/pdf':
            text = extract_text_from_pdf(file)
        elif mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            text = extract_text_from_docx(file)
        else:
            continue

        occurrences = search_words(text, words_to_search)

        if occurrences:
            print(f"In file {file}:")
            for word, count in occurrences.items():
                print(f"  '{word}' found {count} time(s)")
