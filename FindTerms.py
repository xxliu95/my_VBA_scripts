import os
import csv
import re
from pptx import Presentation
from docx import Document
import PyPDF2
import pandas as pd

def load_terms(terms_file):
    terms_df = pd.read_excel(terms_file)
    term_col = next((col for col in terms_df.columns if col.strip().lower() == 'term'), None)
    if not term_col:
        raise ValueError("No column named 'Term' found in the Excel file.")

    terms_dict = {term.strip().lower(): term.strip() for term in terms_df[term_col] if isinstance(term, str)}
    return terms_dict


def search_terms_in_text(terms_dict, text):
    matches = []
    for term_lower, term_original in terms_dict.items():
        pattern = r'\b' + re.escape(term_lower) + r'\b'
        if re.search(pattern, text, flags=re.IGNORECASE):
            matches.append(term_original)
    return matches



def read_pptx(file_path, terms):
    presentation = Presentation(file_path)
    results = []
    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text.strip()
                matches = search_terms_in_text(terms, text)
                for match in matches:
                    results.append([slide_index, match, text, ""])
    return results

def read_docx(file_path, terms):
    doc = Document(file_path)
    results = []
    for paragraph_index, paragraph in enumerate(doc.paragraphs, start=1):
        text = paragraph.text.strip()
        matches = search_terms_in_text(terms, text)
        for match in matches:
            results.append([paragraph_index, match, text, ""])
    return results

def read_pdf(file_path, terms):
    results = []
    with open(file_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page_index, page in enumerate(pdf_reader.pages, start=1):
            text = page.extract_text().strip()
            matches = search_terms_in_text(terms, text)
            for match in matches:
                results.append([page_index, match, text, ""])
    return results

def process_file(file_path, terms):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pptx":
        return read_pptx(file_path, terms)
    elif ext == ".docx":
        return read_docx(file_path, terms)
    elif ext == ".pdf":
        return read_pdf(file_path, terms)
    else:
        print(f"Unsupported file format: {file_path}")
        return []

def write_results_to_csv(results, output_file):
    with open(output_file, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
        writer.writerow(["Page/Slide", "Term", "Paragraph", "Comments"])
        for result in results:
            result[2] = result[2].replace("\n", " ")
            writer.writerow(result)


def main():
    terms_file = "/Users/xxliu95/Documents/my_VBA_scripts/terms.xlsx"
    input_folder = "/Users/xxliu95/Documents/SAP/Liuxin"
    output_folder = "/Users/xxliu95/Documents/my_VBA_scripts/Output"

    os.makedirs(output_folder, exist_ok=True)

    terms_dict = load_terms(terms_file)
    print(f"Loaded {len(terms_dict)} terms from dictionary.")

    allowed_extensions = {".docx", ".pptx", ".pdf"}
    for file_name in os.listdir(input_folder):
        file_path = os.path.join(input_folder, file_name)
        if os.path.isfile(file_path) and os.path.splitext(file_name)[1].lower() in allowed_extensions:
            print(f"Processing {file_name}...")
            results = process_file(file_path, terms_dict)

            base_name = os.path.splitext(file_name)[0]
            output_file = os.path.join(output_folder, f"Source Analysis_{base_name}.csv")

            write_results_to_csv(results, output_file)
            print(f"Output: {output_file} \n")


if __name__ == "__main__":
    main()
