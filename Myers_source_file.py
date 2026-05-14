from pypdf import PdfReader
from docx import Document
import os
from uuid import uuid4
from datetime import date
from rich.console import Console

console = Console()

def generate_unique_filename(base_name, extension, uuid_ref):
    """
    Generates a unique filename using the base name, extension, and a UUID reference.
        - The filename is structured as: {current_date}/{base_name}_{uuid_ref}.{extension}
    """
    dateformat = 'Y%Y_M%m_D%d'
    current_date = date.today().strftime(dateformat)
    path_end = f'Conversion/{current_date}/{base_name}_{uuid_ref}.{extension}'
    return path_end


def read_pdf(file_path, uuid_ref):
    """
    Reads a PDF file and extracts its text content.
    """
    reader = PdfReader(file_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text()

    console.print(f"[blue]Extracted text from PDF file:[/blue] {file_path}")

    name_file = file_path.split("\\")[-1].split(".")[0]
    base_name = f'PDF/{name_file}'
    store_to_txt(text, generate_unique_filename(base_name, "txt", uuid_ref))

    return text

def read_docx(file_path, uuid_ref):
    """
    Reads a DOCX file and extracts its text content.
    """
    doc = Document(file_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"

    console.print(f"[blue]Extracted text from DOCX file:[/blue] {file_path}")

    name_file = file_path.split("\\")[-1].split(".")[0]
    base_name = f'DOCX/{name_file}'
    store_to_txt(text, generate_unique_filename(base_name, "txt", uuid_ref))

    return text

def read_txt(file_path, uuid_ref):
    """
    Reads a TXT file and extracts its text content.
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        console.print(f"[blue]Reading text from[/blue] {file_path}")
        text = f.read()
    
    name_file = file_path.split("\\")[-1].split(".")[0]
    base_name = f'TXT/{name_file}'
    store_to_txt(text, generate_unique_filename(base_name, "txt", uuid_ref))

    return text

def read_file(file_path):
    """
    Reads a file and extracts its text content based on the file type.
        - Supported file types are PDF, DOCX, and TXT.
    """
    uuid_ref = uuid4()
    if file_path.endswith('.pdf'):
        return read_pdf(file_path, uuid_ref)
    elif file_path.endswith('.docx'):
        return read_docx(file_path, uuid_ref)
    elif file_path.endswith('.txt'):
        return read_txt(file_path, uuid_ref)
    else:
        raise ValueError("Unsupported file type. Only .pdf, .docx, and .txt are supported.")
    
def store_to_txt(text, output_path):
    """
    Stores text content to a .txt file.
    """
    # check if output_path file exists, if not create it
    if not os.path.exists(output_path):
        console.print(f"[green]Creating new file at[/green] {output_path}")
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Write the text content to the specified output path.
    console.print(f"[green]Storing text to[/green] {output_path}")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(text)
    
