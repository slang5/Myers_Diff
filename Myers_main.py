from Myers_docx_compared_on_two_lines import *
from Myers_docx_reference_annotations import *
from Myers_source_file import *
from uuid import uuid4
from datetime import date
from rich.console import Console
from tkinter import filedialog

if __name__ == "__main__":

    console = Console()

    # Keep the example small so the generated document is easy to inspect.
    console.print(f'[red][bold]{("Initializing the Myers Diff Algorithm").center(60, "-")}[/bold][/red]')

    # open files for comparison based on user selection
    reference_file = filedialog.askopenfilename(title="Select the reference PDF/TXT/DOCX file", filetypes=[("PDF files", "*.pdf"), ("TXT files", "*.txt"), ("DOCX files", "*.docx")])
    reference_file = str(reference_file).replace("/", "\\")
    console.print(f"[blue]Reference file path:[/blue] {reference_file}")
    
    new_file = filedialog.askopenfilename(title="Select the modified PDF/TXT/DOCX file", filetypes=[("PDF files", "*.pdf"), ("TXT files", "*.txt"), ("DOCX files", "*.docx")])
    new_file = str(new_file).replace("/", "\\")
    console.print(f"[dark_orange]Modified file path:[/dark_orange] {new_file}")

    reference_text = read_file(reference_file)
    new_text = read_file(new_file)

    # process for storing the diff output in docx format
    reference_name = reference_file.split("\\")[-1].split(".")[0]
    new_name = new_file.split("\\")[-1].split(".")[0]
    uuid_ref = uuid4()
    dateformat = 'Y%Y_M%m_D%d'
    current_date = date.today().strftime(dateformat)

    file_name_2LComp = f'Compare/{reference_name}_vs_{new_name}/{current_date}/2LComp_{uuid_ref}.docx'
    file_name_Annota = f'Compare/{reference_name}_vs_{new_name}/{current_date}/Annota_{uuid_ref}.docx'

    # diffs are saved as docx files and with 2 formats: 2LComp for line-by-line comparison and Annota for inline annotations
    output_file = save_diff_docx_from_text(reference_text, new_text, file_name_2LComp)
    output_file = save_reference_annotated_docx_from_text(reference_text, new_text, file_name_Annota)
    console.print(f"[green]Saved 2LComp DOCX output to[/green] {output_file}")
    console.print(f"[green]Saved Annota DOCX output to[/green] {output_file}")

    console.print(f'[red][bold]{("Ending the Myers Diff Algorithm").center(60, "-")}[/bold][/red]')
