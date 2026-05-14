from Myers_docx_compared_on_two_lines import *
from Myers_docx_reference_annotations import *
from Myers_source_file import *
from pathlib import Path
from uuid import uuid4
from datetime import date
from rich.console import Console


if __name__ == "__main__":

    console = Console()

    # Keep the example small so the generated document is easy to inspect.
    console.print(f'[red][bold]{("Initializing the Myers Diff Algorithm").center(60, "-")}[/bold][/red]')
    path_base = Path(__file__).parent / "test_env"
    console.print(f"[blue]Base path for test files:[/blue] {path_base}")

    # open files from test_env folder
    reference_file = str(path_base / "Ref_doc" / "W_juin26_800.pdf")
    console.print(f"[blue]Reference file path:[/blue] {reference_file}")
    
    new_file = str(path_base / "Modified_doc" / 'W_juin27_500.pdf')
    console.print(f"[dark_orange]Modified file path:[/dark_orange] {new_file}")

    old_text = read_file(reference_file)
    new_text = read_file(new_file)

    old_name = reference_file.split("\\")[-1].split(".")[0]
    new_name = new_file.split("\\")[-1].split(".")[0]
    uuid_ref = uuid4()
    dateformat = 'Y%Y_M%m_D%d'
    current_date = date.today().strftime(dateformat)

    file_name_2LComp = f'Compare/{old_name}_vs_{new_name}/{current_date}/2LComp_{uuid_ref}.docx'
    file_name_Annota = f'Compare/{old_name}_vs_{new_name}/{current_date}/Annota_{uuid_ref}.docx'

    output_file = save_diff_docx_from_text(old_text, new_text, file_name_2LComp)
    output_file = save_reference_annotated_docx_from_text(old_text, new_text, file_name_Annota)
    console.print(f"[green]Saved 2LComp DOCX output to[/green] {output_file}")
    console.print(f"[green]Saved Annota DOCX output to[/green] {output_file}")

    console.print(f'[red][bold]{("Ending the Myers Diff Algorithm").center(60, "-")}[/bold][/red]')

