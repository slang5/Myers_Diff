from rich.console import Console
from unidecode import unidecode
import time 

console = Console()

def parsing_clean_txt(txt:str):
    """
    Cleans the input text by removing extra whitespace and normalizing line breaks.
    """
    length_before = len(txt)
    start_time = time.time()
    # Remove leading and trailing whitespace, and replace multiple spaces with a single space
    cleaned_text = ' '.join(txt.split())
    
    # Normalize line breaks to a single newline character
    cleaned_text = cleaned_text.replace('\r\n', '\n').replace('\r', '\n')

    # Remove line when there is only one character that is in the list of characters to ignore
    ignore_characters = ['-', '_', '*', '#', ' ', '.', ',', ';', ':', '!', '?', '(', ')', '[', ']', '{', '}', '"', "'"]
    cleaned_text = '\n'.join(line for line in cleaned_text.split('\n') if not (len(line) == 1 and line in ignore_characters))

    # replace characters with accents with their non-accented counterparts
    cleaned_text = unidecode(cleaned_text)
    duration = time.time() - start_time

    console.print(f'[dark_orange]Cleaned text from {length_before} characters to {len(cleaned_text)} characters in {duration:.2f} seconds[/dark_orange]')

    return cleaned_text