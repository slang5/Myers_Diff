import re

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
    # Remove leading and trailing whitespace, and replace multiple spaces with a single space but preserve line breaks
    cleaned_text = ' '.join(txt.split())
    
    # Normalize line breaks to a single newline character
    cleaned_text = cleaned_text.replace(' \n', '\n').replace('\n\n ', '\n')

    # Remove line when there is only one character that is in the list of characters to ignore
    ignore_characters = ['-', '_', '*', '#', ' ', '.', ',', ';', ':', '!', '?', '"', "'"]
    cleaned_text = '\n'.join(line for line in cleaned_text.split('\n') if not (len(line) == 1 and line in ignore_characters))

    # Add new line after each sentence-ending punctuation (., !, ?) followed by a space and an uppercase letter (indicating the start of a new sentence)
    cleaned_text = re.sub(r'([.!?])\s+(?=[A-Z])', r'\1\n', cleaned_text)

    # Add a space when there is a lowercase letter or a digit followed by an uppercase letter (indicating the start of a new sentence without punctuation)
    cleaned_text = re.sub(r'([a-z0-9])(?=[A-Z])', r'\1 ', cleaned_text)

    # add new line after each bullet point ('-', '*', '•') that is between (letter, digit, or closing parenthesis) and a space
    cleaned_text = re.sub(r'(?<=[a-zA-Z0-9\)])\s*([-*•])\s+', r'\n\1 ', cleaned_text)

    # replace characters with accents with their non-accented counterparts
    cleaned_text = unidecode(cleaned_text)
    duration = time.time() - start_time

    console.print(f'[dark_orange]Cleaned text from {length_before} characters to {len(cleaned_text)} characters in {duration:.2f} seconds[/dark_orange]')

    return cleaned_text