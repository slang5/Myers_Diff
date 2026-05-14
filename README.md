# Myers Diff

A Python port of the Myers diff algorithm with two DOCX output modes:

- a line-by-line comparison view
- a reference-based inline annotation view

The repository keeps the core diff engine separate from presentation so the algorithm stays readable, testable, and easy to reuse.

## Main Idea

Myers diff finds the shortest edit script between two sequences. In this project, the sequences are lines of text, so the core algorithm works on line arrays rather than characters.

The project has two layers:

1. Core diff layer: compute `equal`, `delete`, and `insert` edits with Myers.
2. Presentation layer: turn those edits into human-readable output in `.docx` files.

The important design choice is that the algorithm computes the diff, while the renderers decide how to show it.

## What The Repository Produces

The repo generates two DOCX files for one comparison:

- `2LComp_...docx`: a stacked line-by-line comparison.
- `Annota_...docx`: a reference-based annotated document that inserts inline change markers around the changed text.

The main entry point is [Myers_main.py](Myers_main.py), which asks the user to pick two files, reads them, computes the diff, and writes both documents.

## How The Diff Works

The core implementation is in [Myers_diff.py](Myers_diff.py).

High level flow:

1. Split the input text into lines.
2. Run the Myers forward search.
3. Backtrack to recover the edit script.
4. Group raw edits into change blocks.
5. Optionally run a second token-level diff inside `replace` blocks.

That last step is not part of the classic Myers algorithm itself. It is a presentation aid so the DOCX output can show word-level changes inside a line.

## Files And Their Purpose

### Core algorithm

- [Myers_diff.py](Myers_diff.py): main Myers implementation, change-block grouping, token-level replace diff, and formatting helpers.
- [Myers_Class_Type.py](Myers_Class_Type.py): shared dataclasses and type aliases for lines, edits, change blocks, and algorithm state.

### Input preparation

- [Myers_source_file.py](Myers_source_file.py): reads PDF, DOCX, and TXT files, then stores extracted text into intermediate `.txt` files.
- [Myers_text_parsing.py](Myers_text_parsing.py): cleans and normalizes extracted text before diffing.

### DOCX output

- [Myers_docx_compared_on_two_lines.py](Myers_docx_compared_on_two_lines.py): the stacked line-by-line DOCX renderer.
- [Myers_docx_reference_annotations.py](Myers_docx_reference_annotations.py): the reference-based DOCX renderer with compact inline annotations.

### Entry point and notes

- [Myers_main.py](Myers_main.py): interactive UI entry point that asks the user for two files and writes both DOCX outputs.
- Algorithm notes and Python-oriented explanation of the theoretical Myers implementation.

## How To Run

Install dependencies first:

```bash
pip install -r requirements.txt
```

Then run the main script:

```bash
python Myers_main.py
```

The script opens file pickers for the reference and modified files, then writes the output documents under the `Compare/` folder.

## Why There Are Two DOCX Formats

The two renderers solve different review tasks.

The line-by-line view is easier to scan when you want to understand the global structure of the diff.

The reference-based annotated view is better when you want to stay close to the original text and see only the changed spans.

## What Is Canonical Myers And What Is A Deliberate Choice

Some parts of this repository are the Myers algorithm itself. Other parts are deliberate choices made to improve readability or file output.

### Core algorithm parts that are close to canonical Myers

- `find_shortest_edit()` in [Myers_diff.py](Myers_diff.py) is the actual Myers forward search.
- `backtrack()` in [Myers_diff.py](Myers_diff.py) reconstructs the edit script from the saved trace.
- `compute_diff()` and `diff_lines()` are thin wrappers around the search and backtracking.

### Derived views and presentation choices

These are useful, but they are not the core Myers algorithm:

- `build_change_blocks()` groups raw edits into `equal`, `delete`, `insert`, and `replace` blocks.
- `diff_replace_block()` runs a second token-level diff inside `replace` blocks.
- `build_replace_token_diffs()` prepares token diffs for rendering.
- The DOCX writers use colors, fills, spacing, and inline markers to make the output readable.

### Input normalization choices

These also affect the final diff, because they happen before Myers runs:

- `Myers_text_parsing.py` collapses whitespace and normalizes text.
- `Myers_source_file.py` extracts text from PDF and DOCX before diffing.
- `split_lines()` uses `text.split("\n")` so trailing empty lines are preserved in the same way the current implementation expects.

## Places Where The Project Uses Heuristics

A few parts of the output are intentionally heuristic rather than mathematically required by Myers:

- semantic chunking in the reference-based DOCX writer
- bounded coalescing of same-type token edits
- field-aware splitting around punctuation and separators
- color and fill choices for readability
- turning a replace block into a human-friendly inline annotation

These choices are meant to help a reviewer, not to change the algorithmic result.

## Useful Implementation Notes

- Line numbers are 1-based.
- The diff engine operates on normalized line arrays, not raw character streams.
- `replace` is a derived block type, not a primitive Myers edit operation.
- Equal text in the annotated DOCX should remain plain context whenever possible.
- Whitespace is preserved in the token diff so the document can be reconstructed faithfully.

## Dependencies

The project uses the packages listed in [requirements.txt](requirements.txt):

- `Rich`
- `python-docx`
- `pypdf`
- `unidecode`
- `tk` / `tkinter` runtime support
