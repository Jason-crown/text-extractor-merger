# Text Extractor

Small Python tool to extract all text from given input of files (including .zip), folder and write in single txt-file. Desktop application with simple GUI as well as command line version, similar to [PDF Merger](../PDF_merger).

## Features

- Extracts content from archives like .zip, .tar, .tar.gz/.tgz, .tar.bz2, .tar.xz, .gz, .bz2, .xz and optionally .7z
- Processes nested archive layers up to 5 levels deep
- Unpacks no contents onto your hard drive; everything reads into RAM
- Pulls plain-text data out of any type of plain-text file (.txt, .md, .csv, .py, .json, etc.) and PDFs, MS Word docs (.docx), MS PowerPoint slideshows (.pptx) & MS Excel spreadsheets (.xlsx)
- Adds a label at the top of each extracted file (such as "===== notes.zip/week1/lecture.txt =====") or omits if you prefer
- Provides an inventory list for all files that were excluded and the reasons they were omitted (e.g. images, other binary files, empty files, scanned PDFs)
- Disregards common desktop clutter files (.macosx/, .ds_store & Thumbs.db)`

## Requirements

- Python 3.9 or newer
- Tkinter (included with the standard Python installers for Windows and macOS)
- Optional: [pypdf](https://pypi.org/project/pypdf/) for PDFs, [py7zr](https://pypi.org/project/py7zr/) for `.7z` archives

```
python -m pip install pypdf py7zr
```

Everything else uses the Python standard library.

## Usage

### Desktop window

Run the script with no arguments:

```
python text_extractor.py
```

1. Click **Add files…** to add files or archives, or **Add folder…** to add everything in a folder.
2. Change the order with **Move up** and **Move down**. Remove an item with **Remove** or the Delete key.
3. Untick **Label each file with a header** if you want only the raw text.
4. Click **Extract →** and choose where to save the result.

When it finishes, a message shows how many files were written and which were skipped.

### Command line

```
python text_extractor.py -o combined.txt lectures.zip notes.md project_folder/
```

| Option | Description |
| --- | --- |
| `inputs` | One or more files, folders or archives |
| `-o`, `--output` | Output file (default: `combined.txt`) |
| `--no-headers` | Don't write a header line before each file |
| `-q`, `--quiet` | Don't list skipped files |

## Output

The output is a UTF-8 text file. Files appear in the order you added them; the contents of a folder or archive appear in their stored order. PDFs are split with `--- Page N ---`, slides with `--- Slide N ---` and spreadsheets with `--- Sheet N ---`.

```
===== lectures.zip/week1/intro.txt =====

Welcome to the course...


===== lectures.zip/week1/slides.pdf =====

--- Page 1 ---
...
```

## Limitations

- RAR archives aren't supported. Convert them to `.zip` or `.7z` first.
- Password-protected archives and PDFs can't be opened.
- Scanned PDFs (images of text) have no text to extract and are skipped.
- Old Office formats (`.doc`, `.ppt`, `.xls`) aren't supported; save them as `.docx`, `.pptx` or `.xlsx`.
- Any single file over 100 MB inside an archive is skipped (a guard against "zip bombs"). Change `MAX_FILE_BYTES` at the top of the script to raise this.

## Project structure

```
extractor/
├── text_extractor.py   # Extraction logic, command-line interface and desktop window
├── README.md
└── LICENSE
```

The extraction logic lives in `extract_text()`, so you can also import it from other Python code:

```python
from text_extractor import extract_text

result = extract_text(["lectures.zip", "notes.md"], "combined.txt")
print(len(result.written), "files written,", len(result.skipped), "skipped")
```
