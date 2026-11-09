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

1. To add your documents (files), folders or archives click on **Add Files…**, for adding all of the content within one folder, select **Add Folder…**
2. The list can be reordered using buttons **Move Up**, **Move Down**. A single file may be removed by clicking on **Remove**, or pressing the delete button on your keyboard.
3. If you are working with a very large number of files that need to be separated into individual headers based on their contents, uncheck **Label Each File With a Header** so you have only the raw data.
4. When ready, click **Extract →** and select where the extracted document will be saved.

Once complete, a notification will appear that includes the total amount of files successfully created from the original source as well as which items did not create a usable document.

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

The output is a text file using UTF-8 encoding. The files will be listed in the same order that they were added to your project. A folder or zip archive will also have its contents presented as if they were stored in that way. Pages within pdf documents will be separated by `--- Page N ---`. Slides from pptx/pptm files will be separated by `--- Slide N ---`. Sheets in xlsx/xlsm/xltm files will be separated by `--- Sheet N ---`.

```
===== lectures.zip/week1/intro.txt =====

Welcome to the course...


===== lectures.zip/week1/slides.pdf =====

--- Page 1 ---
...
```

## Limitations

– RAR files cannot be processed. Please convert the RAR file to a .ZIP or .7Z file type prior to processing. 
– Password protected RAR files, PDFs etc. will not open. 
– When scanning PDFs (image of text), there is nothing for the software to pull from the image and therefore these files will be bypassed. 
– Legacy Microsoft Office formats (.doc / .xls / ppt) are not compatible with our system. In order to use the legacy Microsoft office formats please save your documents in the new format (Office 2007/2010 formats). 
– For security reasons we currently limit each individual file within a compressed package (i.e. zip, rar, tar etc.) to 100MB in size. This prevents users from uploading maliciously large packages that could consume all available disk space. The value for MAX_FILE_BYTES can be adjusted near the top of the script to increase this number if needed.

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
