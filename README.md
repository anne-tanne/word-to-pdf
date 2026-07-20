# Word to PDF Converter

A small local macOS app that batch-converts Word documents (`.doc`/`.docx`) to PDF
by driving Microsoft Word — the same result as using **File → Save As → PDF**, but
for many files at once.

## Features

- Add individual files (**Add Files…**) or every Word document in a folder (**Add Folder…**)
- Tick/untick each file to choose exactly what gets converted, or remove it from the list
- Save each PDF next to its Word document, or into one folder you pick
- Progress and per-file status while converting

## Requirements

- macOS
- Microsoft Word installed and signed in with an account that can edit/save
- Python 3.13 with Tk (only needed to run from source or rebuild the app)

## Run from source

```bash
python3 -m venv venv
source venv/bin/activate
python3 main.py
```

## Build a double-clickable `.app`

```bash
source venv/bin/activate
pip install py2app
python3 setup.py py2app
```

The bundled app is written to `dist/Word to PDF.app`.

## Notes

- The conversion opens each document in Word, exports it as PDF, and closes it again.
  Word does the rendering, so the PDF matches Word's own output exactly.
- If a document is already open in Word, close it first — an open document can stall
  the conversion.
- Files that come from email or downloads may show Word's **Enable Editing** banner the
  first time. Open the file in Word once, click **Enable Editing**, save, then convert.
