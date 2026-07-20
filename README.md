# Word to PDF Converter

A small local macOS app that batch-converts Word documents (`.doc`/`.docx`) to PDF
by driving Microsoft Word, the same result as using **File → Save As → PDF**, but
for many files at once.

## Features

- Add individual files (**Add Files…**) or every Word document in a folder (**Add Folder…**),
  or drag files and folders straight from Finder onto the window
- Tick/untick each file to choose exactly what gets converted, or remove it from the list
- Batch-name the PDFs: keep the original names, add text before/after each name, or
  rename them all to a base name plus a sequential number (with zero-padding for clean
  sorting), shown in a live preview
- Save each PDF next to its Word document, or into one folder you pick
- Progress and per-file status while converting
- Multi-language interface (English and German included), auto-selected from the
  macOS system language
- Resizable, scrollable window: the layout adapts to the window width and scrolls
  when short, while the Convert button stays pinned and always reachable

## Languages

All interface text lives in JSON translation files under `locales/`, one per
language (`en.json`, `de.json`). The app picks a language automatically from the
macOS UI language and falls back to English.

- **Add a language:** copy `locales/en.json` to e.g. `locales/fr.json`, translate the
  values, and add its code to `SUPPORTED_LANGUAGES` in `main.py`.
- **Force a language** (for testing): set `WORDTOPDF_LANG=de` (or `en`) in the
  environment before launching.

## Requirements

- macOS
- Microsoft Word installed and signed in with an account that can edit/save
- Python 3.13 with Tk (only needed to run from source or rebuild the app)

## Run from source

```bash
python3 -m venv venv
source venv/bin/activate
pip install tkinterdnd2   # for drag & drop (optional; the app runs without it)
python3 main.py
```

## Build a double-clickable `.app`

```bash
source venv/bin/activate
pip install py2app tkinterdnd2
python3 setup.py py2app
```

The bundled app is written to `dist/Word to PDF.app`.

## Notes

- The conversion opens each document in Word, exports it as PDF, and closes it again.
  Word does the rendering, so the PDF matches Word's own output exactly.
- If a document is already open in Word, close it first; an open document can stall
  the conversion.
- Files that come from email or downloads may show Word's **Enable Editing** banner the
  first time. Open the file in Word once, click **Enable Editing**, save, then convert.
