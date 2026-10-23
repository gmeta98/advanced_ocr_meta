# LayoutLens OCR — Streamlit Cloud package

LayoutLens converts scanned and mixed-layout PDFs into translation-ready Word
documents with GPT-5.6 vision. It can also create an exact-look Word reference.

This folder is the clean deployment package. It contains no API key, local
environment file, uploaded PDF, generated Word document, test output, or cache.

## Changes in version 0.5.1

Editable Word downloads now use the suffix `-word.docx` instead of
`-translation-ready.docx`. Exact-look and JSON filenames are unchanged.

The interface now follows Streamlit's active light or dark theme for its page,
text and controls. Primary buttons and selected options use an explicit
dark-green background with white text. Fixed light backgrounds and conflicting
text colors have been removed. File cards, help text, conversion status and usage
metrics remain readable in both themes.

**Automatically download completed Word files** is on by default. Each PDF's
Word output starts downloading when that PDF finishes, while the remaining batch
continues. Browsers may ask you to allow multiple downloads. Manual download
buttons and the batch ZIP remain available. Turning the option off keeps downloads
manual. Refreshing or changing unrelated controls does not restart downloads.

JSON downloads are off by default. If needed for troubleshooting, open
**Advanced downloads** and enable **Include diagnostic JSON files**. Otherwise,
the app does not export OCR layout JSON or offer API usage JSON downloads, and
the batch ZIP contains only the selected Word outputs. Usage totals still appear
in the interface; their temporary internal report is removed with the job files.

## Word layout improvements in version 0.5.0

The new **Compact translation** style uses Times New Roman 12 pt, single-spaced
paragraphs and modest headings. Choose **Source typography** to retain the
previous font-selection approach. These are editable reconstructions; translated
text can still change the final page count.

Photographs and maps at the top of a page now stay in the body at a readable size.
Adjacent image blocks are placed together when their extracted positions align.
The exporter no longer assumes that every top-of-page image is a small header logo.

Under **Signatures and stamps**, choose between cropped images (the default) and
**Editable text and signature labels**. The latter keeps transcribed wording and
uses bracketed labels where only a mark was detected. Photographs, maps and
barcodes remain images in both modes. The exporter suppresses an exact, overlapping
stamp-text duplicate when separate editable wording is available.

**Skip pages with no detected content** removes pages whose OCR result contains
no text, table or visual. It does not remove pages merely because they contain
little text. Disable it to retain every selected page in editable Word. The
exact-look reference always keeps every selected source page. OCR can miss faint
content, so compare important documents against the source.

Source page labels are now literal text, rather than Word fields that change when
opened. Numbered lists continue across source pages and preserve explicit jumps.
Addresses and other form blocks retain intentional line breaks. Section-break
paragraphs use minimal spacing to reduce otherwise empty overflow pages.

The extraction instructions also distinguish letterhead from running headers,
request tight image crops and preserve source list markers. These prompt changes
still need a live API comparison against your documents; local export tests do not
measure model transcription accuracy.

### Suggested settings for the supplied example

- Word style: **Compact translation**
- Signatures and stamps: **Editable text and signature labels**
- Skip pages with no detected content: enabled
- Automatically download completed Word files: enabled
- Include diagnostic JSON files: off (only enable for troubleshooting)
- Exact-look reference: optional, useful for checking the edited result

The separate three-page layout demonstration reuses the supplied Italian extract
with manually mapped page regions. It shows the intended formatting, preserves
existing OCR errors, and is not an end-to-end conversion or a translation.

### Local use and checks

Use Python 3.12. From this folder:

```sh
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Configure the API key and app password as described below before using the app.
To run the local export and pipeline regression tests (no API key or network calls):

```sh
python -m unittest discover -s tests -v
```

The command-line converter accepts the same export choices:

```sh
python -m ocr_reader.cli input.pdf --mode both --word-style compact --marks editable --json
```

Add `--keep-blank-pages` to retain empty OCR pages. Conversion metadata records
the selected Word options and the omitted page numbers. Layout JSON retains the
original OCR data without the presentation changes.

## Streamlit Community Cloud settings

- Repository: your uploaded GitHub repository
- Branch: `main`
- Main file path: `app.py`
- Python version: `3.12`

Open **Advanced settings** and add this to **Secrets**:

```toml
OPENAI_API_KEY = "your OpenAI API key"
APP_PASSWORD = "choose a long, unique site password"
```

Never save the real API key or password inside this folder or commit them to GitHub.

Visitors must enter `APP_PASSWORD` before any upload or conversion controls are
shown. Authentication lasts for the current browser session, and the app provides
a **Sign out** button. Five failed attempts trigger a 30-second pause for that
browser session.

The app uses the Responses API with original-detail image input. The model picker
offers GPT-5.6 Sol, Terra, and Luna. Every OCR conversion uses the API key stored
in Streamlit and can create API charges.

The password screen is an additional app-level gate. Keeping the Streamlit app
private as well provides stronger access control and is still recommended.
