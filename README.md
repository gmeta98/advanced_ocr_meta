# LayoutLens OCR — Streamlit Cloud package

LayoutLens converts scanned and mixed-layout PDFs into translation-ready Word
documents with GPT-6 vision. It can also create an exact-look Word reference.

This folder is the clean deployment package. It contains no API key, local
environment file, uploaded PDF, generated Word document, test output, or cache.

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
offers **GPT-6 Sol** (the default) and **GPT-6 Luna** (the lower-cost option).
Sol uses medium reasoning; Luna uses low reasoning. Every OCR conversion uses the API key stored
in Streamlit and can create API charges.

The password screen is an additional app-level gate. Keeping the Streamlit app
private as well provides stronger access control and is still recommended.

## Automatic downloads

After a conversion produces a Word file, the app automatically starts a download.
A single Word output downloads directly. Multiple outputs, including an optional
layout JSON, or a multi-PDF batch download as one ZIP containing the available
results and usage reports. Individual download buttons remain available.

Automatic downloads run once per conversion. If your browser blocks the download,
use the download button. Changing options or downloading another result does not
automatically download the previous conversion again.

Command-line model selection uses `--model sol` (default) or `--model luna`.

## Tests

Install `requirements-dev.txt`, then run `python -m pytest -q` from this folder.
The tests use synthetic documents and mocked OCR responses; they make no paid API calls.
