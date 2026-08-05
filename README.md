# LayoutLens OCR — Streamlit Cloud package

LayoutLens converts scanned and mixed-layout PDFs into translation-ready Word
documents with GPT-5.6 vision. It can also create an exact-look Word reference.

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
offers GPT-5.6 Sol, Terra, and Luna. Every OCR conversion uses the API key stored
in Streamlit and can create API charges.

The password screen is an additional app-level gate. Keeping the Streamlit app
private as well provides stronger access control and is still recommended.
