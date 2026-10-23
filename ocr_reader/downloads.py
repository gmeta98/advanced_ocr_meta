"""Session-only browser downloads; no file data goes to another service."""

import base64
import json

import streamlit as st


def auto_download_word(filename: str, data: bytes, download_id: str) -> None:
    content = _download_html(filename, data, download_id)
    if hasattr(st, "iframe"):
        st.iframe(content, height=0, tab_index=-1)
    else:
        # Compatibility with the older Streamlit releases allowed by requirements.
        from streamlit.components.v1 import html
        html(content, height=0, scrolling=False, tab_index=-1)


def _download_html(filename: str, data: bytes, download_id: str) -> str:
    payload = json.dumps({
        "filename": filename,
        "data": base64.b64encode(data).decode("ascii"),
        "id": download_id,
    }).replace("<", "\\u003c")
    return """<!doctype html><html><body><script>
    (() => {
      const file = PAYLOAD;
      const key = "layoutlens-download:" + file.id;
      // Streamlit can recreate an iframe. Do not repeat the same download.
      try { if (sessionStorage.getItem(key)) return; } catch (_) {}
      const bytes = Uint8Array.from(atob(file.data), char => char.charCodeAt(0));
      const blob = new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = file.filename;
      document.body.appendChild(link);
      link.click();
      try { sessionStorage.setItem(key, "attempted"); } catch (_) {}
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 60000);
    })();
    </script></body></html>""".replace("PAYLOAD", payload)
