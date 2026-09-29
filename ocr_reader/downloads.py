from __future__ import annotations

import json
import re

import streamlit as st


def trigger_download(download_id: str) -> None:
    """Click the native download button once, after Streamlit enables it.

    Files stay in Streamlit's download handler; no document data or filenames
    are interpolated into JavaScript. The normal button remains available if
    the browser blocks automatic downloads.
    """
    if not re.fullmatch(r"[0-9a-f]{32}", download_id):
        raise ValueError("Invalid download ID")
    selector = f".st-key-layoutlens_download_{download_id} button"
    st.html(
        """
        <script>
        (() => {
            const id = DOWNLOAD_ID;
            const selector = DOWNLOAD_SELECTOR;
            const started = window.__layoutlensAutoDownloads ??= new Set();
            if (started.has(id)) return;

            const tryDownload = () => {
                if (started.has(id)) return true;
                const button = document.querySelector(selector);
                if (!button || button.disabled) return false;
                started.add(id);
                button.click();
                return true;
            };
            if (tryDownload()) return;

            const observer = new MutationObserver(() => {
                if (tryDownload()) observer.disconnect();
            });
            observer.observe(document.body, {
                childList: true,
                subtree: true,
                attributes: true,
                attributeFilter: ["disabled"],
            });
            setTimeout(() => observer.disconnect(), 30000);
        })();
        </script>
        """.replace("DOWNLOAD_ID", json.dumps(download_id)).replace(
            "DOWNLOAD_SELECTOR", json.dumps(selector)
        ),
        unsafe_allow_javascript=True,
    )
