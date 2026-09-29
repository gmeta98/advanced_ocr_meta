from __future__ import annotations

import hashlib
import hmac
import io
import json
import shutil
import time
import uuid
import zipfile
from dataclasses import dataclass, replace
from pathlib import Path

import streamlit as st

from .config import (
    DEFAULT_MODEL_PRESET,
    MODEL_PRESETS,
    PROJECT_ROOT,
    Settings,
    apply_quality_preset,
    load_environment,
)
from .pdf import inspect_pdf_bytes
from .pipeline import PartialConversionError, convert_pdf
from .downloads import trigger_download


JOBS_DIR = PROJECT_ROOT / ".ocr_jobs"
MAX_UPLOAD_BYTES = 100 * 1024 * 1024
MAX_BATCH_BYTES = 500 * 1024 * 1024
MAX_FILES_PER_BATCH = 20
MAX_PAGES_PER_CONVERSION = 150
JOB_RETENTION_SECONDS = 24 * 60 * 60


@dataclass(frozen=True, slots=True)
class UploadDocument:
    name: str
    data: bytes
    page_count: int
    page_numbers: tuple[int, ...]


def run() -> None:
    st.set_page_config(
        page_title="LayoutLens OCR",
        page_icon="◫",
        layout="centered",
        initial_sidebar_state="collapsed",
    )
    _styles()
    if not _password_gate():
        return
    load_environment()
    if not st.session_state.get("job_cleanup_done"):
        _cleanup_stale_jobs()
        st.session_state["job_cleanup_done"] = True

    st.markdown('<div class="eyebrow">GPT-6 document vision</div>', unsafe_allow_html=True)
    st.markdown(
        '<h1 class="hero-title" style="color:#17251f !important">PDF to Word, without losing the page</h1>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Upload one or more PDFs and get editable Word reconstructions, exact-look copies, or both."
    )
    _feature_row()

    uploaded_files = st.file_uploader(
        "Choose PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
    if not uploaded_files:
        st.markdown(
            '<div class="empty-hint">Drop one or more scanned forms, certificates, reports, or mixed-layout PDFs here.</div>',
            unsafe_allow_html=True,
        )
        _privacy_note()
        return

    if len(uploaded_files) > MAX_FILES_PER_BATCH:
        st.error(f"Choose at most {MAX_FILES_PER_BATCH} PDFs in one batch.")
        return

    digest = hashlib.sha256()
    documents: list[UploadDocument] = []
    total_bytes = 0
    for uploaded in uploaded_files:
        data = uploaded.getvalue()
        total_bytes += len(data)
        digest.update(uploaded.name.encode("utf-8", errors="replace"))
        digest.update(b"\0")
        digest.update(data)
        if len(data) > MAX_UPLOAD_BYTES:
            st.error(
                f"{uploaded.name} is larger than 100 MB. Split or compress it, then try again."
            )
            return
        try:
            info = inspect_pdf_bytes(data)
        except Exception as exc:
            st.error(f"{uploaded.name}: {exc}")
            return
        documents.append(
            UploadDocument(
                name=uploaded.name,
                data=data,
                page_count=info.page_count,
                page_numbers=tuple(range(1, info.page_count + 1)),
            )
        )

    if total_bytes > MAX_BATCH_BYTES:
        st.error("This batch is larger than 500 MB. Split it into two smaller batches.")
        return

    upload_digest = digest.hexdigest()
    if st.session_state.get("upload_digest") != upload_digest:
        st.session_state["upload_digest"] = upload_digest
        st.session_state.pop("downloads", None)
        st.session_state.pop("api_usage_report", None)
        st.session_state.pop("auto_download_pending", None)

    for document in documents:
        st.markdown(
            f'<div class="file-card"><span>{_escape(document.name)}</span>'
            f'<span>{document.page_count} page{"s" if document.page_count != 1 else ""} · '
            f'{len(document.data) / 1024 / 1024:.1f} MB</span></div>',
            unsafe_allow_html=True,
        )

    if len(documents) == 1:
        document = documents[0]
        left, right = st.columns(2)
        with left:
            first_page = st.number_input("First page", 1, document.page_count, 1)
        with right:
            last_page = st.number_input(
                "Last page",
                1,
                document.page_count,
                document.page_count,
            )
        if first_page > last_page:
            st.warning("The first page must come before the last page.")
            return
        documents[0] = replace(
            document,
            page_numbers=tuple(range(int(first_page), int(last_page) + 1)),
        )
    else:
        total_pages = sum(document.page_count for document in documents)
        st.info(
            f"Batch mode will process every page in all {len(documents)} PDFs "
            f"({total_pages} pages total). Upload one PDF by itself to choose a page range."
        )

    st.subheader("Word output")
    editable = st.checkbox(
        "Translation-ready Word",
        value=True,
        help="Natural paragraphs, headings, lists, tables, headers, and footers that reflow while you translate.",
    )
    faithful = st.checkbox(
        "Exact-look reference",
        value=False,
        help="Optional image-based Word copy for visual comparison. It stays local and is not editable.",
    )
    if not editable and not faithful:
        st.info("Choose at least one Word output.")

    quality = st.segmented_control(
        "Quality",
        options=["Balanced", "Maximum"],
        default="Balanced",
        help="Maximum uses a higher-resolution scan and a second OCR verification pass. It is slower and uses more API tokens.",
    )
    reasoning_level = st.segmented_control(
        "OCR model",
        options=list(MODEL_PRESETS),
        default=DEFAULT_MODEL_PRESET,
        help=(
            "GPT-6 Sol is the default for detailed OCR. "
            "GPT-6 Luna is the lower-cost option for higher-volume conversion."
        ),
    )
    include_native_text = st.checkbox(
        "Use embedded PDF text as an OCR aid",
        value=True,
        help="Improves born-digital PDFs. Turn it off if the PDF contains an incorrect hidden OCR layer.",
    )
    include_json = st.checkbox(
        "Include OCR layout data (JSON)",
        value=False,
        help=(
            "Optional machine-readable coordinates and styles for diagnostics. "
            "It is not needed for translation and adds no model call when "
            "Translation-ready Word is already selected."
        ),
    )
    option_fingerprint = (
        upload_digest,
        tuple((document.name, document.page_numbers) for document in documents),
        editable,
        faithful,
        quality or "Balanced",
        reasoning_level or DEFAULT_MODEL_PRESET,
        include_native_text,
        include_json,
    )
    if st.session_state.get("conversion_options") != option_fingerprint:
        st.session_state["conversion_options"] = option_fingerprint
        st.session_state.pop("downloads", None)
        st.session_state.pop("api_usage_report", None)
        st.session_state.pop("auto_download_pending", None)
    _privacy_note()

    selected_page_count = sum(len(document.page_numbers) for document in documents)
    too_many_pages = selected_page_count > MAX_PAGES_PER_CONVERSION
    if too_many_pages:
        st.warning(
            f"A batch can contain at most {MAX_PAGES_PER_CONVERSION} pages total. "
            f"This selection contains {selected_page_count}; split it into smaller batches."
        )

    button_label = (
        "Convert to Word"
        if len(documents) == 1
        else f"Convert {len(documents)} PDFs to Word"
    )
    if st.button(
        button_label,
        type="primary",
        use_container_width=True,
        disabled=not (editable or faithful) or too_many_pages,
    ):
        st.session_state.pop("downloads", None)
        st.session_state.pop("api_usage_report", None)
        st.session_state.pop("auto_download_pending", None)
        _run_batch_conversion(
            documents=documents,
            editable=editable,
            faithful=faithful,
            quality=quality or "Balanced",
            reasoning_level=reasoning_level or DEFAULT_MODEL_PRESET,
            include_native_text=include_native_text,
            include_json=include_json,
        )

    _usage_summary()
    _downloads()


def _password_gate() -> bool:
    expected_password = _site_password()
    if not expected_password:
        st.error("This site is locked, but its password has not been configured.")
        st.caption(
            'Add APP_PASSWORD = "your password" to the Streamlit Secrets box, '
            "then reboot the app."
        )
        return False

    if st.session_state.get("layoutlens_authenticated"):
        _, sign_out_column = st.columns([5, 1])
        with sign_out_column:
            if st.button("Sign out", use_container_width=True):
                st.session_state["layoutlens_authenticated"] = False
                st.session_state.pop("layoutlens_failed_logins", None)
                st.session_state.pop("layoutlens_locked_until", None)
                st.rerun()
        return True

    st.markdown('<div class="eyebrow">Private access</div>', unsafe_allow_html=True)
    st.markdown(
        '<h1 class="hero-title" style="color:#17251f !important">LayoutLens OCR</h1>',
        unsafe_allow_html=True,
    )
    st.caption("Enter the site password to open the PDF-to-Word workspace.")

    locked_until = float(st.session_state.get("layoutlens_locked_until", 0.0))
    remaining = max(0, int(locked_until - time.time()))
    if remaining:
        st.error(f"Too many incorrect attempts. Try again in {remaining + 1} seconds.")
        return False

    with st.form("layoutlens_password_form", clear_on_submit=True):
        entered_password = st.text_input(
            "Site password",
            type="password",
            placeholder="Enter password",
        )
        submitted = st.form_submit_button(
            "Open LayoutLens",
            type="primary",
            use_container_width=True,
        )

    if submitted:
        entered_digest = hashlib.sha256(entered_password.encode("utf-8")).digest()
        expected_digest = hashlib.sha256(expected_password.encode("utf-8")).digest()
        if hmac.compare_digest(entered_digest, expected_digest):
            st.session_state["layoutlens_authenticated"] = True
            st.session_state.pop("layoutlens_failed_logins", None)
            st.session_state.pop("layoutlens_locked_until", None)
            st.rerun()

        failures = int(st.session_state.get("layoutlens_failed_logins", 0)) + 1
        st.session_state["layoutlens_failed_logins"] = failures
        if failures >= 5:
            st.session_state["layoutlens_failed_logins"] = 0
            st.session_state["layoutlens_locked_until"] = time.time() + 30
            st.error("Too many incorrect attempts. Try again in 30 seconds.")
        else:
            st.error("Incorrect password.")
    return False


def _site_password() -> str:
    try:
        from .config import _get_setting

        return (_get_setting("APP_PASSWORD", "") or "").strip()
    except Exception:
        return ""


def _run_batch_conversion(
    *,
    documents: list[UploadDocument],
    editable: bool,
    faithful: bool,
    quality: str,
    reasoning_level: str,
    include_native_text: bool,
    include_json: bool,
) -> None:
    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    JOBS_DIR.mkdir(parents=True, exist_ok=True, mode=0o700)
    JOBS_DIR.chmod(0o700)
    job_dir.mkdir(exist_ok=True, mode=0o700)
    progress_bar = st.progress(0.0)
    status = st.empty()

    try:
        mode = "both" if editable and faithful else "editable" if editable else "faithful"
        needs_ocr = editable or include_json
        settings = Settings.from_environment(require_api_key=needs_ocr)
        preset = MODEL_PRESETS.get(
            reasoning_level,
            MODEL_PRESETS[DEFAULT_MODEL_PRESET],
        )
        settings = replace(
            settings,
            model=preset.model,
            reasoning_effort=preset.reasoning_effort,
            include_native_text=include_native_text,
        )
        settings = apply_quality_preset(settings, quality)
        displayed_progress = 0.0
        downloads: list[tuple[str, str, bytes]] = []
        usage_reports: list[tuple[str, dict]] = []
        failures: list[tuple[str, str]] = []
        files_with_word_output = 0
        file_count = len(documents)

        for index, document in enumerate(documents, start=1):
            document_dir = job_dir / f"document-{index:03d}"
            output_dir = document_dir / "output"
            document_dir.mkdir(exist_ok=True, mode=0o700)
            source_path = document_dir / _safe_filename(document.name)
            source_path.write_bytes(document.data)
            source_path.chmod(0o600)

            def progress(message: str, current: int, total: int) -> None:
                nonlocal displayed_progress
                status.text(f"File {index}/{file_count}: {document.name} — {message}")
                phase_value = _phase_progress(message, current, total)
                overall_value = ((index - 1) + phase_value) / file_count
                displayed_progress = max(displayed_progress, overall_value)
                progress_bar.progress(displayed_progress)

            try:
                result = convert_pdf(
                    source_path,
                    output_dir,
                    page_numbers=list(document.page_numbers),
                    settings=settings,
                    mode=mode,
                    include_layout_json=include_json,
                    progress=progress,
                )
            except PartialConversionError as exc:
                result = exc.result
                failures.append(
                    (
                        document.name,
                        _friendly_error(exc.__cause__ or exc),
                    )
                )
            except Exception as exc:
                failures.append((document.name, _friendly_error(exc)))
                displayed_progress = max(displayed_progress, index / file_count)
                progress_bar.progress(displayed_progress)
                continue

            if result.editable_docx or result.faithful_docx:
                files_with_word_output += 1
            is_batch = file_count > 1
            downloads.extend(
                _load_downloads(
                    result,
                    source_name=document.name if is_batch else None,
                    filename_prefix=f"{index:02d}-" if is_batch else "",
                )
            )
            usage = _load_usage_report(result)
            if usage is not None:
                usage_reports.append((document.name, usage))
            displayed_progress = max(displayed_progress, index / file_count)
            progress_bar.progress(displayed_progress)

        combined_usage = _combine_usage_reports(usage_reports)
        if combined_usage is not None:
            st.session_state["api_usage_report"] = combined_usage
            if file_count > 1:
                downloads.append(
                    (
                        "Download combined API usage report (JSON)",
                        "layoutlens-batch-api-usage.json",
                        json.dumps(
                            combined_usage,
                            ensure_ascii=False,
                            indent=2,
                        ).encode("utf-8"),
                    )
                )
        else:
            st.session_state.pop("api_usage_report", None)

        output_count = sum(
            filename.endswith((".docx", "-layout.json"))
            for _, filename, _ in downloads
        )
        if downloads and (file_count > 1 or output_count > 1):
            downloads.insert(
                0,
                (
                    "Download all results (ZIP)",
                    (
                        "layoutlens-batch-results.zip"
                        if file_count > 1
                        else f"{Path(_safe_filename(documents[0].name)).stem}-results.zip"
                    ),
                    _build_download_zip(downloads),
                ),
            )
        if downloads:
            st.session_state["downloads"] = downloads
            st.session_state["download_id"] = job_id
            if files_with_word_output:
                st.session_state["auto_download_pending"] = job_id
        else:
            st.session_state.pop("downloads", None)

        progress_bar.progress(1.0)
        if failures:
            if file_count == 1:
                if files_with_word_output:
                    status.warning(
                        "OCR could not finish, but the available Word output is ready below."
                    )
                else:
                    status.error(
                        "Conversion stopped before a requested Word file was ready."
                    )
            else:
                status.warning(
                    f"Batch finished: {files_with_word_output} of {file_count} PDFs "
                    "produced a Word file."
                )
            for name, detail in failures:
                st.warning(f"{name}: {detail}")
        else:
            noun = "PDF" if file_count == 1 else "PDFs"
            status.success(f"Finished {file_count} {noun}.")
    except Exception as exc:
        st.session_state.pop("downloads", None)
        st.session_state.pop("api_usage_report", None)
        st.session_state.pop("auto_download_pending", None)
        st.error(f"Conversion stopped: {_friendly_error(exc)}")
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)


def _load_downloads(
    result,
    *,
    source_name: str | None = None,
    filename_prefix: str = "",
) -> list[tuple[str, str, bytes]]:
    downloads: list[tuple[str, str, bytes]] = []

    def add(label: str, path: Path) -> None:
        path.chmod(0o600)
        button_label = (
            f"Download {source_name} — {label}"
            if source_name
            else f"Download {label}"
        )
        downloads.append(
            (
                button_label,
                f"{filename_prefix}{path.name}",
                path.read_bytes(),
            )
        )

    if result.editable_docx:
        add("translation-ready Word", result.editable_docx)
    if result.faithful_docx:
        add("exact-look Word", result.faithful_docx)
    if result.layout_json:
        add("OCR layout data (JSON)", result.layout_json)
    if result.usage_json:
        add("API usage report (JSON)", result.usage_json)
    return downloads


def _load_usage_report(result) -> dict | None:
    if not result.usage_json:
        return None
    try:
        return json.loads(result.usage_json.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _combine_usage_reports(
    reports: list[tuple[str, dict]],
) -> dict | None:
    if not reports:
        return None

    integer_fields = (
        "api_calls",
        "metered_responses",
        "unmetered_attempts",
        "input_tokens",
        "cached_input_tokens",
        "cache_write_input_tokens",
        "output_tokens",
        "reasoning_tokens",
        "total_tokens",
    )
    summaries = [report.get("summary", {}) for _, report in reports]
    summary = {
        field: sum(int(item.get(field, 0) or 0) for item in summaries)
        for field in integer_fields
    }
    costs = [item.get("estimated_cost_usd") for item in summaries]
    summary["estimated_cost_usd"] = (
        None
        if any(cost is None for cost in costs)
        else round(sum(float(cost) for cost in costs), 6)
    )
    summary["known_minimum_cost_usd"] = round(
        sum(float(item.get("known_minimum_cost_usd", 0) or 0) for item in summaries),
        6,
    )

    model_totals: dict[str, dict] = {}
    for _, report in reports:
        for entry in report.get("by_model", []):
            model = str(entry.get("model", ""))
            if not model:
                continue
            total = model_totals.setdefault(
                model,
                {
                    "model": model,
                    "api_calls": 0,
                    "input_tokens": 0,
                    "output_tokens": 0,
                    "reasoning_tokens": 0,
                    "estimated_cost_usd": 0.0,
                    "_cost_unknown": False,
                },
            )
            for field in (
                "api_calls",
                "input_tokens",
                "output_tokens",
                "reasoning_tokens",
            ):
                total[field] += int(entry.get(field, 0) or 0)
            entry_cost = entry.get("estimated_cost_usd")
            if entry_cost is None:
                total["_cost_unknown"] = True
            else:
                total["estimated_cost_usd"] += float(entry_cost)

    by_model = []
    for model in sorted(model_totals):
        entry = model_totals[model]
        unknown = entry.pop("_cost_unknown")
        entry["estimated_cost_usd"] = (
            None
            if unknown
            else round(float(entry["estimated_cost_usd"]), 6)
        )
        by_model.append(entry)

    calls = []
    for document_name, report in reports:
        calls.extend(
            {"document": document_name, **call}
            for call in report.get("calls", [])
        )

    return {
        "pricing": reports[0][1].get("pricing", {}),
        "documents": [name for name, _ in reports],
        "summary": summary,
        "by_model": by_model,
        "calls": calls,
    }


def _build_download_zip(downloads: list[tuple[str, str, bytes]]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for _, filename, data in downloads:
            archive.writestr(filename, data)
    return buffer.getvalue()


def _usage_summary() -> None:
    report = st.session_state.get("api_usage_report")
    if not report:
        return
    summary = report.get("summary", {})
    cost = summary.get("estimated_cost_usd")
    if cost is None:
        cost_label = "Unknown"
    elif cost < 0.01:
        cost_label = f"${cost:,.4f}"
    else:
        cost_label = f"${cost:,.2f}"

    st.subheader("API usage")
    calls, input_tokens, output_tokens, estimated_cost = st.columns(4)
    calls.metric("API calls", f"{summary.get('api_calls', 0):,}")
    input_tokens.metric("Input tokens", f"{summary.get('input_tokens', 0):,}")
    output_tokens.metric("Output tokens", f"{summary.get('output_tokens', 0):,}")
    estimated_cost.metric("Estimated cost", cost_label)
    model_names = ", ".join(
        entry.get("model", "") for entry in report.get("by_model", []) if entry.get("model")
    ) or "No metered response"
    unmetered = summary.get("unmetered_attempts", 0)
    unknown_note = ""
    if unmetered:
        minimum = summary.get("known_minimum_cost_usd", 0)
        unknown_note = (
            f" {unmetered:,} request attempt(s) returned no token data, so the exact "
            f"cost is unknown; the metered minimum is ${minimum:,.4f}."
        )
    st.caption(
        f"Model(s): {model_names} · "
        f"Metered responses: {summary.get('metered_responses', 0):,} · "
        f"Cached input: {summary.get('cached_input_tokens', 0):,} · "
        f"Cache-written input: {summary.get('cache_write_input_tokens', 0):,} · "
        f"Reasoning: {summary.get('reasoning_tokens', 0):,} "
        "(already included in output) · "
        f"Total tokens: {summary.get('total_tokens', 0):,}. "
        "Cost is estimated in USD at default-tier standard API rates; your billing "
        f"dashboard is authoritative.{unknown_note}"
    )


def _downloads() -> None:
    downloads = st.session_state.get("downloads", [])
    if not downloads:
        return
    st.subheader("Your files")
    download_id = st.session_state.get("download_id")
    for index, (label, filename, data) in enumerate(downloads):
        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        if filename.endswith(".json"):
            mime = "application/json"
        elif filename.endswith(".zip"):
            mime = "application/zip"
        container_key = (
            f"layoutlens_download_{download_id}"
            if index == 0 and download_id
            else None
        )
        with st.container(key=container_key):
            st.download_button(
                label, data=data, file_name=filename, mime=mime,
                use_container_width=True, on_click="ignore",
            )
    pending = st.session_state.pop("auto_download_pending", None)
    if pending and pending == download_id:
        trigger_download(pending)
    st.caption("Your download starts automatically. If it does not start, use a download button above.")


def _feature_row() -> None:
    st.markdown(
        """
        <div class="feature-row">
          <div><strong>Text + tables</strong><span>Editable Word elements</span></div>
          <div><strong>Natural reflow</strong><span>Built for translation</span></div>
          <div><strong>Visual fallback</strong><span>Pixel-faithful reference</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _privacy_note() -> None:
    st.markdown(
        '<div class="privacy">When Translation-ready Word or JSON is selected, page images and optional embedded PDF text are sent to the OpenAI API for OCR. Exact-look-only conversions stay local.</div>',
        unsafe_allow_html=True,
    )


def _styles() -> None:
    st.markdown(
        """
        <style>
        :root { --ink: #17251f; --muted: #66756e; --accent: #126149; --paper: #fbfaf5; }
        .stApp { background: linear-gradient(180deg, #f4f1e8 0%, #fbfaf5 32%, #fbfaf5 100%); }
        .block-container { max-width: 780px; padding-top: 4.2rem; padding-bottom: 5rem; }
        h1 { color: var(--ink) !important; font-size: clamp(2.4rem, 6vw, 4.5rem) !important; letter-spacing: -0.055em !important; line-height: .98 !important; max-width: 720px; }
        h2, h3 { color: var(--ink) !important; }
        [data-testid="stWidgetLabel"] p,
        [data-testid="stCheckbox"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p { color: var(--ink) !important; }
        .eyebrow { color: var(--accent); font-size: .78rem; font-weight: 760; letter-spacing: .14em; text-transform: uppercase; margin-bottom: .8rem; }
        [data-testid="stCaptionContainer"] p { color: var(--muted); font-size: 1.08rem; max-width: 650px; }
        .feature-row { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin: 1.8rem 0 1.4rem; }
        .feature-row div { border: 1px solid #d9ddd7; border-radius: 14px; padding: 14px 15px; background: rgba(255,255,255,.66); }
        .feature-row strong, .feature-row span { display: block; }
        .feature-row strong { color: var(--ink); font-size: .92rem; }
        .feature-row span { color: var(--muted); font-size: .78rem; margin-top: 3px; }
        [data-testid="stFileUploaderDropzone"] { background: #ffffff; border: 1.5px dashed #9baba2; border-radius: 18px; min-height: 150px; }
        [data-testid="stFileUploaderDropzone"]:hover { border-color: var(--accent); background: #f7fbf8; }
        .empty-hint, .privacy { color: var(--muted); font-size: .82rem; margin-top: .8rem; }
        .privacy { border-left: 3px solid #b8c9c0; padding-left: 10px; }
        .file-card { display: flex; justify-content: space-between; gap: 16px; align-items: center; padding: 14px 16px; margin: 1rem 0; background: #edf3ef; border-radius: 12px; color: var(--ink); font-weight: 650; }
        .file-card span:last-child { color: var(--muted); font-size: .82rem; font-weight: 500; white-space: nowrap; }
        .stButton > button[kind="primary"] { background: var(--accent); border-radius: 12px; min-height: 50px; font-weight: 720; }
        .stDownloadButton > button { border-color: #9cb4a9; border-radius: 12px; min-height: 46px; }
        [data-testid="stMetric"] { background: rgba(255,255,255,.66); border: 1px solid #d9ddd7; border-radius: 12px; padding: 10px; }
        [data-testid="stMetricLabel"] p, [data-testid="stMetricValue"] { color: var(--ink) !important; }
        @media (max-width: 650px) { .feature-row { grid-template-columns: 1fr; } .file-card { align-items: flex-start; flex-direction: column; } }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _safe_filename(value: str) -> str:
    name = Path(value).name
    cleaned = "".join(char if char.isalnum() or char in {"-", "_", "."} else "-" for char in name)
    return cleaned or "document.pdf"


def _escape(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _phase_progress(message: str, current: int, total: int) -> float:
    ratio = min(1.0, current / total) if total else 0.0
    lowered = message.lower()
    if "render" in lowered:
        return 0.02 + 0.13 * ratio
    if "exact-look" in lowered:
        return 0.15 + 0.08 * ratio
    if "read" in lowered or "vision" in lowered:
        return 0.23 + 0.57 * ratio
    if "translation-ready" in lowered:
        return 0.80 + 0.16 * ratio
    return 0.96 * ratio


def _cleanup_stale_jobs() -> None:
    if not JOBS_DIR.is_dir():
        return
    cutoff = time.time() - JOB_RETENTION_SECONDS
    for candidate in JOBS_DIR.iterdir():
        try:
            if candidate.is_dir() and candidate.stat().st_mtime < cutoff:
                shutil.rmtree(candidate, ignore_errors=True)
        except OSError:
            continue


def _friendly_error(error: Exception) -> str:
    message = str(error).strip()
    if "OPENAI_API_KEY" in message:
        return message
    if "401" in message or "authentication" in message.lower():
        return "The OpenAI API key was rejected. Check the key in your environment file."
    if "rate" in message.lower() and "limit" in message.lower():
        return "The OpenAI API is temporarily rate-limited. Wait a moment, then retry."
    return message or "An unexpected conversion error occurred."
