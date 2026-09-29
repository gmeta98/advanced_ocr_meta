from __future__ import annotations

import io
import zipfile

import pytest

import fitz
from streamlit.testing.v1 import AppTest

from ocr_reader.models import ConversionResult
from ocr_reader.ui import _build_download_zip, _combine_usage_reports


@pytest.fixture(autouse=True)
def isolated_ui_settings(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_PASSWORD", "test-password")
    monkeypatch.setenv("OPENAI_API_KEY", "test-api-key")
    monkeypatch.setattr("ocr_reader.config.load_environment", lambda: None)
    monkeypatch.setattr("ocr_reader.ui.load_environment", lambda: None)
    monkeypatch.setattr("ocr_reader.ui.JOBS_DIR", tmp_path / "jobs")


def _signed_in_app():
    app = AppTest.from_file("app.py").run()
    assert not app.exception
    assert not app.file_uploader
    app.text_input[0].set_value("test-password")
    next(button for button in app.button if button.label == "Open LayoutLens").click().run()
    assert not app.exception
    return app


def test_app_shows_model_choices_and_translation_ready_defaults():
    document = fitz.open()
    document.new_page(width=612, height=792)
    pdf_bytes = document.tobytes()
    document.close()

    app = _signed_in_app()
    app.run()
    assert not app.exception
    assert app.file_uploader[0].accept_multiple_files is True
    app.file_uploader[0].set_value(
        ("sample.pdf", pdf_bytes, "application/pdf")
    ).run()
    assert not app.exception

    controls = {control.label: control for control in app.segmented_control}
    assert controls["Quality"].value == "Balanced"
    assert controls["OCR model"].options == [
        "GPT-6 Sol",
        "GPT-6 Luna",
    ]
    assert controls["OCR model"].value == "GPT-6 Sol"

    checkboxes = {checkbox.label: checkbox.value for checkbox in app.checkbox}
    assert checkboxes["Translation-ready Word"] is True
    assert checkboxes["Exact-look reference"] is False
    assert checkboxes["Include OCR layout data (JSON)"] is False


def test_multiple_uploads_switch_to_all_pages_batch_mode():
    document = fitz.open()
    document.new_page(width=612, height=792)
    pdf_bytes = document.tobytes()
    document.close()

    app = _signed_in_app()
    app.run()
    app.file_uploader[0].set_value(
        [
            ("first.pdf", pdf_bytes, "application/pdf"),
            ("second.pdf", pdf_bytes, "application/pdf"),
        ]
    ).run()
    assert not app.exception
    assert not app.number_input
    assert "Batch mode will process every page" in app.info[0].value
    assert any(button.label == "Convert 2 PDFs to Word" for button in app.button)


def test_batch_conversion_processes_each_pdf_and_offers_zip(monkeypatch):
    document = fitz.open()
    document.new_page(width=612, height=792)
    pdf_bytes = document.tobytes()
    document.close()
    converted: list[tuple[str, list[int]]] = []

    def fake_convert(pdf_path, output_dir, *, page_numbers, **_kwargs):
        converted.append((pdf_path.name, page_numbers))
        output_dir.mkdir(parents=True, exist_ok=True)
        faithful = output_dir / f"{pdf_path.stem}-exact-look.docx"
        faithful.write_bytes(pdf_path.name.encode("utf-8"))
        return ConversionResult(
            editable_docx=None,
            faithful_docx=faithful,
            layout_json=None,
            processed_pages=tuple(page_numbers),
        )

    monkeypatch.setattr("ocr_reader.ui.convert_pdf", fake_convert)
    app = _signed_in_app()
    app.run()
    app.file_uploader[0].set_value(
        [
            ("first.pdf", pdf_bytes, "application/pdf"),
            ("second.pdf", pdf_bytes, "application/pdf"),
        ]
    ).run()
    for checkbox in app.checkbox:
        if checkbox.label == "Translation-ready Word":
            checkbox.set_value(False)
        elif checkbox.label == "Exact-look reference":
            checkbox.set_value(True)
    app.run()
    convert_button = next(
        button for button in app.button if button.label == "Convert 2 PDFs to Word"
    )
    convert_button.click().run()
    assert not app.exception
    assert converted == [("first.pdf", [1]), ("second.pdf", [1])]
    download_labels = [button.label for button in app.download_button]
    assert download_labels[0] == "Download all results (ZIP)"
    assert "Download first.pdf — exact-look Word" in download_labels
    assert "Download second.pdf — exact-look Word" in download_labels


def test_usage_summary_labels_unknown_cost_without_calling_it_zero():
    app = AppTest.from_string(
        "from ocr_reader.ui import _usage_summary\n_usage_summary()"
    )
    app.session_state["api_usage_report"] = {
        "summary": {
            "api_calls": 2,
            "metered_responses": 1,
            "unmetered_attempts": 1,
            "input_tokens": 1_000,
            "cached_input_tokens": 100,
            "cache_write_input_tokens": 50,
            "output_tokens": 200,
            "reasoning_tokens": 75,
            "total_tokens": 1_200,
            "estimated_cost_usd": None,
            "known_minimum_cost_usd": 0.005,
        },
        "by_model": [{"model": "gpt-6-sol"}],
    }
    app.run()
    assert not app.exception
    metrics = {metric.label: metric.value for metric in app.metric}
    assert metrics["API calls"] == "2"
    assert metrics["Estimated cost"] == "Unknown"
    assert "gpt-6-sol" in app.caption[0].value
    assert "metered minimum is $0.0050" in app.caption[0].value


def test_combined_usage_and_batch_zip_include_each_document():
    first = {
        "pricing": {"currency": "USD"},
        "summary": {
            "api_calls": 1,
            "metered_responses": 1,
            "unmetered_attempts": 0,
            "input_tokens": 100,
            "cached_input_tokens": 10,
            "cache_write_input_tokens": 0,
            "output_tokens": 20,
            "reasoning_tokens": 5,
            "total_tokens": 120,
            "estimated_cost_usd": 0.01,
            "known_minimum_cost_usd": 0.01,
        },
        "by_model": [
            {
                "model": "gpt-6-sol",
                "api_calls": 1,
                "input_tokens": 100,
                "output_tokens": 20,
                "reasoning_tokens": 5,
                "estimated_cost_usd": 0.01,
            }
        ],
        "calls": [{"page_number": 1, "stage": "extraction"}],
    }
    second = {
        **first,
        "summary": {
            **first["summary"],
            "input_tokens": 200,
            "total_tokens": 220,
            "estimated_cost_usd": 0.02,
            "known_minimum_cost_usd": 0.02,
        },
        "calls": [{"page_number": 2, "stage": "extraction"}],
    }
    combined = _combine_usage_reports(
        [("first.pdf", first), ("second.pdf", second)]
    )
    assert combined is not None
    assert combined["documents"] == ["first.pdf", "second.pdf"]
    assert combined["summary"]["api_calls"] == 2
    assert combined["summary"]["input_tokens"] == 300
    assert combined["summary"]["estimated_cost_usd"] == 0.03
    assert [call["document"] for call in combined["calls"]] == [
        "first.pdf",
        "second.pdf",
    ]

    archive_bytes = _build_download_zip(
        [
            ("First", "01-first.docx", b"first"),
            ("Second", "02-second.docx", b"second"),
        ]
    )
    with zipfile.ZipFile(io.BytesIO(archive_bytes)) as archive:
        assert archive.namelist() == ["01-first.docx", "02-second.docx"]
        assert archive.read("02-second.docx") == b"second"


def _one_page_upload(app):
    document = fitz.open()
    document.new_page(width=612, height=792)
    data = document.tobytes()
    document.close()
    app.file_uploader[0].set_value(("sample.pdf", data, "application/pdf")).run()
    return app


@pytest.mark.parametrize("model_label,model_id,effort", [
    ("GPT-6 Sol", "gpt-6-sol", "medium"),
    ("GPT-6 Luna", "gpt-6-luna", "low"),
])
def test_selected_model_routes_and_auto_download_runs_once(monkeypatch, model_label, model_id, effort):
    requests = []
    triggers = []

    def fake_convert(pdf_path, output_dir, *, page_numbers, settings, **kwargs):
        requests.append((settings.model, settings.reasoning_effort))
        output_dir.mkdir(parents=True, exist_ok=True)
        word = output_dir / "sample-translation-ready.docx"
        word.write_bytes(b"test Word output")
        usage = output_dir / "sample-api-usage.json"
        usage.write_text('{"summary": {"api_calls": 1, "estimated_cost_usd": 0.01}}')
        return ConversionResult(word, None, None, (1,), usage)

    monkeypatch.setattr("ocr_reader.ui.convert_pdf", fake_convert)
    monkeypatch.setattr("ocr_reader.ui.trigger_download", triggers.append)
    app = _one_page_upload(_signed_in_app())
    next(control for control in app.segmented_control if control.label == "OCR model").set_value(model_label).run()
    next(button for button in app.button if button.label == "Convert to Word").click().run()
    assert not app.exception
    assert requests == [(model_id, effort)]
    assert len(triggers) == 1
    assert app.session_state["downloads"][0][1] == "sample-translation-ready.docx"
    app.run()
    assert len(triggers) == 1
    next(button for button in app.button if button.label == "Convert to Word").click().run()
    assert len(triggers) == 2
    assert triggers[0] != triggers[1]


def test_both_word_outputs_auto_download_as_one_zip(monkeypatch):
    triggers = []

    def fake_convert(pdf_path, output_dir, **kwargs):
        output_dir.mkdir(parents=True, exist_ok=True)
        editable = output_dir / "sample-translation-ready.docx"
        faithful = output_dir / "sample-exact-look.docx"
        editable.write_bytes(b"editable")
        faithful.write_bytes(b"faithful")
        return ConversionResult(editable, faithful, None, (1,))

    monkeypatch.setattr("ocr_reader.ui.convert_pdf", fake_convert)
    monkeypatch.setattr("ocr_reader.ui.trigger_download", triggers.append)
    app = _one_page_upload(_signed_in_app())
    next(box for box in app.checkbox if box.label == "Exact-look reference").set_value(True).run()
    next(button for button in app.button if button.label == "Convert to Word").click().run()
    assert not app.exception
    assert len(triggers) == 1
    name, data = app.session_state["downloads"][0][1:]
    assert name == "sample-results.zip"
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        assert sorted(archive.namelist()) == ["sample-exact-look.docx", "sample-translation-ready.docx"]


def test_failed_conversion_does_not_auto_download_usage_only(monkeypatch):
    from ocr_reader.pipeline import PartialConversionError
    triggers = []

    def fake_convert(pdf_path, output_dir, **kwargs):
        output_dir.mkdir(parents=True, exist_ok=True)
        usage = output_dir / "sample-api-usage.json"
        usage.write_text('{"summary": {"api_calls": 1, "estimated_cost_usd": null}}')
        result = ConversionResult(None, None, None, (1,), usage)
        raise PartialConversionError("OCR failed", result)

    monkeypatch.setattr("ocr_reader.ui.convert_pdf", fake_convert)
    monkeypatch.setattr("ocr_reader.ui.trigger_download", triggers.append)
    app = _one_page_upload(_signed_in_app())
    next(button for button in app.button if button.label == "Convert to Word").click().run()
    assert not app.exception
    assert triggers == []
    assert len(app.download_button) == 1
