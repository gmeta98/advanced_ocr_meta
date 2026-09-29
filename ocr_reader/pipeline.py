from __future__ import annotations

import hashlib
import json
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Callable, Literal

from . import __version__
from .config import Settings
from .exporters import export_editable, export_faithful, export_layout_json
from .models import APICallUsage, ConversionResult, PageResult, RenderedPage
from .pdf import render_pages
from .usage import usage_report, write_usage_report
from .vision import PROMPT_VERSION, VisionOCR


ProgressCallback = Callable[[str, int, int], None]
Mode = Literal["editable", "faithful", "both"]


class PartialConversionError(RuntimeError):
    def __init__(self, message: str, result: ConversionResult):
        super().__init__(message)
        self.result = result


class _OCRBatchError(RuntimeError):
    def __init__(
        self,
        message: str,
        *,
        api_usage: tuple[APICallUsage, ...],
        request_attempts: int,
    ):
        super().__init__(message)
        self.api_usage = api_usage
        self.request_attempts = request_attempts


def convert_pdf(
    pdf_path: Path,
    output_dir: Path,
    *,
    page_numbers: list[int],
    settings: Settings,
    mode: Mode = "both",
    include_layout_json: bool = False,
    keep_rendered: bool = False,
    progress: ProgressCallback | None = None,
) -> ConversionResult:
    output_dir.mkdir(parents=True, exist_ok=True)
    needs_ocr = mode in {"editable", "both"} or include_layout_json
    render_dir = output_dir / "_rendered"
    _notify(progress, "Rendering PDF pages", 0, len(page_numbers))
    try:
        rendered = render_pages(
            pdf_path,
            render_dir,
            page_numbers,
            dpi=settings.render_dpi,
            enhance_scans=settings.enhance_scans and needs_ocr,
        )
    except Exception:
        _remove_rendered(render_dir, keep_rendered)
        raise
    _notify(progress, "PDF pages rendered", len(rendered), len(rendered))

    stem = _safe_stem(pdf_path.stem)
    editable_path = None
    faithful_path = None
    json_path = None
    usage_path = None
    api_usage: tuple[APICallUsage, ...] = ()
    api_request_attempts = 0

    if mode in {"faithful", "both"}:
        _notify(progress, "Building exact-look Word document", 0, 1)
        try:
            faithful_path = export_faithful(
                rendered, output_dir / f"{stem}-exact-look.docx"
            )
        except Exception:
            _remove_rendered(render_dir, keep_rendered)
            raise
        _notify(progress, "Exact-look Word document ready", 1, 1)

    page_results: list[PageResult] = []
    if needs_ocr:
        try:
            page_results, api_usage, api_request_attempts = _ocr_pages(
                rendered,
                settings,
                progress,
            )
        except Exception as exc:
            if isinstance(exc, _OCRBatchError):
                api_usage = exc.api_usage
                api_request_attempts = exc.request_attempts
                try:
                    usage_path = write_usage_report(
                        output_dir / f"{stem}-api-usage.json",
                        api_usage,
                        request_attempts=api_request_attempts,
                    )
                except Exception:
                    _remove_rendered(render_dir, keep_rendered)
                    raise
            _remove_rendered(render_dir, keep_rendered)
            _write_manifest(
                output_dir,
                pdf_path,
                rendered,
                [],
                settings,
                mode,
                failure=str(exc),
                api_usage=api_usage,
                api_request_attempts=api_request_attempts,
            )
            if faithful_path is None and usage_path is None:
                raise
            partial = ConversionResult(
                editable_docx=None,
                faithful_docx=faithful_path,
                layout_json=None,
                processed_pages=tuple(page.page_number for page in rendered),
                usage_json=usage_path,
            )
            if faithful_path is not None:
                message = (
                    "OCR could not finish, but the exact-look Word reference and "
                    "available API usage are ready."
                )
            else:
                message = (
                    f"OCR could not finish: {exc}. Available API usage was recorded."
                )
            raise PartialConversionError(
                message,
                partial,
            ) from exc

    try:
        if needs_ocr:
            usage_path = write_usage_report(
                output_dir / f"{stem}-api-usage.json",
                api_usage,
                request_attempts=api_request_attempts,
            )

        if mode in {"editable", "both"}:
            _notify(progress, "Building translation-ready Word document", 0, 1)
            editable_path = export_editable(
                page_results, output_dir / f"{stem}-translation-ready.docx"
            )
            _notify(progress, "Translation-ready Word document ready", 1, 1)

        if include_layout_json:
            json_path = export_layout_json(page_results, output_dir / f"{stem}-layout.json")

        _write_manifest(
            output_dir,
            pdf_path,
            rendered,
            page_results,
            settings,
            mode,
            api_usage=api_usage,
            api_request_attempts=api_request_attempts,
        )
    except Exception:
        _remove_rendered(render_dir, keep_rendered)
        raise
    _remove_rendered(render_dir, keep_rendered)
    return ConversionResult(
        editable_docx=editable_path,
        faithful_docx=faithful_path,
        layout_json=json_path,
        processed_pages=tuple(page.page_number for page in rendered),
        usage_json=usage_path,
    )


def _ocr_pages(
    pages: list[RenderedPage],
    settings: Settings,
    progress: ProgressCallback | None,
) -> tuple[list[PageResult], tuple[APICallUsage, ...], int]:
    engine = VisionOCR(settings)
    completed: list[PageResult] = []
    total = len(pages)
    failure: Exception | None = None
    _notify(progress, f"Reading page layout with {settings.model}", 0, total)
    try:
        with ThreadPoolExecutor(max_workers=min(settings.max_workers, total)) as executor:
            futures = {executor.submit(engine.extract_page, page): page for page in pages}
            for future in as_completed(futures):
                page = futures[future]
                layout = future.result()
                completed.append(PageResult(source=page, layout=layout))
                _notify(progress, f"Read page {page.page_number}", len(completed), total)
    except Exception as exc:
        failure = exc
    finally:
        try:
            engine.close()
        except Exception as exc:
            if failure is None:
                failure = exc
    if failure is not None:
        raise _OCRBatchError(
            str(failure),
            api_usage=engine.api_usage,
            request_attempts=engine.request_attempts,
        ) from failure
    completed.sort(key=lambda result: result.source.page_number)
    return completed, engine.api_usage, engine.request_attempts


def _notify(callback: ProgressCallback | None, message: str, current: int, total: int) -> None:
    if callback:
        callback(message, current, total)


def _remove_rendered(render_dir: Path, keep_rendered: bool) -> None:
    if not keep_rendered:
        shutil.rmtree(render_dir, ignore_errors=True)


def _safe_stem(value: str) -> str:
    cleaned = "".join(char if char.isalnum() or char in {"-", "_"} else "-" for char in value)
    return cleaned.strip("-") or "document"


def _write_manifest(
    output_dir: Path,
    pdf_path: Path,
    pages: list[RenderedPage],
    results: list[PageResult],
    settings: Settings,
    mode: Mode,
    failure: str | None = None,
    api_usage: tuple[APICallUsage, ...] = (),
    api_request_attempts: int = 0,
) -> None:
    manifest = {
        "layoutlens_version": __version__,
        "source_sha256": _file_sha256(pdf_path),
        "model": settings.model,
        "prompt_version": PROMPT_VERSION,
        "reasoning_effort": settings.reasoning_effort,
        "render_dpi": settings.render_dpi,
        "scan_enhancement": settings.enhance_scans,
        "verification_pass": settings.verify_ocr,
        "native_text_aid": settings.include_native_text,
        "mode": mode,
        "pages": [page.page_number for page in pages],
        "ocr_summary": [
            {
                "page": result.source.page_number,
                "languages": result.layout.detected_languages,
                "elements": len(result.layout.elements),
                "minimum_confidence": min(
                    (element.confidence for element in result.layout.elements), default=None
                ),
                "mean_confidence": (
                    sum(element.confidence for element in result.layout.elements)
                    / len(result.layout.elements)
                    if result.layout.elements
                    else None
                ),
            }
            for result in results
        ],
        "api_usage": usage_report(
            api_usage,
            request_attempts=api_request_attempts,
        ),
        "failure": failure,
    }
    (output_dir / "conversion.json").write_text(
        json.dumps(manifest, indent=2), encoding="utf-8"
    )


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()
