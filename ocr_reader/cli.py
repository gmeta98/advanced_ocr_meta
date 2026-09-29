from __future__ import annotations

import argparse
from dataclasses import replace
from pathlib import Path

from .config import MODEL_PRESETS, Settings, apply_quality_preset
from .pdf import inspect_pdf, parse_page_spec
from .pipeline import PartialConversionError, convert_pdf


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="layoutlens",
        description="Convert PDF pages to translation-ready and exact-look Word documents.",
    )
    parser.add_argument("pdf", type=Path, help="Input PDF")
    parser.add_argument("--output", type=Path, default=Path("output"), help="Output folder")
    parser.add_argument("--pages", help="Page selection such as 1-3,7")
    parser.add_argument(
        "--mode",
        choices=("editable", "faithful", "both"),
        default="both",
        help="Word export type",
    )
    parser.add_argument(
        "--quality",
        choices=("balanced", "maximum"),
        default="balanced",
        help="OCR quality preset",
    )
    parser.add_argument(
        "--model",
        choices=("sol", "luna"),
        default="sol",
        help="OCR model: sol=GPT-6 Sol (default), luna=GPT-6 Luna",
    )
    parser.add_argument("--json", action="store_true", help="Also save structured OCR layout JSON")
    parser.add_argument(
        "--no-native-text",
        action="store_true",
        help="Do not send embedded PDF text to the OCR model as a transcription aid",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    info = inspect_pdf(args.pdf)
    pages = parse_page_spec(args.pages, info.page_count)
    needs_ocr = args.mode in {"editable", "both"} or args.json
    settings = Settings.from_environment(require_api_key=needs_ocr)
    preset = MODEL_PRESETS[f"GPT-6 {args.model.title()}"]
    settings = replace(
        settings,
        model=preset.model,
        reasoning_effort=preset.reasoning_effort,
        include_native_text=not args.no_native_text,
    )
    settings = apply_quality_preset(settings, args.quality)

    def progress(message: str, current: int, total: int) -> None:
        suffix = f" ({current}/{total})" if total else ""
        print(f"{message}{suffix}", flush=True)

    try:
        result = convert_pdf(
            args.pdf,
            args.output,
            page_numbers=pages,
            settings=settings,
            mode=args.mode,
            include_layout_json=args.json,
            progress=progress,
        )
    except PartialConversionError as exc:
        result = exc.result
        _print_result_paths(result)
        raise SystemExit(str(exc)) from exc
    _print_result_paths(result)


def _print_result_paths(result) -> None:
    for path in (
        result.editable_docx,
        result.faithful_docx,
        result.layout_json,
        result.usage_json,
    ):
        if path:
            print(path.resolve())


if __name__ == "__main__":
    main()
