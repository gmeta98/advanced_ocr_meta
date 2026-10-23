from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import fitz
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

from .models import RenderedPage


MAX_RENDER_PIXELS = 60_000_000


@dataclass(frozen=True, slots=True)
class PdfInfo:
    page_count: int
    title: str
    author: str


def inspect_pdf(path: Path) -> PdfInfo:
    _validate_pdf(path)
    with fitz.open(path) as document:
        _reject_locked_pdf(document)
        metadata = document.metadata or {}
        return PdfInfo(
            page_count=document.page_count,
            title=(metadata.get("title") or path.stem).strip(),
            author=(metadata.get("author") or "").strip(),
        )


def inspect_pdf_bytes(data: bytes) -> PdfInfo:
    if not data.startswith(b"%PDF-"):
        raise ValueError("The uploaded file is not a valid PDF.")
    with fitz.open(stream=data, filetype="pdf") as document:
        _reject_locked_pdf(document)
        metadata = document.metadata or {}
        return PdfInfo(
            page_count=document.page_count,
            title=(metadata.get("title") or "PDF document").strip(),
            author=(metadata.get("author") or "").strip(),
        )


def render_pages(
    path: Path,
    output_dir: Path,
    page_numbers: Iterable[int],
    *,
    dpi: int,
    enhance_scans: bool = False,
) -> list[RenderedPage]:
    _validate_pdf(path)
    output_dir.mkdir(parents=True, exist_ok=True)
    selected = sorted(set(page_numbers))
    if not selected:
        raise ValueError("Select at least one page.")

    rendered: list[RenderedPage] = []
    with fitz.open(path) as document:
        _reject_locked_pdf(document)
        for page_number in selected:
            if page_number < 1 or page_number > document.page_count:
                raise ValueError(
                    f"Page {page_number} is outside this {document.page_count}-page PDF."
                )
            page = document.load_page(page_number - 1)
            rect = page.rect
            width_px = max(1, round(rect.width * dpi / 72.0))
            height_px = max(1, round(rect.height * dpi / 72.0))
            if width_px * height_px > MAX_RENDER_PIXELS:
                raise ValueError(
                    f"Page {page_number} would render to {width_px:,} x {height_px:,} pixels. "
                    "Reduce OCR_RENDER_DPI or use a PDF with a standard page size."
                )
            pixmap = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
            image_path = output_dir / f"page-{page_number:04d}.png"
            pixmap.save(image_path)
            native_text = _native_text_with_layout(page)
            ocr_image_path = None
            if enhance_scans and _looks_scanned(page, native_text):
                ocr_image_path = output_dir / f"page-{page_number:04d}-ocr.png"
                _enhance_ocr_image(image_path, ocr_image_path)
            rendered.append(
                RenderedPage(
                    page_number=page_number,
                    width_pt=float(rect.width),
                    height_pt=float(rect.height),
                    image_path=image_path,
                    native_text=native_text,
                    ocr_image_path=ocr_image_path,
                )
            )
    return rendered


def parse_page_spec(spec: str | None, page_count: int) -> list[int]:
    if not spec:
        return list(range(1, page_count + 1))
    pages: set[int] = set()
    for raw_part in spec.split(","):
        part = raw_part.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = part.split("-", 1)
            start, end = int(start_text), int(end_text)
            if start > end:
                start, end = end, start
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part))
    if not pages:
        raise ValueError("The page selection is empty.")
    invalid = sorted(page for page in pages if page < 1 or page > page_count)
    if invalid:
        raise ValueError(f"Invalid page selection: {', '.join(map(str, invalid))}")
    return sorted(pages)


def _validate_pdf(path: Path) -> None:
    if not path.is_file():
        raise FileNotFoundError(path)
    with path.open("rb") as handle:
        if handle.read(5) != b"%PDF-":
            raise ValueError(f"{path.name} is not a valid PDF.")


def _reject_locked_pdf(document: fitz.Document) -> None:
    if document.needs_pass:
        raise ValueError(
            "This PDF is password-protected. Remove the password in your PDF app, then try again."
        )


def _native_text_with_layout(page: fitz.Page) -> str:
    blocks = []
    for block in page.get_text("blocks", sort=True):
        if len(block) < 7 or block[6] != 0:
            continue
        x0, y0, x1, y1, text = block[:5]
        value = " ".join(part.strip() for part in str(text).splitlines() if part.strip())
        if not value:
            continue
        printable = sum(character.isprintable() for character in value) / max(1, len(value))
        if printable < 0.94 or value.count("\ufffd") > max(1, len(value) // 100):
            continue
        nx0 = max(0.0, min(1000.0, x0 / max(1.0, page.rect.width) * 1000.0))
        ny0 = max(0.0, min(1000.0, y0 / max(1.0, page.rect.height) * 1000.0))
        nx1 = max(nx0, min(1000.0, x1 / max(1.0, page.rect.width) * 1000.0))
        ny1 = max(ny0, min(1000.0, y1 / max(1.0, page.rect.height) * 1000.0))
        blocks.append(f"[box {nx0:.1f},{ny0:.1f},{nx1:.1f},{ny1:.1f}] {value}")
    joined = "\n".join(blocks).strip()
    if len(joined) < 20:
        return ""
    return joined


def _enhance_ocr_image(source: Path, destination: Path) -> None:
    with Image.open(source) as image:
        rgb = image.convert("RGB")
        enhanced = ImageOps.autocontrast(rgb, cutoff=0.35, preserve_tone=True)
        enhanced = ImageEnhance.Contrast(enhanced).enhance(1.06)
        enhanced = enhanced.filter(ImageFilter.UnsharpMask(radius=1.1, percent=130, threshold=3))
        enhanced.save(destination, format="PNG", optimize=True)


def _looks_scanned(page: fitz.Page, native_text: str) -> bool:
    """Detect image-backed scans, including scans with a hidden OCR text layer."""
    if not native_text:
        return True

    page_area = max(1.0, page.rect.get_area())
    for image in page.get_images(full=True):
        xref = image[0]
        try:
            rectangles = page.get_image_rects(xref)
        except (RuntimeError, ValueError):
            continue
        for rectangle in rectangles:
            visible = rectangle & page.rect
            if not visible.is_empty and visible.get_area() / page_area >= 0.60:
                return True
    return False
