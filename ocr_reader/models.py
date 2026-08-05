from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict


ElementKind = Literal[
    "title",
    "heading",
    "paragraph",
    "list_item",
    "table",
    "form_field",
    "header",
    "footer",
    "page_number",
    "graphic",
    "signature",
    "stamp",
    "barcode",
    "other",
]

FontFamily = Literal["sans_serif", "serif", "monospace", "handwriting", "unknown"]
Alignment = Literal["left", "center", "right", "justified"]


class NormalizedBox(BaseModel):
    """A rectangle in a 0-1000 coordinate space."""

    model_config = ConfigDict(extra="forbid")

    x: float
    y: float
    width: float
    height: float


class VisionElement(BaseModel):
    model_config = ConfigDict(extra="forbid")

    kind: ElementKind
    bbox: NormalizedBox
    text: str
    reading_order: int
    font_family: FontFamily
    font_size_pt: float
    bold: bool
    italic: bool
    underline: bool
    color_hex: str
    background_hex: str
    alignment: Alignment
    table_cells: list[list[str]]
    table_header_rows: int
    table_header_fill_hex: str
    table_header_text_hex: str
    table_border_hex: str
    table_column_widths: list[float]
    confidence: float


class VisionPage(BaseModel):
    """Schema returned by the vision model for one rendered PDF page."""

    model_config = ConfigDict(extra="forbid")

    page_number: int
    detected_languages: list[str]
    rotation_degrees: int
    elements: list[VisionElement]


@dataclass(frozen=True, slots=True)
class RenderedPage:
    page_number: int
    width_pt: float
    height_pt: float
    image_path: Path
    native_text: str
    ocr_image_path: Path | None = None


@dataclass(frozen=True, slots=True)
class PageResult:
    source: RenderedPage
    layout: VisionPage


@dataclass(frozen=True, slots=True)
class APICallUsage:
    page_number: int
    stage: str
    model: str
    input_tokens: int
    cached_input_tokens: int
    cache_write_input_tokens: int
    output_tokens: int
    reasoning_tokens: int
    total_tokens: int


@dataclass(frozen=True, slots=True)
class ConversionResult:
    editable_docx: Path | None
    faithful_docx: Path | None
    layout_json: Path | None
    processed_pages: tuple[int, ...]
    usage_json: Path | None = None
