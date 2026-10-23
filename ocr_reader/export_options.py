from dataclasses import dataclass
from typing import Literal


@dataclass(frozen=True, slots=True)
class ExportOptions:
    """Word presentation options; these never change the OCR transcription."""

    style: Literal["compact", "source"] = "compact"
    marks: Literal["images", "editable"] = "images"
    skip_blank_pages: bool = True

    def __post_init__(self):
        if self.style not in {"compact", "source"}:
            raise ValueError("Unknown Word style")
        if self.marks not in {"images", "editable"}:
            raise ValueError("Unknown mark treatment")
