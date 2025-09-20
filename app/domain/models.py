from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Tuple, List

Rect = Tuple[float, float, float, float]

@dataclass(frozen=True)
class AreaSpec:
    title: str
    rect: Rect

@dataclass(frozen=True)
class OcrSettings:
    mode: str           # "Default" | "OCR-All" | "Text1st+Image-beta"
    dpi: int
    tessdata_dir: Optional[Path] = None
    scale: Optional[float] = None  # optional upscale for images (kept for parity)

@dataclass(frozen=True)
class ExtractionRequest:
    pdf_paths: Iterable[Path]
    output_excel: Path
    areas: List[AreaSpec]
    revision_area: Optional[AreaSpec]
    revision_regex: str
    ocr: OcrSettings
    pdf_root: Optional[Path] = None  # used for relative folder column (defaults to first file's parent)
