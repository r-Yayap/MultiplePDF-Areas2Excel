from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

Rect = Tuple[float, float, float, float]  # x0, y0, x1, y1

@dataclass(frozen=True)
class AreaSpec:
    title: str
    rect: Rect

@dataclass(frozen=True)
class OcrSettings:
    mode: str              # "Default" | "OCR-All" | "Text1st+Image-beta"
    dpi: int
    tessdata_dir: Optional[Path]

@dataclass(frozen=True)
class ExtractionRequest:
    pdf_paths: List[Path]
    output_excel: Path
    areas: List[AreaSpec]
    revision_area: Optional[AreaSpec]
    revision_regex: str
    ocr: OcrSettings
