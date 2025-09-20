from __future__ import annotations
from pathlib import Path
from typing import Optional, Tuple
import pymupdf as fitz

class OcrAdapter:
    """
    Uses MuPDF pdfocr bindings (same as your original `pdfocr_tobytes`).
    Returns text prefixed with `_OCR_` for styling later.
    """
    def __init__(self, tessdata_dir: Optional[Path]):
        self.tessdata_dir = str(tessdata_dir) if tessdata_dir else None

    def ocr_clip_to_text(
        self,
        page: "fitz.Page",
        clip: Tuple[float, float, float, float],
        dpi: int,
        scale: Optional[float]
    ) -> str:
        if scale is not None:
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat, clip=fitz.Rect(clip))
        else:
            pix = page.get_pixmap(clip=fitz.Rect(clip), dpi=dpi)

        pdfdata = pix.pdfocr_tobytes(language="eng", tessdata=self.tessdata_dir)
        try:
            with fitz.open("pdf", pdfdata) as clipdoc:
                return "_OCR_" + clipdoc[0].get_text()
        finally:
            try:
                del pdfdata
            except Exception:
                pass
