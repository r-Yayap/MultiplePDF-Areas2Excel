import pymupdf as fitz
from typing import Tuple
from app.ports.pdf_port import PdfPort, Rect

class FitzPdf(PdfPort):
    def page_count(self, pdf_path: str) -> int:
        with fitz.open(pdf_path) as doc:
            return doc.page_count

    def get_page_size(self, pdf_path: str, page_index: int) -> Tuple[float, float]:
        with fitz.open(pdf_path) as doc:
            r = doc[page_index].rect
            return float(r.width), float(r.height)

    def extract_text_in_rect(self, pdf_path: str, page_index: int, rect: Rect) -> str:
        x0, y0, x1, y1 = rect
        with fitz.open(pdf_path) as doc:
            page = doc[page_index]
            # PyMuPDF uses a Rect object for region text extraction
            region = fitz.Rect(x0, y0, x1, y1)
            return page.get_text("text", clip=region).strip()
