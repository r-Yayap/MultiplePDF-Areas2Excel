from typing import Optional
from app.ports.ocr_port import OcrPort
from app.domain.models import OcrSettings
from app.ports.pdf_port import Rect

class NoopOCR(OcrPort):
    def ocr_image_region(self, pdf_path: str, page_index: int, rect: Rect, settings: OcrSettings) -> Optional[str]:
        return None  # Replace with real OCR later if needed
