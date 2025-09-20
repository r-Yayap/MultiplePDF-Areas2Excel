from typing import Protocol, Optional
from app.domain.models import OcrSettings
from app.ports.pdf_port import Rect

class OcrPort(Protocol):
    def ocr_image_region(self, pdf_path: str, page_index: int, rect: Rect, settings: OcrSettings) -> Optional[str]: ...
