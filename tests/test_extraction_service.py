from pathlib import Path
from typing import Tuple
from app.domain.models import AreaSpec, OcrSettings, ExtractionRequest
from app.services.extraction_service import ExtractionService
from app.ports.pdf_port import Rect

# Fakes
class FakePdf:
    def __init__(self, text_map: dict[Tuple[int, Rect], str], page_counts: dict[str, int]):
        self.text_map = text_map
        self.page_counts = page_counts
    def page_count(self, pdf_path: str) -> int:
        return self.page_counts[pdf_path]
    def get_page_size(self, pdf_path: str, page_index: int):
        return (1000.0, 1000.0)
    def extract_text_in_rect(self, pdf_path: str, page_index: int, rect: Rect) -> str:
        return self.text_map.get((page_index, rect), "")

class FakeExcel:
    def __init__(self):
        self.headers = None
        self.rows = []
        self.saved = False
        self.path = None
    def open(self, output_path: str, headers):
        self.path = output_path
        self.headers = headers
    def append_row(self, row):
        self.rows.append(row)
    def save_and_close(self):
        self.saved = True

def test_extract_simple(tmp_path):
    # Arrange
    pdf_path = tmp_path / "doc.pdf"
    pdf_path.write_bytes(b"%PDF-FAKE")  # dummy

    rect1 = (0.0, 0.0, 100.0, 50.0)
    rect2 = (0.0, 60.0, 100.0, 120.0)

    fake_pdf = FakePdf(
        text_map={
            (0, rect1): "A1",
            (0, rect2): "Title line",
        },
        page_counts={str(pdf_path): 1},
    )
    fake_xlsx = FakeExcel()

    svc = ExtractionService(pdf=fake_pdf, excel=fake_xlsx, ocr=None)

    req = ExtractionRequest(
        pdf_paths=[pdf_path],
        output_excel=tmp_path / "out.xlsx",
        areas=[AreaSpec("Drawing No", rect1), AreaSpec("Drawing Title", rect2)],
        revision_area=None,
        revision_regex=r"[A-Z]\d?",
        ocr=OcrSettings(mode="Default", dpi=150, tessdata_dir=None),
    )

    # Act
    out = svc.extract(req)

    # Assert
    assert out.name == "out.xlsx"
    assert fake_xlsx.headers == ["File", "Page", "Drawing No", "Drawing Title"]
    assert fake_xlsx.rows == [[str(pdf_path), "1", "A1", "Title line"]]
    assert fake_xlsx.saved is True
