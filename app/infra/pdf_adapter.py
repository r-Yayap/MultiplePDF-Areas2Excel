from __future__ import annotations
from contextlib import contextmanager
from pathlib import Path
from typing import Iterable, List, Optional, Tuple
import pymupdf as fitz

RectT = Tuple[float, float, float, float]

class PdfAdapter:
    def page_count(self, path: str | Path) -> int:
        with fitz.open(str(path)) as doc:
            return doc.page_count

    @contextmanager
    def open(self, path: str | Path):
        doc = fitz.open(str(path), filetype="pdf")
        try:
            if not doc.is_pdf:
                raise ValueError("Not a valid PDF")
            fitz.TOOLS.set_small_glyph_heights(True)
            yield doc
        finally:
            doc.close()

    def get_text(self, page: "fitz.Page", clip: RectT) -> str:
        return page.get_text("text", clip=fitz.Rect(clip))

    def words_count(self, page: "fitz.Page", clip: RectT) -> int:
        try:
            return len(page.get_text("words", clip=fitz.Rect(clip)))
        except Exception:
            return -1

    def find_table_rows(self, page: "fitz.Page", clip: RectT) -> Optional[List[List[str]]]:
        data = None
        try:
            tabs = page.find_tables(clip=fitz.Rect(clip))
            if tabs and tabs.tables:
                data = tabs.tables[0].extract()
        except Exception:
            data = None

        if data and len(data) >= 2:
            return data

        # fallback: cut & paste to temp page
        tmp = fitz.open()
        try:
            r = fitz.Rect(clip)
            dst = tmp.new_page(width=r.width, height=r.height)
            dst.show_pdf_page(fitz.Rect(0, 0, r.width, r.height), page.parent, page.number, clip=r)
            try:
                tabs = dst.find_tables()
                if tabs and tabs.tables:
                    data = tabs.tables[0].extract()
            except Exception:
                data = None
        finally:
            tmp.close()

        return data if data and len(data) >= 2 else None

    def page_rect(self, page: "fitz.Page") -> RectT:
        r = page.rect
        return (r.x0, r.y0, r.x1, r.y1)

    def render_pixmap(self, page: "fitz.Page", clip: RectT, dpi: int = 150, scale: Optional[float] = None):
        if scale is not None:
            mat = fitz.Matrix(scale, scale)
            return page.get_pixmap(matrix=mat, clip=fitz.Rect(clip))
        return page.get_pixmap(clip=fitz.Rect(clip), dpi=dpi)

    def remove_rotation(self, page: "fitz.Page"):
        # Normalize rotation so coordinates work in unrotated basis.
        try:
            page.remove_rotation()
        except Exception:
            pass
