# pdf_adapter.py
from __future__ import annotations
from contextlib import contextmanager
from pathlib import Path
from typing import List, Optional, Tuple
import os
import pymupdf as fitz

RectT = tuple[float, float, float, float]

def _safe_clip(page: "fitz.Page", clip: RectT) -> Optional[fitz.Rect]:
    r = fitz.Rect(clip).normalize() & page.rect
    if r.is_empty or r.width < 2 or r.height < 2:
        return None
    return r

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

    def page_rect(self, page: "fitz.Page") -> RectT:
        r = page.rect
        return (r.x0, r.y0, r.x1, r.y1)

    def get_text(self, page: "fitz.Page", clip: RectT) -> str:
        return page.get_text("text", clip=fitz.Rect(clip))

    def words_count(self, page: "fitz.Page", clip: RectT) -> int:
        try:
            return len(page.get_text("words", clip=fitz.Rect(clip)))
        except Exception:
            return -1

    def render_pixmap(self, page: "fitz.Page", clip: RectT, dpi: int = 150, scale: Optional[float] = None):
        if scale is not None:
            mat = fitz.Matrix(scale, scale)
            return page.get_pixmap(matrix=mat, clip=fitz.Rect(clip))
        return page.get_pixmap(clip=fitz.Rect(clip), dpi=dpi)

    def remove_rotation(self, page: "fitz.Page"):
        try:
            page.remove_rotation()
        except Exception:
            pass

    def find_table_rows(self, page: "fitz.Page", clip: RectT) -> Optional[List[List[str]]]:
        """
        Lean table finder for revision tables.
        Fallback mini-doc is disabled by default; enable with REV_TABLE_FALLBACK=1.
        """
        r = _safe_clip(page, clip)
        if r is None:
            return None

        # quick heuristics
        try:
            wc = len(page.get_text("words", clip=r))
        except Exception:
            wc = -1
        if 0 <= wc < 6:
            return None
        MAX_WORDS = int(os.getenv("REV_MAX_WORDS", "1500"))
        if wc > MAX_WORDS:
            return None

        # 1) try on-page
        try:
            tabs = page.find_tables(clip=r)
            if tabs and getattr(tabs, "tables", None):
                data = tabs.tables[0].extract()
                if data and len(data) >= 2:
                    return [[(c if isinstance(c, str) else ("" if c is None else str(c))).strip()
                             for c in row] for row in data]
        except Exception:
            pass

        # 2) optional fallback (heavy)
        if os.getenv("REV_TABLE_FALLBACK", "0") != "1":
            return None

        tmp = None
        try:
            tmp = fitz.open()
            dst = tmp.new_page(width=r.width, height=r.height)
            if r.width >= 2 and r.height >= 2:
                dst.show_pdf_page(fitz.Rect(0, 0, r.width, r.height), page.parent, page.number, clip=r)
                try:
                    tabs = dst.find_tables()
                    if tabs and getattr(tabs, "tables", None):
                        data = tabs.tables[0].extract()
                        if data and len(data) >= 2:
                            return [[(c if isinstance(c, str) else ("" if c is None else str(c))).strip()
                                     for c in row] for row in data]
                except Exception:
                    pass
        finally:
            if tmp is not None:
                tmp.close()
            try:
                fitz.TOOLS.store_shrink(100)
            except Exception:
                pass
        return None
