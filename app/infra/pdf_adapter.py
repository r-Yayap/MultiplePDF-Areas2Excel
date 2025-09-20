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
        Table finder for revision tables.
        - Try on-page first (cheap).
        - If that fails and the page is rotated (90/270), copy the *clipped region*
          to a tiny temp page and draw it with a compensating rotation so the table
          is upright; then run find_tables() there.
        - Optional heavy fallback (mini-doc without rotation) is still controlled by
          REV_TABLE_FALLBACK=1.
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

        # 1) fast path: on original page
        try:
            tabs = page.find_tables(clip=r)
            if tabs and getattr(tabs, "tables", None):
                data = tabs.tables[0].extract()
                if data and len(data) >= 2:
                    return [[(c if isinstance(c, str) else ("" if c is None else str(c))).strip()
                             for c in row] for row in data]
        except Exception:
            pass

        # 2) rotation-normalized fallback for 90/270 pages
        try:
            rotation = getattr(page, "rotation", 0) % 360
        except Exception:
            rotation = 0

        if rotation in (90, 270):
            tmp = None
            try:
                # swap width/height for 90/270 to avoid clipping after rotate
                dst_w, dst_h = (r.height, r.width)
                tmp = fitz.open()
                dst = tmp.new_page(width=dst_w, height=dst_h)

                # draw the clipped region, compensating the page rotation so the content is upright
                # rotate expects multiples of 90; use (360-rotation) to deskew
                dst.show_pdf_page(
                    fitz.Rect(0, 0, dst_w, dst_h),
                    page.parent, page.number,
                    clip=r,
                    rotate=(360 - rotation) % 360
                )
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
                # aggressively release MuPDF caches used during show_pdf_page
                try:
                    fitz.TOOLS.store_shrink(100)
                except Exception:
                    pass

        # 3) optional heavy fallback without rotation (kept behind env flag)
        if os.getenv("REV_TABLE_FALLBACK", "0") == "1":
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
