# app/services/extraction_service.py
from __future__ import annotations
import csv, gc, json, os, secrets, shutil
from dataclasses import asdict
from pathlib import Path
from typing import Callable, Iterable, Optional, Tuple

from app.domain.models import ExtractionRequest, AreaSpec
from app.infra.pdf_adapter import PdfAdapter
from app.infra.ocr_adapter import OcrAdapter
from app.services.revision_parser import RevisionParser
from app.infra.excel_writer import write_from_csv, copy_ndjson

from app.common.geometry import adjust_coordinates_for_rotation


import pymupdf as fitz

# ===== Helpers (kept top-level for Windows pickling) =====

def _open_temp_writers(temp_dir: Path, unid_prefix: str, with_revisions: bool):
    csv_path = temp_dir / f"temp_{unid_prefix}.csv"
    csv_f = open(csv_path, "w", newline="", encoding="utf-8")
    csv_w = csv.writer(csv_f)
    # line-buffered text file (buffering=1) so it flushes per newline
    ndjson_f = open(temp_dir / f"temp_{unid_prefix}.ndjson", "w", encoding="utf-8", buffering=1) if with_revisions else None
    return csv_w, csv_f, ndjson_f

def _process_single_pdf_star(args):
    """
    Safe wrapper for pool: never raises; returns dict with only primitives.
    """
    pdf_path, req, temp_dir, prefix = args
    try:
        pages = _process_single_pdf(pdf_path, req, temp_dir, prefix)
        return {"ok": True, "pages": int(pages)}
    except Exception as e:
        # strip to a plain string so it's 100% picklable
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}

def _sanitize_clip(clip: tuple, page_rect: tuple[float, float, float, float]) -> Optional[tuple[float, float, float, float]]:
    """
    Ensure x0<x1, y0<y1, intersect with page bounds, and reject tiny/empty rects.
    Rects are in page coordinate space (after any rotation adjustment).
    """
    try:
        x0, y0, x1, y1 = map(float, clip)
        px0, py0, px1, py1 = map(float, page_rect)

        # order
        if x1 < x0: x0, x1 = x1, x0
        if y1 < y0: y0, y1 = y1, y0

        # intersect with page bounds
        x0 = max(px0, min(x0, px1))
        x1 = max(px0, min(x1, px1))
        y0 = max(py0, min(y0, py1))
        y1 = max(py0, min(y1, py1))

        # strictly positive size; reject near-zero areas
        if (x1 - x0) <= 1e-3 or (y1 - y0) <= 1e-3:
            return None
        return (x0, y0, x1, y1)
    except Exception:
        return None

def _rel_folder(pdf_path: Path, root: Path) -> str:
    try:
        return os.path.relpath(pdf_path.parent, root)
    except Exception:
        return str(pdf_path.parent)

def _infer_pdf_root(pdf_paths: Iterable[Path]) -> Path:
    parents = [Path(p).parent for p in pdf_paths]
    if not parents:
        return Path.cwd()

    try:
        common = os.path.commonpath([str(p) for p in parents])
    except Exception:
        return parents[0]

    return Path(common)

def _clean_text(text: str) -> str:
    if not text:
        return ""
    t = text.replace("\n", " ").replace("\r", " ").strip()
    # replace control chars
    import re
    t = re.sub(r'[\x00-\x1F\x7F-\x9F]', '■', t)
    return re.sub(r'\s+', ' ', t)

def _prepare_headers(areas: Iterable[AreaSpec]) -> Tuple[list[str], dict]:
    headers = ["Size (Bytes)", "Date Last Modified", "Folder", "Filename", "Page No", "Page Size"]
    unique = {}
    counts = {}
    for i, a in enumerate(areas):
        title = a.title or f"Area {i+1}"
        if title not in counts:
            counts[title] = 1
            u = title
        else:
            counts[title] += 1
            u = f"{title} ({counts[title]})"
        headers.append(u)
        unique[i] = u
    return headers, unique

def _write_temp_csv(csv_path: Path, ndjson_path: Optional[Path], rows: Iterable[list], write_revisions: bool):
    with open(csv_path, "w", newline="", encoding="utf-8") as cf:
        w = csv.writer(cf)
        # header row for combined step (matches extractor.py combine)
        # UNID + BASE + each area header are appended later; here we just stream data rows
        # The combine step will prepend headers; rows here are final data rows.
        for r in rows:
            w.writerow(r)
    if write_revisions and ndjson_path:
        with open(ndjson_path, "w", encoding="utf-8") as jf:
            for r in rows:
                unid = r[0]
                rev_json = r[-1] if r else "[]"
                try:
                    revs = json.loads(rev_json) if rev_json.strip().startswith("[") else []
                except Exception:
                    revs = []
                jf.write(json.dumps({"unid": unid, "revisions": revs}, ensure_ascii=False) + "\n")

def _process_single_pdf(pdf_path: Path, req: dict, temp_dir: Path, unid_prefix: str) -> int:
    pdf = PdfAdapter()
    ocr = OcrAdapter(req["ocr_tess"])
    parser = RevisionParser(req.get("rev_regex"))

    areas_rects: list[tuple] = list(req["areas_rects"])
    revision_rect: Optional[tuple] = req.get("rev_area_rect")
    ocr_mode = req["ocr_mode"]
    dpi = max(1, int(req["ocr_dpi"] or 150))
    scale = req.get("ocr_scale")
    pdf_root = Path(req["pdf_root"])

    # open temp writers once, write per page
    csv_w, csv_f, ndjson_f = _open_temp_writers(temp_dir, unid_prefix, with_revisions=bool(revision_rect))
    pages_written = 0

    try:
        with pdf.open(pdf_path) as doc:
            page_count = doc.page_count

            for page_no in range(page_count):
                page = doc[page_no]

                # ---- metadata (size / modtime) ----
                try:
                    st = pdf_path.stat()
                    size = st.st_size
                    last_mod = __import__("datetime").datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    size, last_mod = 0, ""

                pr = tuple(pdf.page_rect(page))  # (x0,y0,x1,y1)
                pw, ph = (pr[2]-pr[0], pr[3]-pr[1])
                rotation = getattr(page, "rotation", 0)

                # ---- areas ----
                area_texts: list[str] = []
                for idx, raw in enumerate(areas_rects):
                    # Text should use rotation-adjusted rect
                    adj = adjust_coordinates_for_rotation(raw, rotation, ph, pw)

                    pr = tuple(pdf.page_rect(page))
                    clip_text = _sanitize_clip(adj, pr)  # for get_text
                    clip_img = _sanitize_clip(raw, pr)  # for pixmap & OCR (raw like legacy)

                    text_area = ""
                    try:
                        if ocr_mode == "Default":
                            if clip_text:
                                text_area = pdf.get_text(page, clip_text)
                            if (not text_area.strip()) and clip_img:
                                # OCR on the image crop (raw coords)
                                text_area = ocr.ocr_clip_to_text(page, clip_img, dpi, scale)

                        elif ocr_mode == "OCR-All":
                            if clip_img:
                                text_area = ocr.ocr_clip_to_text(page, clip_img, dpi, scale)
                            else:
                                text_area = ""

                        elif ocr_mode == "Text1st+Image-beta":
                            # 1) text first with adjusted rect
                            if clip_text:
                                text_area = pdf.get_text(page, clip_text)

                            # 2) always save image using the raw rect (visual orientation)
                            if clip_img:
                                try:
                                    pix = pdf.render_pixmap(page, clip_img, dpi=dpi, scale=scale)
                                    out_img = temp_dir / f"{pdf_path.name}_page{page_no + 1}_area{idx}.png"
                                    pix.save(str(out_img))
                                finally:
                                    try:
                                        del pix
                                    except:
                                        pass
                                    gc.collect()
                                # 30MB guard
                                try:
                                    if out_img.stat().st_size > 30 * 1024 * 1024:
                                        out_img.unlink(missing_ok=True)
                                except Exception:
                                    pass

                            # 3) OCR fallback on the same image crop (raw rect)
                            if (not text_area.strip()) and clip_img:
                                try:
                                    text_area = ocr.ocr_clip_to_text(page, clip_img, dpi, scale)
                                except Exception:
                                    text_area = "OCR_ERROR"

                        else:
                            # Fallback mode: plain text with adjusted rect
                            if clip_text:
                                text_area = pdf.get_text(page, clip_text)
                    except Exception:
                        text_area = ""

                    area_texts.append(_clean_text(text_area) if text_area.strip() else "")
                # ---- revision table ----
                revisions = []
                if revision_rect:
                    try:
                        # do NOT call page.remove_rotation(); rotate the rect instead
                        pr2 = tuple(pdf.page_rect(page))
                        pw2, ph2 = (pr2[2] - pr2[0], pr2[3] - pr2[1])
                        rotation = getattr(page, "rotation", 0)
                        adj_rev = adjust_coordinates_for_rotation(revision_rect, rotation, ph2, pw2)
                        rclip = _sanitize_clip(adj_rev, pr2)
                        if rclip:
                            rows = pdf.find_table_rows(page, rclip)
                            if rows:
                                revisions = parser.parse_table_rows(rows)

                                # free big locals
                                try:
                                    del rows
                                except Exception:
                                    pass
                    except Exception:
                        revisions = []

                # ---- final row (write immediately) ----
                folder = _rel_folder(pdf_path, pdf_root)
                filename = pdf_path.name
                pr_final = tuple(pdf.page_rect(page))
                page_size_str = f"{pr_final[2]-pr_final[0]:.1f} x {pr_final[3]-pr_final[1]:.1f}"

                latest_rev = latest_desc = latest_date = ""
                if isinstance(revisions, list) and revisions:
                    last = revisions[-1]
                    latest_rev  = last.get("rev", "")
                    latest_desc = last.get("desc", "")
                    latest_date = last.get("date", "")

                unid = f"{unid_prefix}-{page_no+1}"
                row = [unid, size, last_mod, folder, filename, page_no+1, page_size_str] \
                      + area_texts + [latest_rev, latest_desc, latest_date, json.dumps(revisions, ensure_ascii=False)]
                csv_w.writerow(row)

                if ndjson_f is not None:
                    ndjson_f.write(json.dumps({"unid": unid, "revisions": revisions}, ensure_ascii=False) + "\n")

                pages_written += 1

                # periodic flush + GC + store shrink keeps memory flat
                if (page_no + 1) % 10 == 0:
                    try:
                        csv_f.flush()
                        if ndjson_f: ndjson_f.flush()
                    except Exception:
                        pass
                    import gc as _gc
                    _gc.collect()
                    try:
                        import pymupdf as fitz
                        fitz.TOOLS.store_shrink(100)
                    except Exception:
                        pass


    finally:
        try:
            csv_f.close()
        except Exception:
            pass
        if ndjson_f:
            try:
                ndjson_f.close()
            except Exception:
                pass



    return pages_written if pages_written > 0 else 1

# ===== Main Service =====

class ExtractionService:
    """
    Orchestrates multi-file extraction using temp CSV/NDJSON then streams to Excel.
    Mirrors original behavior, but separated into adapters.
    """
    def __init__(self):
        self.pdf = PdfAdapter()

    def extract(self, req: ExtractionRequest, on_progress: Optional[Callable[[int, int], None]] = None, should_cancel: Optional[Callable[[], bool]] = None,) -> Path:
        # temp dir under app folder (secure random suffix)
        app_dir = Path(getattr(__import__("sys"), "executable", __file__)).parent \
            if getattr(__import__("sys"), "frozen", False) else Path(__file__).parent
        base_temp = app_dir / "temp"
        base_temp.mkdir(exist_ok=True)
        temp_dir = base_temp / secrets.token_hex(8)
        temp_dir.mkdir(exist_ok=True, parents=True)

        # compute pdf_root
        pdf_paths = [Path(p) for p in req.pdf_paths]
        explicit_root = Path(req.pdf_root) if req.pdf_root else None
        pdf_root = explicit_root or _infer_pdf_root(pdf_paths)

        # headers for areas (used later in Excel writer)
        _, unique_headers = _prepare_headers(req.areas)

        # total pages (for progress bar)
        try:
            total_pages = sum(self.pdf.page_count(p) for p in pdf_paths)
        except Exception:
            total_pages = len(pdf_paths)  # fallback

        processed = 0

        areas_rects = [tuple(a.rect) for a in req.areas]
        rev_area_rect = tuple(req.revision_area.rect) if req.revision_area else None

        # extract a plain pattern string or None
        rev_pattern = None
        if req.revision_regex:
            rev_pattern = getattr(req.revision_regex, "pattern", None) \
                          or (req.revision_regex if isinstance(req.revision_regex, str) else None)

        req_dict = {
            "areas_rects": areas_rects,
            "rev_area_rect": rev_area_rect,
            "rev_regex": rev_pattern,  # <-- use clean pattern here
            "ocr_mode": req.ocr.mode,
            "ocr_dpi": req.ocr.dpi,
            "ocr_scale": req.ocr.scale,
            "ocr_tess": str(req.ocr.tessdata_dir) if req.ocr.tessdata_dir else None,
            "pdf_root": str(pdf_root),
        }

        import multiprocessing as mp
        ctx = mp.get_context("spawn")

        rev_mode = req.revision_area is not None

        # Pool sizing & worker lifetime
        if rev_mode:
            procs = max(1, os.cpu_count() - 2)  # conservative in rev mode
            maxtasks = 1  # each worker handles 1 PDF then dies (kills leaks)
            batch_size = int(os.getenv("PDF_BATCH_SIZE", "30"))
        else:
            procs = max(1, os.cpu_count())
            maxtasks = 25
            batch_size = int(os.getenv("PDF_BATCH_SIZE", "500"))


        # Build jobs once
        jobs = [(p, req_dict, temp_dir, str(10000 + i)) for i, p in enumerate(pdf_paths)]
        errors: list[str] = []

        # Helper: chunk jobs
        def _chunked(seq, n):
            for i in range(0, len(seq), n):
                yield seq[i:i + n]

        cancelled = False

        # ---- run batches ----
        for bidx, batch in enumerate(_chunked(jobs, batch_size), 1):
            if cancelled:
                break

            with ctx.Pool(processes=procs, maxtasksperchild=maxtasks) as pool:
                try:
                    for res in pool.imap_unordered(_process_single_pdf_star, batch, chunksize=1):
                        if should_cancel and should_cancel():
                            cancelled = True
                            pool.terminate()
                            break

                        if not res.get("ok"):
                            errors.append(res.get("error", "Unknown worker error"))
                            continue

                        processed += res["pages"]
                        if on_progress:
                            on_progress(processed, total_pages)
                finally:
                    try:
                        pool.close()
                    except Exception:
                        pass
                    try:
                        pool.join()
                    except Exception:
                        pass

            # Hard memory reset between batches
            try:
                fitz.TOOLS.store_shrink(100)  # trim MuPDF global store
            except Exception:
                pass
            try:
                gc.collect()  # free Python objects
            except Exception:
                pass

        # combine temp CSVs (and NDJSON)
        combined_csv = temp_dir / "streamed_output.csv"
        combined_ndjson = temp_dir / "streamed_revisions.ndjson"
        self._combine_temp_files(temp_dir, combined_csv, combined_ndjson, unique_headers)

        # write final Excel (streamed)
        needs_images = (req.ocr.mode == "Text1st+Image-beta")
        excel_out = write_from_csv(
            combined_csv, req.output_excel, temp_dir, unique_headers, needs_images, pdf_root
        )

        if errors:
            errlog = req.output_excel.with_suffix(".errors.txt")
            try:
                errlog.write_text("\n".join(errors), encoding="utf-8")
            except Exception:
                pass

        # place NDJSON copy near Excel (if revisions were enabled)
        if req.revision_area:
            copy_ndjson(combined_ndjson, excel_out)

        # cleanup temp dir
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass

        return excel_out

    def _combine_temp_files(self, temp_dir: Path, combined_csv: Path, combined_ndjson: Path, unique_headers: dict):
        # CSVs
        import glob
        temp_csvs = sorted(Path(p) for p in glob.glob(str(temp_dir / "temp_*.csv")))
        with open(combined_csv, "w", newline="", encoding="utf-8") as outf:
            w = csv.writer(outf)
            w.writerow(["UNID", "Size (Bytes)", "Date Last Modified", "Folder", "Filename", "Page No", "Page Size"]
                       + [unique_headers[i] for i in range(len(unique_headers))]
                       + ["Latest Revision", "Latest Description", "Latest Date", "__revisions__"])
            for f in temp_csvs:
                with open(f, "r", encoding="utf-8") as inf:
                    r = csv.reader(inf)
                    for row in r:
                        w.writerow(row)
                try: f.unlink()
                except Exception: pass

        # NDJSON
        import glob
        temp_ndjsons = sorted(Path(p) for p in glob.glob(str(temp_dir / "temp_*.ndjson")))
        if temp_ndjsons:
            with open(combined_ndjson, "w", encoding="utf-8") as outj:
                for f in temp_ndjsons:
                    with open(f, "r", encoding="utf-8") as inj:
                        shutil.copyfileobj(inj, outj)
                    try: f.unlink()
                    except Exception: pass
