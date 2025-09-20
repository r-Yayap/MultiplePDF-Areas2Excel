from __future__ import annotations
import json, os, shutil, csv
from pathlib import Path
from typing import Dict, List
from datetime import datetime
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.cell import WriteOnlyCell

BASE_HEADERS = ["Size (Bytes)", "Date Last Modified", "Folder", "Filename", "Page No", "Page Size"]

def _max_revisions(csv_path: Path) -> int:
    m = 0
    with open(csv_path, "r", encoding="utf-8") as f:
        r = csv.reader(f)
        next(r, None)
        for row in r:
            rev_str = row[-1] if row else "[]"
            try:
                revs = json.loads(rev_str) if rev_str.strip().startswith("[") else []
            except json.JSONDecodeError:
                revs = []
            m = max(m, len(revs))
    return m

def write_from_csv(
    combined_csv: Path,
    out_path: Path,
    temp_image_folder: Path,
    unique_headers_mapping: Dict[int, str],
    needs_images: bool,
    pdf_root: Path
) -> Path:
    area_headers = [unique_headers_mapping[i] for i in range(len(unique_headers_mapping))]
    extra_headers = ["Latest Revision", "Latest Description", "Latest Date"]
    max_revisions = _max_revisions(combined_csv)
    revision_headers = [f"Rev{i+1}" for i in range(max_revisions)]
    headers = ["UNID"] + BASE_HEADERS + area_headers + extra_headers + revision_headers

    wb = Workbook(write_only=not needs_images)
    ws = wb.create_sheet("Sheet1") if wb.write_only else wb.active
    ws.append(headers)

    filename_col_idx0 = headers.index("Filename")
    folder_col_idx0 = headers.index("Folder")
    area_first_idx0 = 1 + len(BASE_HEADERS)
    ocr_font = Font(color="FF3300")

    with open(combined_csv, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader, None)
        for row in reader:
            unid = row[0]
            base = row[1:7]
            areas = row[7:7 + len(area_headers)]
            base_index = 7 + len(area_headers)
            latest_rev  = row[base_index]     if len(row) > base_index else ""
            latest_desc = row[base_index + 1] if len(row) > base_index + 1 else ""
            latest_date = row[base_index + 2] if len(row) > base_index + 2 else ""

            # rev json (last col)
            try:
                revisions = json.loads(row[-1]) if row and row[-1].strip().startswith("[") else []
            except json.JSONDecodeError:
                revisions = []

            flat = []
            for it in revisions:
                if isinstance(it, dict):
                    r = (it.get("rev") or "").strip()
                    d = (it.get("desc") or "").strip()
                    dt = (it.get("date") or "").strip()
                    flat.append(" | ".join(x for x in (r, d, dt) if x))
                else:
                    flat.append("" if it is None else str(it))

            padded = (flat[:max_revisions] + [""] * max(0, max_revisions - len(flat)))
            row_values = [unid] + base + areas + [latest_rev, latest_desc, latest_date] + padded

            if needs_images:
                ws.append(row_values)
                r = ws.max_row

                folder_val = row_values[folder_col_idx0]
                filename_val = row_values[filename_col_idx0]
                if folder_val and filename_val:
                    abs_path = os.path.abspath(pdf_root / folder_val / filename_val)
                    cell = ws.cell(row=r, column=filename_col_idx0 + 1)
                    cell.hyperlink = abs_path
                    cell.font = Font(color="0000FF", underline="single")

                for i in range(len(area_headers)):
                    col = area_first_idx0 + i
                    cell = ws.cell(row=r, column=col + 1)
                    if isinstance(cell.value, str) and "_OCR_" in cell.value:
                        cell.font = ocr_font
                        cell.value = cell.value.replace("_OCR_", "").strip()

                    try:
                        page_no = row_values[headers.index("Page No")]
                        image_filename = f"{filename_val}_page{page_no}_area{i}.png"
                        img_path = temp_image_folder / image_filename
                        if img_path.exists():
                            img = ExcelImage(str(img_path))
                            img.anchor = f"{get_column_letter(col + 1)}{r}"
                            ws.add_image(img)
                    except Exception:
                        pass

            else:
                row_cells = []
                folder_val = row_values[folder_col_idx0]
                filename_val = row_values[filename_col_idx0]
                abs_path = os.path.abspath(pdf_root / folder_val / filename_val) if folder_val and filename_val else None

                for idx0, val in enumerate(row_values):
                    c = WriteOnlyCell(ws, value=val)
                    if idx0 == filename_col_idx0 and abs_path:
                        c.font = Font(color="0000FF", underline="single")
                        c.hyperlink = abs_path
                    if area_first_idx0 <= idx0 < area_first_idx0 + len(area_headers):
                        if isinstance(val, str) and "_OCR_" in val:
                            c.value = val.replace("_OCR_", "").strip()
                            c.font = ocr_font
                    row_cells.append(c)
                ws.append(row_cells)

    # versioned filename if already exists
    final_out = Path(out_path)
    if final_out.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = final_out.with_suffix("")
        final_out = final_out.with_name(f"{stem.name}_{ts}{final_out.suffix}")

    wb.save(str(final_out))
    return final_out

def copy_ndjson(temp_stream_path: Path, excel_out: Path):
    ndjson_out = excel_out.with_name(excel_out.stem + "_revisions.ndjson")
    try:
        if temp_stream_path.exists():
            shutil.copy(str(temp_stream_path), str(ndjson_out))
    except Exception as e:
        print(f"❌ Failed to copy NDJSON: {e}")
