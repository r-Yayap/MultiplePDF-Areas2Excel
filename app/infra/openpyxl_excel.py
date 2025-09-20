from __future__ import annotations
from openpyxl import Workbook
from typing import List
from app.ports.excel_port import ExcelPort
from pathlib import Path
from app.errors import ExcelWriteError, CancelledError

import os, tempfile, shutil

class OpenpyxlExcel(ExcelPort):
    def __init__(self):
        self._wb = None
        self._ws = None
        self._path = None

    def open(self, output_path: str, headers: List[str]) -> None:
        self._path = output_path
        self._wb = Workbook()
        self._ws = self._wb.active
        self._ws.title = "Extraction"
        self._ws.append(headers)

    def append_row(self, row: List[str]) -> None:
        self._ws.append(row)

    def save_and_close(self) -> None:
        self._wb.save(self._path)
        self._wb.close()
        self._wb = self._ws = self._path = None

    def save_rows(self, rows, target_path: str, should_cancel) -> str:
        tmp_dir = Path(target_path).parent
        tmp_dir.mkdir(parents=True, exist_ok=True)

        # create a temp file in the same directory, so the final rename is atomic on Windows
        fd, tmp_path = tempfile.mkstemp(prefix="xtractor_", suffix=".xlsx", dir=tmp_dir)
        os.close(fd)

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Rectangles"
            ws.append(["Title", "x0", "y0", "x1", "y1"])
            for r in rows:
                if should_cancel():
                    raise CancelledError("User cancelled during Excel write.")
                ws.append(r)

            wb.save(tmp_path)
            wb.close()

            # atomic-ish replace (Windows: fall back to unlink + move)
            final = Path(target_path)
            try:
                os.replace(tmp_path, final)        # preferred
            except Exception:
                if final.exists():
                    final.unlink()
                shutil.move(tmp_path, final)

            return str(final)

        except CancelledError:
            # remove temp on cancel; propagate
            try: os.unlink(tmp_path)
            except Exception: pass
            raise
        except Exception as e:
            # remove temp on failure
            try: os.unlink(tmp_path)
            except Exception: pass
            raise ExcelWriteError(str(e)) from e
