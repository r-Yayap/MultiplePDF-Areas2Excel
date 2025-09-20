import multiprocessing as mp
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

from app.domain.models import ExtractionRequest
from extractor import TextExtractor  # your existing implementation

@dataclass
class ExtractionJob:
    process: mp.Process
    total_files: mp.Value
    progress: mp.Value
    final_output_path: "mp.managers.ValueProxy[str]"
    cancel_event: mp.Event
    manager: mp.Manager
    started_at: float

class ExtractController:
    """
    UI-agnostic orchestration:
    - start(req) returns a job handle
    - poll(job) -> (processed, total) or None when finished
    - finish(job) -> output path (or None if cancelled)
    - cancel(job)
    """

    def start(self, req: ExtractionRequest) -> ExtractionJob:
        progress_counter = mp.Value('i', 0)
        total_files = mp.Value('i', 0)
        cancel_event = mp.Event()
        manager = mp.Manager()
        final_output_path = manager.Value("s", "")

        extractor = TextExtractor(
            pdf_folder=str(Path(req.pdf_paths[0]).parent),   # keeps your current constructor happy
            output_excel_path=str(req.output_excel),
            areas=[{"title": a.title, "coordinates": list(a.rect)} for a in req.areas],
            ocr_settings={"enable_ocr": req.ocr.mode, "dpi_value": req.ocr.dpi, "tessdata_folder": str(req.ocr.tessdata_dir) if req.ocr.tessdata_dir else None},
            revision_regex=req.revision_regex,
        )
        extractor.revision_area = {"title": req.revision_area.title, "coordinates": list(req.revision_area.rect)} if req.revision_area else None

        # IMPORTANT: keep your current target signature
        p = mp.Process(
            target=extractor.start_extraction,
            args=(progress_counter, total_files, final_output_path, [str(p) for p in req.pdf_paths], cancel_event),
            daemon=False
        )
        p.start()

        return ExtractionJob(
            process=p,
            total_files=total_files,
            progress=progress_counter,
            final_output_path=final_output_path,
            cancel_event=cancel_event,
            manager=manager,
            started_at=time.time(),
        )

    def poll(self, job: ExtractionJob) -> Optional[tuple[int, int]]:
        if job.process.is_alive():
            return (job.progress.value, job.total_files.value)
        return None

    def finish(self, job: ExtractionJob) -> Optional[Path]:
        try:
            job.process.join(timeout=2)
            job.process.close()
        except Exception:
            pass
        out = (job.final_output_path.value or "").strip()
        try:
            job.manager.shutdown()
        except Exception:
            pass
        return Path(out) if out else None

    def cancel(self, job: ExtractionJob) -> None:
        try:
            job.cancel_event.set()
        except Exception:
            pass
