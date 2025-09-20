from __future__ import annotations
import multiprocessing as mp, time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Iterable

from app.domain.models import ExtractionRequest
from app.services.extraction_service import ExtractionService
from app.infra.pdf_adapter import PdfAdapter


@dataclass
class ExtractionJob:
    process: mp.Process
    total_pages: mp.Value
    progress: mp.Value
    cancel_event: mp.Event
    manager: mp.Manager
    final_output_path: "mp.managers.ValueProxy[str]"
    started_at: float

def _run_service(
    req: ExtractionRequest,
    progress: mp.Value,
    total: mp.Value,
    final_output_path: "mp.managers.ValueProxy[str]",
    cancel_evt: mp.Event
):
    svc = ExtractionService()
    pdf = PdfAdapter()

    try:
        total.value = sum(pdf.page_count(p) for p in req.pdf_paths)
    except Exception:
        total.value = 0

    def on_progress(proc: int, tot: int):
        progress.value = proc

    def should_cancel() -> bool:
        return cancel_evt.is_set()

    out = svc.extract(req, on_progress=on_progress, should_cancel=should_cancel)
    final_output_path.value = str(out)

class ExtractController:
    def start(self, req: ExtractionRequest) -> ExtractionJob:
        progress = mp.Value('i', 0)
        total = mp.Value('i', 0)
        cancel_evt = mp.Event()
        manager = mp.Manager()
        final_out = manager.Value("s", "")

        p = mp.Process(
            target=_run_service,
            args=(req, progress, total, final_out, cancel_evt),
            daemon=False
        )
        p.start()

        return ExtractionJob(
            process=p,
            total_pages=total,
            progress=progress,
            cancel_event=cancel_evt,
            manager=manager,
            final_output_path=final_out,
            started_at=time.time(),
        )

    def poll(self, job: ExtractionJob) -> Optional[tuple[int, int]]:
        if job.process.is_alive():
            return (job.progress.value, job.total_pages.value)
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
