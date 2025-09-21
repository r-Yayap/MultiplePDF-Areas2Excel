# app/logging_setup.py
from __future__ import annotations
import logging, os, sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

APP_NAME = "xtractor"

def _log_dir() -> Path:
    base = Path(os.getenv("XTRACTOR_LOG_DIR", Path.home() / f".{APP_NAME}" / "logs"))
    base.mkdir(parents=True, exist_ok=True)
    return base

def configure_logging(level: int | None = None) -> logging.Logger:
    # Resolve level from env or default
    level_name = (os.getenv("XTRACTOR_LOG_LEVEL") or "").upper()
    lvl = getattr(logging, level_name, None) if level_name else None
    lvl = lvl or level or logging.INFO

    root = logging.getLogger()
    # Idempotency: if we've already added our handler, bail
    if any(getattr(h, "name", "") == "xtractor_file" for h in root.handlers):
        return logging.getLogger(APP_NAME)

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(processName)s | %(name)s:%(lineno)d | %(message)s"
    )

    fh = RotatingFileHandler(
        _log_dir() / "app.log",
        maxBytes=1_000_000, backupCount=3, encoding="utf-8", delay=True
    )
    fh.set_name("xtractor_file")
    fh.setFormatter(fmt)
    fh.setLevel(lvl)
    root.addHandler(fh)

    if os.environ.get("XTRACTOR_CONSOLE_LOG", "0") == "1":
        ch = logging.StreamHandler()
        ch.set_name("xtractor_console")
        ch.setFormatter(fmt)
        ch.setLevel(lvl)
        root.addHandler(ch)

    root.setLevel(lvl)
    logging.captureWarnings(True)

    # Tame noisy libs
    logging.getLogger("PIL").setLevel(logging.WARNING)
    logging.getLogger("fitz").setLevel(logging.WARNING)

    return logging.getLogger(APP_NAME)

def install_excepthook():
    def _hook(exc_type, exc, tb):
        logging.getLogger(APP_NAME).exception("Uncaught exception", exc_info=(exc_type, exc, tb))
    sys.excepthook = _hook
