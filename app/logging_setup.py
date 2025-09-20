# app/logging_setup.py
from __future__ import annotations
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import os

APP_NAME = "xtractor"

def _log_dir() -> Path:
    # ~/.xtractor/logs  (works for dev and packaged app)
    base = Path.home() / f".{APP_NAME}" / "logs"
    base.mkdir(parents=True, exist_ok=True)
    return base

def configure_logging(level: int = logging.INFO) -> logging.Logger:
    logger = logging.getLogger(APP_NAME)
    if logger.handlers:  # already configured
        return logger

    logger.setLevel(level)

    # File handler (rotating, ~1 MB × 3 files)
    fh = RotatingFileHandler(
        _log_dir() / "app.log", maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(processName)s | %(name)s:%(lineno)d | %(message)s"
    )
    fh.setFormatter(fmt)
    fh.setLevel(level)
    logger.addHandler(fh)

    # Console handler in dev runs
    if os.environ.get("XTRACTOR_CONSOLE_LOG", "0") == "1":
        ch = logging.StreamHandler()
        ch.setFormatter(fmt)
        ch.setLevel(level)
        logger.addHandler(ch)

    # quiet noisy libs (optional)
    logging.getLogger("PIL").setLevel(logging.WARNING)
    logging.getLogger("fitz").setLevel(logging.WARNING)

    return logger
