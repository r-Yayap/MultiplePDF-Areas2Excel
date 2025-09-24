# app/logging_setup.py
from __future__ import annotations
import logging, os, sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

APP_NAME = "xtractor"
APP_NAME = "xtractor"

def _app_dir() -> Path:
    # When frozen (PyInstaller/Nuitka), write next to the executable;
    # otherwise write in the project root (parent of the 'app' package).
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]

def _log_dir() -> Path:
    # Allow optional override via env var, else use "<app>/logs"
    base = Path(os.getenv("XTRACTOR_LOG_DIR") or (_app_dir() / "logs"))
    base.mkdir(parents=True, exist_ok=True)
    return base

def log_file_path() -> Path:
    return _log_dir() / "app.log"

def configure_logging(level: int | None = None) -> logging.Logger:
    # Resolve level from env or default
    level_name = (os.getenv("XTRACTOR_LOG_LEVEL") or "").upper()
    lvl = getattr(logging, level_name, None) if level_name else None
    lvl = lvl or level or logging.INFO

    root = logging.getLogger()
    # Idempotent: if our file handler already exists, don’t add again
    if any(getattr(h, "name", "") == "xtractor_file" for h in root.handlers):
        return logging.getLogger(APP_NAME)

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(processName)s | %(name)s:%(lineno)d | %(message)s"
    )

    fh = RotatingFileHandler(log_file_path(), maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    fh.setFormatter(fmt)
    fh.name = "xtractor_file"

    root.addHandler(fh)
    root.setLevel(lvl)

    return logging.getLogger(APP_NAME)

