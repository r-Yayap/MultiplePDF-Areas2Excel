from pathlib import Path
import sys

def resource_path(rel: str) -> str:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)) if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[2]
    return str(base_dir / rel)
