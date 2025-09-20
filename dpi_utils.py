# dpi_utils.py
import platform

def init_windows_dpi_awareness():
    """Call before creating any Tk/CTk window."""
    if platform.system() != "Windows":
        return
    try:
        import ctypes
        # Prefer Per-Monitor-V2, fall back to Per-Monitor-V1
        try:
            ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)  # PER_MONITOR_AWARE_V2
        except Exception:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)          # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        pass

def get_dpi_for_window(hwnd: int) -> int:
    """Return current window DPI (Windows only)."""
    import ctypes
    user32 = ctypes.windll.user32
    try:
        return user32.GetDpiForWindow(hwnd)  # requires Windows 10+
    except Exception:
        # Fallback to system DPI
        hdc = user32.GetDC(0)
        LOGPIXELSX = 88
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, LOGPIXELSX)
        user32.ReleaseDC(0, hdc)
        return dpi

def apply_scaling(root, preferred_ui=1.0):
    """
    Sync Tk + CustomTkinter scaling with the monitor DPI.
    Call after root is created and whenever DPI changes.
    preferred_ui lets you nudge overall size (e.g., 1.1 for slightly bigger UI).
    Returns (dpi, scale_96).
    """
    import customtkinter as ctk

    hwnd = root.winfo_id()
    dpi = get_dpi_for_window(hwnd) if hwnd else 96
    scale_96 = dpi / 96.0
    # Tk expects pixels per typographic point (1/72")
    root.tk.call('tk', 'scaling', dpi / 72.0)

    # CustomTkinter has two knobs — keep them in sync with Windows DPI
    ctk.set_widget_scaling(scale_96 * preferred_ui)
    ctk.set_window_scaling(scale_96 * preferred_ui)
    return dpi, scale_96

def install_dpi_watcher(root, callback, poll_ms=500):
    """
    Polls the window DPI; if it changes (e.g., user drags window to another monitor),
    re-apply scaling and call `callback(dpi, scale_96)`.
    """
    from functools import lru_cache
    @lru_cache(maxsize=1)
    def _get():
        from .dpi_utils import get_dpi_for_window  # safe self-import
        return get_dpi_for_window(root.winfo_id())

    state = {"last_dpi": None}

    def _tick():
        try:
            dpi_now = get_dpi_for_window(root.winfo_id())
            if dpi_now and dpi_now != state["last_dpi"]:
                state["last_dpi"] = dpi_now
                callback(dpi_now, dpi_now/96.0)
        finally:
            root.after(poll_ms, _tick)

    _tick()
