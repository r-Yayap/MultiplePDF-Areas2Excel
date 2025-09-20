# main.py
# ────────────────────────────────────────────────────────────
# Xtractor launcher with splash screen, DPI scaling & warm-up
# ────────────────────────────────────────────────────────────
import multiprocessing
import threading
import sys
from pathlib import Path

import customtkinter as ctk
from app.ui.gui import CTkDnD, XtractorGUI   # CTkDnD ensures tkdnd is loaded
from app.ui.constants import (
    INITIAL_WIDTH, INITIAL_HEIGHT,
    INITIAL_X_POSITION, INITIAL_Y_POSITION,
    VERSION_TEXT
)
from app.ui.dpi_utils import init_windows_dpi_awareness, apply_scaling, install_dpi_watcher


from app.logging_setup import configure_logging
logger = configure_logging()
logger.info("App starting…")

# ────────────────────────────────────────────────────────────
#  Helper: resource path (PyInstaller/Nuitka/Python-safe)
# ────────────────────────────────────────────────────────────
def resource_path(rel: str) -> str:
    if getattr(sys, "frozen", False):
        base_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        base_dir = Path(__file__).parent
    return str(base_dir / rel)

def asset(rel_name: str) -> str:
    """Return path to UI assets, working in dev and frozen builds."""
    p = Path(resource_path(f"app/ui/style/{rel_name}"))
    if p.exists():
        return str(p)
    # fallback for dev layouts that still use /style at project root
    return resource_path(f"style/{rel_name}")


# ────────────────────────────────────────────────────────────
#  Splash-screen utilities
# ────────────────────────────────────────────────────────────
def create_splash(master):
    """Returns (splash_window, cancel_animation_callable)."""
    splash = ctk.CTkToplevel(master)
    splash.overrideredirect(True)

    # Transparent-ish splash
    splash.configure(fg_color="#262626")
    try:
        splash.wm_attributes("-transparentcolor", "#262626")
    except Exception:
        pass
    try:
        splash.attributes("-alpha", 0.88)
    except Exception:
        pass

    w, h = 360, 260
    x = splash.winfo_screenwidth()  // 2 - w // 2
    y = splash.winfo_screenheight() // 2 - h // 2
    splash.geometry(f"{w}x{h}+{x}+{y}")

    # Logo
    from PIL import Image
    logo_path = asset("xtractor-logo.png")
    logo_img   = ctk.CTkImage(light_image=Image.open(logo_path), size=(120, 120))
    logo_lbl   = ctk.CTkLabel(splash, image=logo_img, text="")
    logo_lbl.pack(pady=(30, 10))
    splash.logo_ref = logo_img

    # Animated text
    lbl = ctk.CTkLabel(splash, text="loading…", text_color="white", font=("Segoe UI", 14, "bold"))
    lbl.pack()

    job_id = None
    def animate(i=0):
        nonlocal job_id
        lbl.configure(text="loading" + ". " * (i % 4))
        job_id = splash.after(120, animate, i + 1)
    animate()

    def cancel():
        if job_id:
            splash.after_cancel(job_id)

    return splash, cancel


# ────────────────────────────────────────────────────────────
#  Heavy imports in background (warm-up)
# ────────────────────────────────────────────────────────────
def warm_up(event):
    import pymupdf as fitz     # noqa: F401
    import openpyxl            # noqa: F401
    import PIL.Image           # noqa: F401
    event.set()


# ────────────────────────────────────────────────────────────
#  Build and show main UI (called after warm-up)
# ────────────────────────────────────────────────────────────
def _build_and_show(root, splash, cancel_anim):
    # Window meta
    root.title("Xtractor " + VERSION_TEXT)
    root.geometry(f"{INITIAL_WIDTH}x{INITIAL_HEIGHT}+{INITIAL_X_POSITION}+{INITIAL_Y_POSITION}")
    try:
        root.iconbitmap(asset("xtractor-logo.ico"))
    except Exception:
        pass  # non-Windows or icon missing

    app = XtractorGUI(root)

    # Re-scale + refresh when moving across monitors with different DPI
    def _on_dpi_change(new_dpi, scale_96):
        apply_scaling(root)
        try:
            app.pdf_viewer.update_display()
            app.on_window_resize()
            if hasattr(app, "update_floating_controls"):
                app.update_floating_controls()
        except Exception:
            pass

    install_dpi_watcher(root, _on_dpi_change, poll_ms=500)

    # Reveal app, hide splash
    cancel_anim()
    root.deiconify()
    root.update_idletasks()
    splash.destroy()


# ────────────────────────────────────────────────────────────
#  Program entry-point
# ────────────────────────────────────────────────────────────
def main():
    # A) DPI awareness BEFORE any Tk window is created
    init_windows_dpi_awareness()

    # B) Root (DnD-enabled) hidden at first
    root = CTkDnD()
    root.withdraw()

    # C) Apply DPI scaling AFTER root exists
    apply_scaling(root, preferred_ui=1.0)

    # D) Theme (after root is created is fine)
    ctk.set_default_color_theme(asset("xtractor-dark-red.json"))
    ctk.set_appearance_mode("dark")

    # E) Splash
    splash, cancel_anim = create_splash(root)

    # F) Background warm-up
    ready = threading.Event()
    threading.Thread(target=warm_up, args=(ready,), daemon=True).start()

    # G) Switch from splash to main once warmed
    def finish():
        _build_and_show(root, splash, cancel_anim)

    def check_ready():
        if ready.is_set():
            root.after(360, finish)      # small delay so animation ends cleanly
        else:
            root.after(120, check_ready)

    root.after(100, check_ready)
    root.mainloop()


if __name__ == "__main__":
    multiprocessing.freeze_support()  # Windows packaging safety
    main()
