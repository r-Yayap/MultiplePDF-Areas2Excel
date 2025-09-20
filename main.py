# main.py
# ────────────────────────────────────────────────────────────
# Xtractor launcher with splash screen & background warm-up
# ────────────────────────────────────────────────────────────
import multiprocessing
import threading
import sys, os
import customtkinter as ctk
from gui import CTkDnD                         #  needs gui for class definition
from constants import (
    INITIAL_WIDTH, INITIAL_HEIGHT,
    INITIAL_X_POSITION, INITIAL_Y_POSITION,
    VERSION_TEXT
)

from pathlib import Path

# ────────────────────────────────────────────────────────────
#  Helper: resource path (PyInstaller-safe)
# ────────────────────────────────────────────────────────────
def resource_path(rel: str) -> str:
    """
    Return an absolute path to a bundled resource that works for:
    • normal `python main.py` runs
    • Nuitka --standalone builds
    • PyInstaller one-file / one-dir builds
    """
    # -------- if we are inside a frozen app ---------------------------
    if getattr(sys, "frozen", False):
        # PyInstaller defines _MEIPASS; other freezers (Nuitka, cx_Freeze) don't.
        base_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        # running from source – use the directory where *this* file lives
        base_dir = Path(__file__).parent

    return str(base_dir / rel)


# Apply theme once (fast)
ctk.set_default_color_theme(resource_path("style/xtractor-dark-red.json"))
ctk.set_appearance_mode("dark")

# ────────────────────────────────────────────────────────────
#  Splash-screen utilities
# ────────────────────────────────────────────────────────────
def create_splash(master):
    """
    Returns (splash_window, cancel_animation_callable).
    """
    splash = ctk.CTkToplevel(master)
    splash.overrideredirect(True)

    # ── transparent / semi-transparent window ──────────────────────
    splash.configure(fg_color="#262626")           # dark base colour
    splash.wm_attributes("-transparentcolor", "#262626")
    splash.attributes("-alpha", 0.88)              # 0=fully transparent, 1=opaque
    # splash.attributes("-transparentcolor", "#262626")  # <-- true cut-out transparency (Windows only)
                                                      #     uncomment if you prefer a non-rectangular splash
    # ----------------------------------------------------------------

    w, h = 360, 260
    x = splash.winfo_screenwidth()  // 2 - w // 2
    y = splash.winfo_screenheight() // 2 - h // 2
    splash.geometry(f"{w}x{h}+{x}+{y}")

    # ── logo ───────────────────────────────────────────────────────
    from PIL import Image
    logo_path  = resource_path("style/xtractor-logo.png")
    logo_img   = ctk.CTkImage(light_image=Image.open(logo_path), size=(120, 120))
    logo_lbl   = ctk.CTkLabel(splash, image=logo_img, text="")
    logo_lbl.pack(pady=(30, 10))
    splash.logo_ref = logo_img      # keep a reference so it isn't GC-d

    # ── animated text ──────────────────────────────────────────────
    lbl = ctk.CTkLabel(
        splash,
        text="loading…",
        text_color="white",
        font=("Segoe UI", 14, "bold")
    )
    lbl.pack()

    job_id = None

    def animate(i=0):
        nonlocal job_id
        lbl.configure(text="loading" + ". " * (i % 4))
        job_id = splash.after(120, animate, i + 1)  # ← was 300

    animate()

    def cancel():
        if job_id:
            splash.after_cancel(job_id)

    return splash, cancel


# ────────────────────────────────────────────────────────────
#  Heavy imports in background
# ────────────────────────────────────────────────────────────
def warm_up(event):
    import pymupdf as fitz
    import openpyxl, PIL.Image          # heavy libs
    event.set()

# ────────────────────────────────────────────────────────────
#  Build the main GUI (runs after warm-up)
# ────────────────────────────────────────────────────────────
def build_gui(root):
    from gui import XtractorGUI               # already imported earlier, but OK
    root.title("Xtractor " + VERSION_TEXT)
    root.geometry(f"{INITIAL_WIDTH}x{INITIAL_HEIGHT}+"
                  f"{INITIAL_X_POSITION}+{INITIAL_Y_POSITION}")
    root.iconbitmap(resource_path("style/xtractor-logo.ico"))
    XtractorGUI(root)                         # build the interface


# ────────────────────────────────────────────────────────────
#  Program entry-point
# ────────────────────────────────────────────────────────────
def main():
    # 0️⃣  Root (hidden at first)
    root = CTkDnD()
    root.withdraw()

    # 1️⃣  Splash
    splash, cancel_anim = create_splash(root)

    # 2️⃣  Background warm-up
    ready = threading.Event()
    threading.Thread(target=warm_up, args=(ready,), daemon=True).start()

    # 3️⃣  Poll until warm-up done, then switch windows
    # 3️⃣  Poll until warm-up done, then switch windows
    def finish():
        cancel_anim()               # stop animation
        build_gui(root)             # build the heavy UI (user still sees splash)
        root.deiconify()            # show finished window
        root.update_idletasks()
        splash.destroy()            # now hide splash – no blank gap

    def check_ready():
        if ready.is_set():
            # wait long enough for two animation ticks (2 × 40 ms = 80 ms)
            # a little extra margin so you reliably see ".." and "..."
            root.after(360, finish)
        else:
            root.after(120, check_ready)     # poll at the same rate as animation


    root.after(100, check_ready)            # start polling

    # 4️⃣  Main event loop
    root.mainloop()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
