# constants.py

VERSION_TEXT = "1.3-test"

# ───────── Geometry (base DIP values @ 96 DPI) ─────────
INITIAL_WIDTH = 965
GOLDEN_RATIO = (1 + 5 ** 0.5) / 2
INITIAL_HEIGHT = int(INITIAL_WIDTH / GOLDEN_RATIO)  # <-- was float; Tk wants ints

CANVAS_WIDTH = int(INITIAL_WIDTH * 0.98)
CANVAS_HEIGHT = int(INITIAL_HEIGHT * 0.98)

CANVAS_PADDING = 20
INITIAL_X_POSITION = 100
INITIAL_Y_POSITION = 100
CURRENT_ZOOM = 2.0

# Sidebar (tab view)
SIDEBAR_WIDTH = 280
SIDEBAR_PADDING = 10  # left/right margin around the sidebar

# Canvas margins and scrollbars (base DIP; scale in GUI with px())
CANVAS_LEFT_MARGIN = SIDEBAR_WIDTH + SIDEBAR_PADDING * 2
CANVAS_TOP_MARGIN = 15
CANVAS_BOTTOM_MARGIN = 25
SCROLLBAR_THICKNESS = 16
CANVAS_EXTRA_MARGIN = 5

BUTTON_FONT = "Verdana"

DPI_OPTIONS = {
    "50": 50,
    "75": 75,
    "150": 150,
    "300": 300,
    "450": 450,
    "600": 600,
    "750": 750,
    "900": 900,
    "1200": 1200,
}

# ───────── Lazy-load tool actions to avoid heavy imports at import-time ─────────
import importlib


def _lazy(module: str, attr: str):
    def _call(*args, **kwargs):
        mod = importlib.import_module(module)
        return getattr(mod, attr)(*args, **kwargs)

    return _call


launch_pdf_dwg_gui = _lazy("standalone.sc_pdf_dwg_list", "launch_pdf_dwg_gui")
generate_file_list_and_excel = _lazy("standalone.sc_dir_list", "generate_file_list_and_excel")
bulk_rename_gui = _lazy("standalone.sc_bulk_rename", "bulk_rename_gui")
bim_checker_main = _lazy("standalone.sc_bim_file_checker", "main")

OPTION_ACTIONS = {
    "PDF/DWG List": launch_pdf_dwg_gui,
    "Directory List": generate_file_list_and_excel,
    "Bulk Renamer": bulk_rename_gui,
    "BIM File Checker": bim_checker_main,
}

# Auto-scroll behavior for rectangle drawing (base DIP; scale in GUI if needed)
SCROLL_MARGIN = 20
SCROLL_INCREMENT_THRESHOLD = 3
# Avoid mutable state in constants; keep scroll counters inside the widget classes.
# (If you still reference `scroll_counter` today, migrate it into PDFViewer and then delete this.)
scroll_counter = 0

RESIZE_DELAY = 700  # ms

# Tesseract OCR data folder (resolved at runtime)
TESSDATA_FOLDER = None

tool_definitions = {
    "📐 PDF & DWG Checker": {
        "action": launch_pdf_dwg_gui,
        "instructions": """
1. Select the folder containing PDF & DWG files.

2. Optionally, if DWG files are in another folder, uncheck “Same folder” and select the folder where DWG files are located.

3. It will generate an Excel report comparing filenames:
   • Matching filenames will be on the same row
   • Missing = empty
   • Duplicates will be shown on the last columns
   • Includes file sizes, modified dates, and relative folders."""
    },
    "🧮 BIM File Checker": {
        "action": bim_checker_main,
        "instructions": """
The BIM File Checker scans a folder and visually shows in Excel:
    -Which file types are present (RVT, IFC, DWG, etc.)
    -Which ones are missing
    -Which files are duplicated
It creates a clear table so you can quickly spot what's complete and what is missing.

How to use it?

1. Select the directory containing your BIM-related files (RVT, IFC, NWD, etc.).

2. Once the folder is selected, save the output somewhere.

3. It generates an Excel report highlighting:
   • ✓ Green    - File type found
   • Black      - Empty (missing) types
   • Red        - Duplicate file names for the same type
"""
    },
    "✏️ Bulk Rename Tool": {
        "action": bulk_rename_gui,
        "instructions": """
1. Load a mapping file (.csv or Excel) with original and new filenames.
    The csv or excel file should follow this:
    - First row will be treated as a header and will not be included in renaming
    - Original Filename should be on the FIRST column
    - New Filename should be on the second column
    - Do not forget the file format for both column!

2. Select the root folder where files are located.

4. Click “Start Rename” to apply the changes.

• Errors (e.g. files not found or rename failed) will be listed and copied to your clipboard."""
    },
    "📊 Folder File Exporter": {
        "action": generate_file_list_and_excel,
        "instructions": """
1. Select a folder to scan.

2. Select where to save the output.

2. It will lists all files with:
   • File name, size, extension, modified date, folder name
   • Also contains hyperlinks in the filename column."""
    },
}
