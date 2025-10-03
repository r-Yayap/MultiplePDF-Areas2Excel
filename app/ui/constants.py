# constants.py

VERSION_TEXT = "1.3.250924"

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


CARD_MIN_W  = 200     # try 220–260 depending on how compact you want the cards
CARD_GAP    = 16      # horizontal gap between cards (matches grid padx total)
WRAPPER_PAD = 16      # inner left+right padding inside the overlay wrapper
MIN_APP_H   = 540     # a sensible minimum app height

CARD_BG_DEFAULT   = "#3a3a3a"  # unselected bg
CARD_BG_SELECTED  = "#910206"  # selected bg (pick what you like)
CARD_BORDER_DEFAULT  = "gray35"
CARD_BORDER_SELECTED = "#C10206"
CARD_BORDER_HOVER = "#8a8a8a"
CARD_BG_HOVER     = "#2f2f2f"

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

# Auto-scroll behavior for rectangle drawing (base DIP; scale in GUI if needed)
SCROLL_MARGIN = 20
SCROLL_INCREMENT_THRESHOLD = 3
# Avoid mutable state in constants; keep scroll counters inside the widget classes.
# (If you still reference `scroll_counter` today, migrate it into PDFViewer and then delete this.)
scroll_counter = 0

RESIZE_DELAY = 700  # ms

# Tesseract OCR data folder (resolved at runtime)
TESSDATA_FOLDER = None

# ───────── Tools: import directly from standalone so PyInstaller bundles them
from standalone.sc_pdf_dwg_list import open_window as pdf_dwg_list_open
from standalone.sc_dir_list import generate_file_list_and_excel as dir_list_open
from standalone.sc_bulk_rename import open_window as bulk_rename_open
from standalone.sc_bim_file_checker import open_window as bim_checker_open



tool_definitions = {
    "PDF & DWG Checker": {
        "action": pdf_dwg_list_open,
        "needs_master": True,   # opens a CTkToplevel(master)
        "blurb": "Compare PDFs and DWGs, spot missing/duplicates, export to Excel.",
        "instructions": """
1. Select the folder containing PDF & DWG files.

2. Optionally, if DWG files are in another folder, uncheck “Same folder” and select the folder where DWG files are located.

3. It will generate an Excel report comparing filenames:
   • Matching filenames will be on the same row
   • Missing = empty
   • Duplicates will be shown on the last columns
   • Includes file sizes, modified dates, and relative folders."""
    },
    "BIM File Checker": {
        "action": bim_checker_open,
        "needs_master": True,   # we’ll make it use dialogs parented to master
        "blurb": "Scan BIM types (RVT, IFC, NWC/NWD, DWG, etc.) and mark present/dup/missing.",
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
    "Bulk Rename Tool": {
        "action": bulk_rename_open,
        "needs_master": True,   # we’ll make it use CTkToplevel(master)
        "blurb": "Batch-rename files using a CSV/Excel name map.",
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
    "Folder File Exporter": {
        "action": dir_list_open,
        "needs_master": False,  # uses file dialogs only
        "blurb": "Export a folder’s file inventory (with hyperlinks) to Excel.",
        "instructions": """
1. Select a folder to scan.

2. Select where to save the output.

2. It will lists all files with:
   • File name, size, extension, modified date, folder name
   • Also contains hyperlinks in the filename column."""
    },
}
