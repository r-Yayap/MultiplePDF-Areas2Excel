# constants.py

VERSION_TEXT = "Version 0.240915-03"

# Application settings
INITIAL_WIDTH = 965
GOLDEN_RATIO = (1 + 5 ** 0.5) / 2
INITIAL_HEIGHT = INITIAL_WIDTH / GOLDEN_RATIO
CANVAS_WIDTH = int(INITIAL_WIDTH * 0.98)  # was 0.95
CANVAS_HEIGHT = int(INITIAL_HEIGHT * 0.80)  # was 0.75
CANVAS_PADDING = 20
INITIAL_X_POSITION = 100
INITIAL_Y_POSITION = 100
CURRENT_ZOOM = 2.0

# Sidebar (tab view)
SIDEBAR_WIDTH = 280
SIDEBAR_PADDING = 10  # left/right margin around the sidebar

# Canvas margins and scrollbars
CANVAS_LEFT_MARGIN = SIDEBAR_WIDTH + SIDEBAR_PADDING * 2  # sidebar + padding
CANVAS_TOP_MARGIN = 100
CANVAS_BOTTOM_MARGIN = 40
SCROLLBAR_THICKNESS = 16
CANVAS_EXTRA_MARGIN = 5  # extra spacing between canvas and scrollbar


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
    "1200": 1200
}

from standalone.sc_pdf_dwg_list import launch_pdf_dwg_gui
from standalone.sc_dir_list import generate_file_list_and_excel
from standalone.sc_bulk_rename import bulk_rename_gui
from standalone import sc_bim_file_checker

OPTION_ACTIONS = {
    "PDF/DWG List": launch_pdf_dwg_gui,
    "Directory List": generate_file_list_and_excel,
    "Bulk Renamer": bulk_rename_gui,
    "BIM File Checker": sc_bim_file_checker.main,
}


SCROLL_MARGIN = 20  # Distance from the canvas edge to start scrolling
SCROLL_INCREMENT_THRESHOLD = 3  # Adjust this for slower/faster auto-scroll
scroll_counter = 0  # This will be updated in your main code
RESIZE_DELAY = 700  # milliseconds delay

# Tesseract OCR data folder (will be set in code)
TESSDATA_FOLDER = None


