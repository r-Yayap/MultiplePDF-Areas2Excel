# constants.py

VERSION_TEXT = "Version 0.231219-14"

# Application settings
INITIAL_WIDTH = 965
GOLDEN_RATIO = (1 + 5 ** 0.5) / 2
INITIAL_HEIGHT = INITIAL_WIDTH / GOLDEN_RATIO
CANVAS_WIDTH = INITIAL_WIDTH - 30
CANVAS_HEIGHT = INITIAL_HEIGHT - 135
INITIAL_X_POSITION = 100
INITIAL_Y_POSITION = 100
CURRENT_ZOOM = 2.0

BUTTON_FONT = "Verdana"

DPI_OPTIONS = {
    "75": 75,
    "150": 150,
    "300": 300,
    "450": 450,
    "600": 600,
    "750": 750,
    "900": 900,
    "1200": 1200
}

OPTION_ACTIONS = {
    "PDF/DWG List": "pdf_dwg_counter",
    "Directory List": "generate_file_list_and_excel",
    "Bulk Renamer": "bulk_rename_gui"
}


SCROLL_MARGIN = 20  # Distance from the canvas edge to start scrolling
SCROLL_INCREMENT_THRESHOLD = 3  # Adjust this for slower/faster auto-scroll
scroll_counter = 0  # This will be updated in your main code
RESIZE_DELAY = 700  # milliseconds delay

# Tesseract OCR data folder (will be set in code)
TESSDATA_FOLDER = None

