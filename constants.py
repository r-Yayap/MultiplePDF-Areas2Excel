# constants.py

VERSION_TEXT = "Version 0.240915-00"

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



'''
Version 00
- subprocessing (now x5 speed!!)

Start of v0.240915

End of v0.231219
Changelog 14
-- Overhaul - Separated into files and lessened the use of Global Variables.
-- Added header Selection

Changelog 13
- Tooltips added
- Fixed updating rectangles
- Autoscroll
- Right Click to Delete Rectangle

Changelog 12
- added Image extraction (would not work for PDFs with multiple pages)
- added Last Modified Date and Size for Extractor and other features
- updated to pymupdf v 1.24.10
- import/export areas
- recent pdf/close pdf

Changelog 11
- Fixed "Illegal Text" error
- Text sorting hotfix

Changelog 10
- Fixed OCR

Changelog 09
- Fixed Zoom
- OCR options and DPI options

Changelog 08
- Optimized-ish? XD
- Progress Bar (at last!)

Changelog 07
- Resize Display along with th windows
- Added Option button for other features
- Added Bulk Rename and Directory List
- Scroll & Shift+Scroll on Canvas

Changelog 06
- UI Overhaul
- Zoom implemented
- Rectangles stays when zoomed
- Removed progress bar
- Can now edit coordinates

Changelog 05
-  time lapsed counter
-  Progress bar during extraction
-  List files on a table (DWG and PDf counter) [INTEGRATED!!!]
-  excel output: add time created on filename
-  open generated excel file (or directory)
-  Includes all pages

Changelog 04
- added DWG and PDF counter (numbers only)
- include subfolder

Changelog 03
- coordinates based on pdf's rotation
- can now read text regardless of pdf inherent rotation

Changelog 02
- scrollbar (not placed well, but working)
- area selection now working (coordinates are now correct)
- text extraction is now working

Changelog 01
- scrollbar (not placed well, but working)
- area selection in display (areas not fixed yet)
'''
