# constants.py

VERSION_TEXT = "Version 0.240915-01" #01

# Application settings
INITIAL_WIDTH = 965
GOLDEN_RATIO = (1 + 5 ** 0.5) / 2
INITIAL_HEIGHT = INITIAL_WIDTH / GOLDEN_RATIO
CANVAS_WIDTH = int(INITIAL_WIDTH * 0.95)  # 95% of window width
CANVAS_HEIGHT = int(INITIAL_HEIGHT * 0.75)  # 75% of window height
INITIAL_X_POSITION = 100
INITIAL_Y_POSITION = 100
CURRENT_ZOOM = 2.0

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

from sc_pdf_dwg_list import pdf_dwg_counter
from sc_dir_list import generate_file_list_and_excel
from sc_merger import MergerGUI
from sc_bulk_rename import bulk_rename_gui

OPTION_ACTIONS = {
    "PDF/DWG List": pdf_dwg_counter,
    "Directory List": generate_file_list_and_excel,
    "Bulk Renamer": bulk_rename_gui
}


SCROLL_MARGIN = 20  # Distance from the canvas edge to start scrolling
SCROLL_INCREMENT_THRESHOLD = 3  # Adjust this for slower/faster auto-scroll
scroll_counter = 0  # This will be updated in your main code
RESIZE_DELAY = 700  # milliseconds delay

# Tesseract OCR data folder (will be set in code)
TESSDATA_FOLDER = None



'''
Version 02
- fixed duplicate Title issue
- revised creation of temp folder for images.

Version 01
- re-added error handling for folders/textboxes


Version 00
- subprocessing (now x5 speed!!)
- Progress bar added again
- List of Drawings merger added on other features

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
