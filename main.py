# main.py
"""
Xtractor Application - A modular PDF and data extraction tool.

This application is organized into separate modules to ensure maintainability, scalability, and a clear separation of concerns.
The following provides an overview of each module and how they work together:

1. **main.py**
   - Contains the main application class `XtractorApp`.
   - Initializes the application window, creates the main GUI instance, and starts the Tkinter main loop.
   - Acts as the entry point for the Xtractor Application.

2. **gui.py**
   - Contains the `XtractorGUI` class, which constructs the graphical user interface (GUI) for the application.
   - Manages window widgets like buttons, entry fields, and checkboxes.
   - Responds to user actions, triggering functions such as file opening, PDF area selection, or text extraction.
   - Interfaces with `pdf_viewer.py` to display PDFs and with `extractor.py` for text extraction functionalities.

3. **pdf_viewer.py**
   - Houses the `PDFViewer` class, which handles PDF display, zooming, panning, and area selection functionalities.
   - Collaborates with `gui.py` to enable users to define regions on a PDF for targeted text extraction.
   - Stores coordinates of user-defined areas and forwards this information to `extractor.py` for text extraction.

4. **extractor.py**
   - Contains the `TextExtractor` class, which manages text extraction based on user-selected areas.
   - Extracts text within coordinates specified by `PDFViewer` in `pdf_viewer.py`.
   - Supports OCR (Optical Character Recognition) based on user-defined settings.
   - Exports extracted data to an Excel file, adding file metadata such as size and last modified date.

5. **utils.py**
   - Provides reusable utilities and custom widgets, including coordinate handling and tooltip functionality.
   - Contains `EditableTreeview` for enhanced interaction with area attributes, like updating titles or coordinates.
   - Centralizes reusable functions, improving code clarity and reducing redundancy across modules.

6. **constants.py**
   - Stores constants for configuration settings used throughout the application.
   - Includes window dimensions, fonts, DPI values, and other configuration parameters.
   - Allows easy adjustments to global settings, keeping configuration centralized.

---

### Module Interactions

- `main.py` initializes the `XtractorApp` class, loading GUI components in `gui.py` and starting the main loop.
- `XtractorGUI` in `gui.py` generates instances of `PDFViewer` and `TextExtractor` for displaying PDFs and extracting data.
- `PDFViewer` (in `pdf_viewer.py`) displays PDFs and manages selected areas, passing coordinates to `TextExtractor` (in `extractor.py`) for text extraction.
- `utils.py` provides shared utilities, including enhanced widgets for GUI interactivity, and `constants.py` centralizes configuration.

---

This modular setup enables easy maintainability, allowing new features or changes with minimal impact on other modules.
"""
import multiprocessing

import customtkinter as ctk
from gui import XtractorGUI,CTkDnD
from constants import INITIAL_WIDTH, INITIAL_HEIGHT, INITIAL_X_POSITION, INITIAL_Y_POSITION, VERSION_TEXT

class XtractorApp:
    def __init__(self):
        self.root = CTkDnD()
        self.root.title("Xtractor " + VERSION_TEXT + " --FINAL")
        self.root.geometry(f"{INITIAL_WIDTH}x{INITIAL_HEIGHT}+{INITIAL_X_POSITION}+{INITIAL_Y_POSITION}")
        self.gui = XtractorGUI(self.root)

    def run(self):
        self.root.mainloop()

def main():
    app = XtractorApp()
    app.run()

if __name__ == '__main__':
    multiprocessing.freeze_support()  # This helps PyInstaller handle multiprocessing.
    main()

