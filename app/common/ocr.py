import os
from tkinter import filedialog

def find_tessdata():
    """
    Find the Tesseract 'tessdata' folder; prompt if not found.
    """
    app_directory = os.path.dirname(os.path.abspath(__file__))
    # When packaged, tessdata may live beside the app root; go two levels up from /app/common
    xtractor_tessdata = os.path.join(os.path.dirname(os.path.dirname(app_directory)), "tessdata")

    locations = [
        xtractor_tessdata,
        os.path.join("C:", os.sep, "Program Files", "Tesseract-OCR", "tessdata"),
        os.path.join(os.getenv("LOCALAPPDATA") or "", "Programs", "Tesseract-OCR", "tessdata"),
        os.path.join(os.getenv("APPDATA") or "", "Tesseract-OCR", "tessdata"),
    ]

    for path in locations:
        if path and os.path.exists(path):
            os.environ["TESSDATA_PREFIX"] = path
            return path

    manual_path = filedialog.askdirectory(title="Select Tesseract TESSDATA folder manually")
    if manual_path and os.path.exists(manual_path):
        os.environ["TESSDATA_PREFIX"] = manual_path
        return manual_path

    return None
