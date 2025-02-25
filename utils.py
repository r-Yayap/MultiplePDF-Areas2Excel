# utils.py

import os
from tkinter import filedialog
from CTkToolTip import *
from openpyxl.utils import get_column_letter

from constants import *
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk

class EditableTreeview(ttk.Treeview):
    def __init__(self, root_window, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.root_window = root_window

        # Class code here

        self._entry = None
        self._col = None

        # Other initialization code for your EditableTreeview
        # Bind right-click to show context menu
        self.bind("<Button-3>", self.show_context_menu)
        self.bind("<Double-Button-1>", self.on_double_click)

        # Create context menu
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Remove Row", command=self.remove_row)

    def on_double_click(self, event):
        item = self.focus()
        col = self.identify_column(event.x)
        if item and col and col != "#0":
            self._col = col
            cell_values = self.item(item, "values")
            if cell_values:
                col_index = int(col.split("#")[-1]) - 1
                cell_value = cell_values[col_index]
                self.edit_cell(item, col, cell_value)

    def show_context_menu(self, event):
        item = self.identify_row(event.y)
        if item:
            self.context_menu.post(event.x_root, event.y_root)

    def remove_row(self):
        item = self.focus()
        if item:
            # Locate the index of the rectangle to remove
            rectangle_index = self.index(item)  # Simplified lookup based on Treeview row

            # Remove the rectangle from both Treeview and canvas
            self.root_window.pdf_viewer.canvas.delete(self.root_window.pdf_viewer.rectangle_list[rectangle_index])
            del self.root_window.pdf_viewer.areas[rectangle_index]
            del self.root_window.pdf_viewer.rectangle_list[rectangle_index]

            # Remove from Treeview and update display
            self.delete(item)
            self.root_window.update_areas_treeview()
            self.root_window.pdf_viewer.update_rectangles()

            print("Removed rectangle and updated canvas.")

    def edit_cell(self, item, col, _):
        def on_ok():
            new_value = entry_var.get()
            if new_value:
                current_values = list(self.item(item, "values"))
                current_values[col_index] = new_value
                self.item(item, values=tuple(current_values))
                self.update_areas_list()  # Update areas list when cell is edited
            top.destroy()

        bbox = self.bbox(item, col)
        x, y, _, _ = bbox
        col_index = int(col.replace("#", "")) - 1  # Subtract 1 for 0-based indexing

        # Create the top-level window without specifying a parent
        top = ctk.CTkToplevel()
        top.title("Edit Cell")

        entry_var = ctk.StringVar()
        entry_var.set(self.item(item, "values")[col_index])

        entry = ctk.CTkEntry(top, justify="center", textvariable=entry_var,
                             width=100, height=20, font=(BUTTON_FONT, 9),
                             border_width=1, corner_radius=3)

        entry.pack(pady=5)

        ok_button = ctk.CTkButton(top, text="OK", command=on_ok)
        ok_button.pack()

        top.geometry(f"+{x}+{y}")
        top.grab_set()  # Make the pop-up modal

        entry.focus_set()
        top.wait_window(top)  # Wait for the window to be closed

    def on_focus_out(self, _event):
        if self._entry is not None:
            self.stop_editing()

    def stop_editing(self, event=None):
        if self._entry is not None:
            new_value = self._entry.get()
            item = self.focus()

            if event and getattr(event, "keysym", "") == "Return" and item:
                current_values = self.item(item, "values")
                updated_values = [new_value if i == 0 else val for i, val in enumerate(current_values)]
                self.item(item, values=updated_values)
                self.update_areas_list()  # Update areas list when cell is edited

            self._entry.destroy()
            self._entry = None
            self._col = None

    def update_areas_list(self):
        """Updates the areas in the main application when a Treeview cell is edited."""
        updated_areas = []
        for row_id in self.get_children():
            values = self.item(row_id, "values")
            title, x0, y0, x1, y1 = values
            updated_areas.append({
                "title": title,
                "coordinates": [float(x0), float(y0), float(x1), float(y1)]
            })

        # Access pdf_viewer through main_app, not root_window
        self.root_window.pdf_viewer.areas = updated_areas
        self.root_window.pdf_viewer.update_rectangles()  # Refresh rectangles on the canvas


def create_tooltip(widget, message,
                   delay=0.3,
                   font=("Verdana", 9),
                   border_width=1,
                   border_color="gray50",
                   corner_radius=6,
                   justify="left"):
    return CTkToolTip(widget,
                      delay=delay,
                      justify=justify,
                      font=font,
                      border_width=border_width,
                      border_color=border_color,
                      corner_radius=corner_radius,
                      message=message)



def find_tessdata():
    """
    Searches for the Tesseract OCR 'tessdata' folder in multiple locations.
    Priority:
    1. The folder beside the Xtractor app (same directory as the script/executable).
    2. Common installation locations: Program Files, Local AppData, Roaming AppData.
    3. Manual selection if not found.
    """

    global tessdata_folder

    # Get the directory where Xtractor is running
    app_directory = os.path.dirname(os.path.abspath(__file__))
    xtractor_tessdata = os.path.join(app_directory, "tessdata")

    # List of locations to check (prioritized)
    locations = [
        xtractor_tessdata,  # 🔹 1st: Look for 'tessdata' beside Xtractor
        os.path.join("C:", os.sep, "Program Files", "Tesseract-OCR", "tessdata"),  # 2nd: Program Files
        os.path.join(os.getenv("LOCALAPPDATA"), "Programs", "Tesseract-OCR", "tessdata"),  # 3rd: Local AppData
        os.path.join(os.getenv("APPDATA"), "Tesseract-OCR", "tessdata"),  # 4th: Roaming AppData
    ]

    # Check each location
    for path in locations:
        print(f"Checking: {path}")  # Debugging: Print the paths being checked
        if os.path.exists(path):
            print(f"✅ Found tessdata at: {path}")  # Debugging: Print where tessdata is found
            tessdata_folder = path
            os.environ["TESSDATA_PREFIX"] = path
            return tessdata_folder

    # If not found, prompt user to manually select 'tessdata' folder
    manual_path = filedialog.askdirectory(title="Select Tesseract TESSDATA folder manually")
    if os.path.exists(manual_path):
        print(f"✅ Manually selected tessdata at: {manual_path}")  # Debugging
        tessdata_folder = manual_path
        os.environ["TESSDATA_PREFIX"] = manual_path
        return tessdata_folder

    print("❌ Tessdata folder not found.")
    return None

def adjust_coordinates_for_rotation(coordinates, rotation, pdf_height, pdf_width):
    """
    Adjusts the given coordinates based on the rotation of a PDF page.

    Args:
        coordinates (list): The original coordinates [x0, y0, x1, y1].
        rotation (int): The rotation angle of the page (0, 90, 180, 270 degrees).
        pdf_height (int): The height of the PDF page.
        pdf_width (int): The width of the PDF page.

    Returns:
        list: Adjusted coordinates based on the specified rotation.
    """
    if rotation == 0:
        return coordinates
    elif rotation == 90:
        x0, y0, x1, y1 = coordinates
        return [y0, pdf_width - x1, y1, pdf_width - x0]
    elif rotation == 180:
        x0, y0, x1, y1 = coordinates
        return [pdf_width - x1, pdf_height - y1, pdf_width - x0, pdf_height - y0]
    elif rotation == 270:
        x0, y0, x1, y1 = coordinates
        return [pdf_height - y1, x0, pdf_height - y0, x1]
    else:
        raise ValueError("Invalid rotation angle. Must be 0, 90, 180, or 270 degrees.")


def clean_text(text):
    # Replace illegal characters with a placeholder
    cleaned_text = ''.join(char if char.isprintable() else '�' for char in text)
    return cleaned_text.strip()

def get_cell_dimensions(sheet, cell):
    col_letter = get_column_letter(cell.column)
    col_width = sheet.column_dimensions[col_letter].width or 8.43  # Default width if not set
    row_height = sheet.row_dimensions[cell.row].height or 15  # Default height if not set
    return col_width * 7, row_height  # Approximate width in pixels

