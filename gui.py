# gui.py

import json
import customtkinter as ctk
from tkinter import filedialog, messagebox, StringVar
from constants import *
from utils import create_tooltip, EditableTreeview
from pdf_viewer import PDFViewer
from extractor import TextExtractor
from sc_pdf_dwg_list import pdf_dwg_counter
from sc_dir_list import generate_file_list_and_excel
from sc_bulk_rename import bulk_rename_gui
from utils import find_tessdata


class XtractorGUI:
    def __init__(self, root):
        self.root = root
        self.pdf_viewer = PDFViewer(self, self.root)  # Pass GUI instance and root window

        self.include_subfolders = False
        self.pdf_folder = ''
        self.output_excel_path = ''
        self.ocr_settings = {'enable_ocr': 'Off', 'dpi_value': 150, 'tessdata_folder': TESSDATA_FOLDER}
        self.recent_pdf_path = None

        self.setup_widgets()
        self.setup_bindings()
        self.setup_tooltips()

    def export_rectangles(self):
        """Exports the currently selected areas (rectangles) to a JSON file."""
        export_file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Save Rectangles As"
        )
        if export_file_path:
            try:
                with open(export_file_path, 'w', encoding='utf-8') as json_file:
                    json.dump(self.pdf_viewer.areas, json_file, indent=4)
                print(f"Exported areas to {export_file_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Could not export areas: {e}")

    def import_rectangles(self):
        """Imports area selections from a JSON file."""
        import_file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Import Rectangles"
        )
        if import_file_path:
            try:
                with open(import_file_path, 'r') as json_file:
                    imported_areas = json.load(json_file)
                self.pdf_viewer.areas = imported_areas  # Update the areas in the PDF viewer
                self.pdf_viewer.update_rectangles()  # Refresh the rectangles on the canvas
                self.update_areas_treeview()  # Refresh the Treeview to show imported areas
                print(f"Imported areas from {import_file_path}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Could not import areas: {e}")

    def clear_all_areas(self):
        """Clears all areas and updates the display."""
        self.pdf_viewer.clear_areas()  # Clear areas from the PDF viewer
        self.areas_tree.delete(*self.areas_tree.get_children())  # Clear all entries in the Treeview
        print("All areas cleared.")

    def update_areas_treeview(self):
        """Refreshes the Treeview to display the current areas and their titles."""

        # Clear existing entries
        self.areas_tree.delete(*self.areas_tree.get_children())

        # Insert each area with its title and coordinates into the Treeview and keep track of each item ID
        self.treeview_item_ids = {}  # Dictionary to track Treeview item IDs to canvas rectangle IDs
        for index, area in enumerate(self.pdf_viewer.areas):
            coordinates = area["coordinates"]
            title = area["title"]
            # Insert row into the Treeview and get its item ID
            item_id = self.areas_tree.insert("", "end", values=(title, *coordinates))
            # Store the item ID associated with the canvas rectangle index
            self.treeview_item_ids[item_id] = index

    def open_sample_pdf(self):
        # Opens a file dialog to select a PDF file, then displays it in the PDFViewer
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if pdf_path:
            self.pdf_viewer.display_pdf(pdf_path)
            self.recent_pdf_path = pdf_path  # Store the recent PDF path
            print(f"Opened sample PDF: {pdf_path}")

    def open_recent_pdf(self):
        """Opens the most recently viewed PDF."""
        if self.recent_pdf_path:
            self.pdf_viewer.display_pdf(self.recent_pdf_path)
            print(f"Reopened recent PDF: {self.recent_pdf_path}")
        else:
            messagebox.showinfo("Info", "No recent PDF found.")

    def close_pdf(self):
        """Delegates PDF closing to the PDF viewer."""
        self.pdf_viewer.close_pdf()

    def remove_row(self):
        """Removes the selected row from the Treeview and updates the canvas to remove the associated rectangle."""
        selected_item = self.areas_tree.selection()
        if selected_item:
            # Get the rectangle index associated with the selected Treeview item
            index = self.treeview_item_ids.get(selected_item[0])
            if index is not None:
                # Remove the rectangle from the canvas
                rectangle_id = self.pdf_viewer.rectangle_list[index]
                self.pdf_viewer.canvas.delete(rectangle_id)
                # Remove the area from PDFViewer's areas and rectangle_list
                del self.pdf_viewer.areas[index]
                del self.pdf_viewer.rectangle_list[index]

                # Remove the item from Treeview
                self.areas_tree.delete(selected_item[0])

                # Update Treeview and canvas display
                self.update_areas_treeview()
                self.pdf_viewer.update_rectangles()

                print("Removed rectangle at index", index)

    def setup_widgets(self):
        # PDF Folder Entry
        self.pdf_folder_entry = ctk.CTkEntry(self.root, width=270, height=20, font=(BUTTON_FONT, 9),
                                             placeholder_text="Select Folder with PDFs", border_width=1,
                                             corner_radius=3)
        self.pdf_folder_entry.place(x=50, y=10)
        self.pdf_folder_button = ctk.CTkButton(self.root, text="...", command=self.browse_pdf_folder,
                                               font=(BUTTON_FONT, 9),
                                               width=25, height=10)
        self.pdf_folder_button.place(x=20, y=10)

        # OCR Option Menu
        self.ocr_menu_var = StringVar(value="OCR-Off")
        self.ocr_menu = ctk.CTkOptionMenu(
            self.root,
            values=["Off", "Text-first", "OCR-All", "Text1st+Image-beta"],
            command=self.ocr_menu_callback,  # Use self.ocr_menu_callback here
            font=("Verdana Bold", 9),
            variable=self.ocr_menu_var,
            width=85,
            height=18
        )

        self.ocr_menu.place(x=330, y=10)

        # DPI Option Menu
        self.dpi_var = ctk.IntVar(value=150)
        self.dpi_menu = ctk.CTkOptionMenu(self.root, values=["75", "150", "300", "450", "600"],
                                          command=self.dpi_callback, font=("Verdana Bold", 7),
                                          variable=self.dpi_var, width=43, height=14)
        self.dpi_menu.place(x=372, y=30)
        self.dpi_label = ctk.CTkLabel(self.root, text="DPI:", text_color="gray59", font=("Verdana Bold", 8))
        self.dpi_label.place(x=348, y=32)

        # Zoom Slider
        self.zoom_var = ctk.DoubleVar(value=self.pdf_viewer.current_zoom)  # Initialize with the current zoom level
        self.zoom_slider = ctk.CTkSlider(self.root, from_=0.1, to=3.5, variable=self.zoom_var,
                                         command=self.update_zoom, width=170)
        self.zoom_slider.place(x=780, y=70)

        # Open Sample PDF Button
        self.open_sample_button = ctk.CTkButton(self.root, text="Open PDF", command=self.open_sample_pdf,
                                                font=(BUTTON_FONT, 9),
                                                width=25, height=10)
        self.open_sample_button.place(x=20, y=35)

        # Recent PDF Button
        self.recent_pdf_button = ctk.CTkButton(self.root, text="Recent PDF", command=self.open_recent_pdf,
                                               font=(BUTTON_FONT, 9), width=88, height=10)
        self.recent_pdf_button.place(x=87, y=35)

        # Close PDF Button
        self.close_pdf_button = ctk.CTkButton(self.root, text="Close PDF", command=self.close_pdf,
                                              font=(BUTTON_FONT, 9), width=88, height=10)
        self.close_pdf_button.place(x=143, y=35)

        # Output Excel Path
        self.output_path_entry = ctk.CTkEntry(self.root, width=270, height=20, font=(BUTTON_FONT, 9),
                                              placeholder_text="Select Folder for Excel output",
                                              border_width=1, corner_radius=3)
        self.output_path_entry.place(x=50, y=60)
        self.output_path_button = ctk.CTkButton(self.root, text="...", command=self.browse_output_path,
                                                font=(BUTTON_FONT, 9),
                                                width=25, height=10)
        self.output_path_button.place(x=20, y=60)

        # Include Subfolders Checkbox
        self.include_subfolders_var = ctk.IntVar()
        self.include_subfolders_checkbox = ctk.CTkCheckBox(self.root, text="Include Subfolders?",
                                                           variable=self.include_subfolders_var,
                                                           command=self.toggle_include_subfolders,
                                                           font=(BUTTON_FONT, 9))
        self.include_subfolders_checkbox.place(x=196, y=34)

        # Extract Button
        self.extract_button = ctk.CTkButton(self.root, text="EXTRACT", font=("Arial Black", 12),
                                            corner_radius=10, width=75, height=30, command=self.start_extraction)
        self.extract_button.place(x=330, y=55)

        # Areas Treeview setup
        self.areas_frame = ctk.CTkFrame(self.root, height=1, width=200, border_width=0)
        self.areas_frame.place(x=425, y=10)

        self.areas_tree = EditableTreeview(
            self,
            self.areas_frame,
            columns=("Title", "x0", "y0", "x1", "y1"),  # Ensure "Title" is included
            show="headings",
            height=3
        )

        # Set up static headers and fixed column widths
        self.areas_tree.heading("Title", text="Title")
        self.areas_tree.column("Title", width=50, anchor="center")
        for col in ("x0", "y0", "x1", "y1"):
            self.areas_tree.heading(col, text=col)
            self.areas_tree.column(col, width=45, anchor="center")

        # Pack the Treeview into the frame
        self.areas_tree.pack(side="left")

        # Import, Export, and Clear Areas Buttons
        self.import_button = ctk.CTkButton(self.root, text="Import Areas", command=self.import_rectangles,
                                           font=(BUTTON_FONT, 9), width=88, height=10)
        self.import_button.place(x=670, y=15)

        self.export_button = ctk.CTkButton(self.root, text="Export Areas", command=self.export_rectangles,
                                           font=(BUTTON_FONT, 9), width=88, height=10)
        self.export_button.place(x=670, y=40)

        self.clear_areas_button = ctk.CTkButton(self.root, text="Clear Areas", command=self.clear_all_areas,
                                                font=(BUTTON_FONT, 9), width=88, height=10)
        self.clear_areas_button.place(x=670, y=65)

        # Option Menu for Other Features
        self.optionmenu_var = StringVar(value="Other Features")
        self.optionmenu = ctk.CTkOptionMenu(self.root, values=["PDF/DWG List", "Directory List", "Bulk Renamer"],
                                            command=self.optionmenu_callback, font=(BUTTON_FONT, 9),
                                            variable=self.optionmenu_var, width=105, height=15)
        self.optionmenu.place(x=850, y=10)

        # Version Label with Tooltip
        self.version_label = ctk.CTkLabel(self.root, text=VERSION_TEXT, fg_color="transparent",
                                          text_color="gray59",
                                          font=(BUTTON_FONT, 9.5))
        self.version_label.place(x=835, y=30)
        self.version_label.bind("<Button-1>", self.display_version_info)

    def setup_bindings(self):
        self.pdf_folder_entry.bind("<KeyRelease>", self.update_pdf_folder)
        self.output_path_entry.bind("<KeyRelease>", self.update_output_path)
        self.root.bind("<Configure>", self.on_window_resize)

    def setup_tooltips(self):
        create_tooltip(self.ocr_menu, "OCR options - select an OCR mode for text extraction")
        create_tooltip(self.dpi_menu, "DPI resolution")
        create_tooltip(self.pdf_folder_entry, "Select the main folder containing PDF files")
        create_tooltip(self.open_sample_button, "Open a sample PDF to set areas")
        create_tooltip(self.output_path_entry, "Select folder for the Excel output")
        create_tooltip(self.include_subfolders_checkbox, "Include files from subfolders for extraction")
        create_tooltip(self.extract_button, "Start the extraction process")
        create_tooltip(self.import_button, "Import a saved template of selected areas")
        create_tooltip(self.export_button, "Export the selected areas as a template")
        create_tooltip(self.clear_areas_button, "Clear all selected areas")
        create_tooltip(self.optionmenu, "Select additional features")

    # OCR Menu Callback
    def ocr_menu_callback(self, choice):
        print("OCR menu dropdown clicked:", choice)

        def enable_ocr_menu(enabled):
            color = "green4" if enabled else "gray29"
            self.ocr_menu.configure(fg_color=color, button_color=color)
            self.dpi_menu.configure(state="normal" if enabled else "disabled", fg_color=color, button_color=color)

        # If OCR is "Off", don't check for tessdata and disable OCR options
        if choice == "Off":
            enable_ocr_menu(False)
            print("OCR disabled.")
            self.ocr_settings['enable_ocr'] = "Off"
            return

        # Check tessdata only for OCR modes that need it
        if choice in ("Text-first", "OCR-All", "Text1st+Image-beta"):
            found_tesseract_path = find_tessdata()
            if found_tesseract_path:
                self.ocr_settings['tessdata_folder'] = found_tesseract_path
                enable_ocr_menu(True)
                if choice == "Text-first":
                    print("OCR will start if no text is extracted.")
                elif choice == "OCR-All":
                    print("OCR will be enabled for every area.")
                elif choice == "Text1st+Image-beta":
                    print("OCR will start if no text is extracted and images will also be extracted.")
            else:
                enable_ocr_menu(False)
                print("Tessdata folder not found. OCR disabled.")

        self.ocr_settings['enable_ocr'] = choice
        print("OCR mode:", self.ocr_settings['enable_ocr'])

    # DPI Menu Callback
    def dpi_callback(self, dpi_value):
        self.ocr_settings['dpi_value'] = int(dpi_value)
        print(f"DPI set to: {dpi_value}")

    # Other methods as per previous code

    def browse_pdf_folder(self):
        self.pdf_folder = filedialog.askdirectory()
        self.pdf_folder_entry.delete(0, ctk.END)
        self.pdf_folder_entry.insert(0, self.pdf_folder)

    def browse_output_path(self):
        """Opens a dialog to specify the output Excel file path."""
        self.output_excel_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if self.output_excel_path:  # Only update the entry if a file was selected
            self.output_path_entry.delete(0, ctk.END)
            self.output_path_entry.insert(0, self.output_excel_path)

    def update_zoom_slider(self, zoom_level):
        """Updates the zoom slider to reflect the current zoom level in PDFViewer."""
        self.zoom_var.set(zoom_level)

    def update_pdf_folder(self, event):
        self.pdf_folder = self.pdf_folder_entry.get()

    def update_output_path(self, event):
        self.output_excel_path = self.output_path_entry.get()

    def update_zoom(self, value):
        """Adjusts the zoom level of the PDFViewer based on slider input."""
        zoom_level = float(value)
        self.pdf_viewer.set_zoom(zoom_level)  # Update zoom in PDFViewer

    def toggle_include_subfolders(self):
        self.include_subfolders = self.include_subfolders_var.get()

    def start_extraction(self):
        # Close any open PDF before extraction
        self.pdf_viewer.close_pdf()

        # Gather necessary data and start extraction
        extractor = TextExtractor(
            root=self.root,
            pdf_folder=self.pdf_folder,
            output_excel_path=self.output_excel_path,
            areas=self.pdf_viewer.areas,
            ocr_settings=self.ocr_settings,
            include_subfolders=self.include_subfolders
        )
        extractor.start_extraction_thread()

    def optionmenu_callback(self, choice):
        if choice == "PDF/DWG List":
            pdf_dwg_counter()
        elif choice == "Directory List":
            generate_file_list_and_excel()
        elif choice == "Bulk Renamer":
            bulk_rename_gui()

    def on_window_resize(self, event):
        """Handles window resizing and adjusts the canvas dimensions."""
        self.pdf_viewer.resize_canvas()

    def display_version_info(self, event):
        version_text = """
        Created by: Rei Raphael Reveral
        
        Links:
        https://github.com/r-Yayap/MultiplePDF-Areas2Excel
        https://www.linkedin.com/in/rei-raphael-reveral
        """
        window = ctk.CTkToplevel(self.root)
        window.title("Version Info")
        text_widget = ctk.CTkTextbox(window, wrap="word", width=400, height=247)
        text_widget.insert("end", version_text)
        text_widget.pack(padx=10, pady=10, side="left")
        window.grab_set()


