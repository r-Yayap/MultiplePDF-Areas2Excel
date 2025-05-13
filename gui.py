# gui.py
import multiprocessing
import os
import time
import customtkinter as ctk
from tkinter import filedialog, messagebox, StringVar
from openpyxl import Workbook, load_workbook
from constants import *
from extractor import TextExtractor

from pdf_viewer import PDFViewer
from utils import create_tooltip, EditableTreeview
from utils import find_tessdata
from utils import REVISION_PATTERNS



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
        export_file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Rectangles As"
        )
        if not export_file_path:
            return

        try:
            wb = Workbook()
            ws_area = wb.active
            ws_area.title = "Rectangles"
            ws_area.append(["Title", "x0", "y0", "x1", "y1"])
            for area in self.pdf_viewer.areas:
                ws_area.append([area["title"]] + area["coordinates"])

            if self.pdf_viewer.revision_area:
                ws_rev = wb.create_sheet("RevisionTable")
                ws_rev.append(["Title", "x0", "y0", "x1", "y1"])
                rev = self.pdf_viewer.revision_area
                ws_rev.append([rev["title"]] + rev["coordinates"])

            wb.save(export_file_path)
            wb.close()
            print(f"Exported areas to {export_file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export areas: {e}")

    def import_rectangles(self):
        import_file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsm"), ("All files", "*.*")],
            title="Import Rectangles"
        )
        if not import_file_path:
            return

        try:
            wb = load_workbook(import_file_path)
            ws_area = wb["Rectangles"] if "Rectangles" in wb.sheetnames else wb.active

            self.pdf_viewer.areas = []
            for row in ws_area.iter_rows(min_row=2, values_only=True):
                title, x0, y0, x1, y1 = row
                self.pdf_viewer.areas.append({"title": title, "coordinates": [x0, y0, x1, y1]})

            # Handle revision area
            revision_area_set = False
            if "RevisionTable" in wb.sheetnames:
                ws_rev = wb["RevisionTable"]
                for row in ws_rev.iter_rows(min_row=2, values_only=True):
                    title, x0, y0, x1, y1 = row
                    # ✅ Only set revision area if all coordinates are present
                    if all(isinstance(coord, (int, float)) for coord in [x0, y0, x1, y1]):
                        self.pdf_viewer.revision_area = {"title": title, "coordinates": [x0, y0, x1, y1]}
                        revision_area_set = True

            if not revision_area_set:
                self.pdf_viewer.revision_area = None  # ✅ Clear revision area if nothing valid was set

            self.pdf_viewer.update_rectangles()
            self.update_areas_treeview()
            print(f"Imported areas from {import_file_path}")
        except Exception as e:
            messagebox.showerror("Import Error", f"Could not import areas: {e}")

    def clear_all_areas(self):
        """Clears all areas and updates the display."""
        self.pdf_viewer.clear_areas()  # Clear area rectangles
        self.pdf_viewer.revision_area = None  # ✅ Clear revision area rectangle
        self.areas_tree.delete(*self.areas_tree.get_children())  # Clear all entries in the Treeview
        self.pdf_viewer.update_rectangles()  # ✅ Ensure canvas refresh reflects the change
        print("All areas and revision table cleared.")

    def update_areas_treeview(self):
        """Updates the Treeview with only area mode rectangles (excludes revision)."""
        self.areas_tree.delete(*self.areas_tree.get_children())
        self.treeview_item_ids = {}

        for index, area in enumerate(self.pdf_viewer.areas):
            coordinates = area["coordinates"]
            title = area["title"]
            item_id = self.areas_tree.insert("", "end", values=(title, *coordinates))
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
        self.pdf_folder_button = ctk.CTkButton(self.root, text="📂", command=self.browse_pdf_folder,
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

        # Mode Toggle Buttons
        self.mode_area_btn = ctk.CTkButton(self.root, text="🟥 Area Mode", width=85, height=10,
                                           font=(BUTTON_FONT, 9), command=self.set_mode_area)
        self.mode_area_btn.place(x=740, y=15) #set to negative to hide

        self.mode_revision_btn = ctk.CTkButton(self.root, text="🟩 Revision Mode", width=85, height=10,
                                               font=(BUTTON_FONT, 9), command=self.set_mode_revision)
        self.mode_revision_btn.place(x=740, y=40) #set to negative to hide

        # Create preview dropdown options
        pattern_options = [f"{k} — {', '.join(v['examples'])}" for k, v in REVISION_PATTERNS.items()]
        self.revision_dropdown_map = {f"{k} — {', '.join(v['examples'])}": k for k, v in REVISION_PATTERNS.items()}

        self.revision_pattern_var = StringVar(value=pattern_options[-1])  # Default: Fuzzy Search
        self.revision_pattern_menu = ctk.CTkOptionMenu(
            self.root,
            font=("Verdana", 8),
            values=pattern_options,
            variable=self.revision_pattern_var,
            width=200,
            height=18,
            fg_color="#5A6C89",  # Main fill
            button_color="#5A6C89",  # Button (caret) background
            text_color="white"  # Text color (works in dark mode)
        )

        self.revision_pattern_menu.place(x=740, y=65)  #set to negative to hide
        create_tooltip(self.revision_pattern_menu, "Choose the revision format pattern")

        # Zoom Slider
        self.zoom_var = ctk.DoubleVar(value=self.pdf_viewer.current_zoom)  # Initialize with the current zoom level
        self.zoom_slider = ctk.CTkSlider(self.root, from_=0.1, to=3.5, variable=self.zoom_var,
                                         command=self.update_zoom, width=170)

        # Create a frame to hold the zoom slider and labels
        self.zoom_frame = ctk.CTkFrame(self.root, fg_color="transparent")

        self.zoom_out_label = ctk.CTkLabel(self.zoom_frame, text="➖", font=(BUTTON_FONT, 14))

        self.zoom_slider = ctk.CTkSlider(self.zoom_frame, from_=0.1, to=4, variable=self.zoom_var,
                                         command=self.update_zoom, width=170)

        self.zoom_in_label = ctk.CTkLabel(self.zoom_frame, text="➕", font=(BUTTON_FONT, 14))


        # Open Sample PDF Button
        self.open_sample_button = ctk.CTkButton(self.root, text="📑 Select", command=self.open_sample_pdf,
                                                font=(BUTTON_FONT, 9),
                                                width=25, height=10)
        self.open_sample_button.place(x=20, y=35)

        # Recent PDF Button
        self.recent_pdf_button = ctk.CTkButton(self.root, text="↩ Recent", command=self.open_recent_pdf,
                                               font=(BUTTON_FONT, 9), width=28, height=10)
        self.recent_pdf_button.place(x=85, y=35)

        # Close PDF Button
        self.close_pdf_button = ctk.CTkButton(self.root, text="❌", command=self.close_pdf,anchor="center",
                                              font=(BUTTON_FONT, 9), width=10, height=10, fg_color="red4")
        self.close_pdf_button.place(x=149, y=35)

        # Output Excel Path
        self.output_path_entry = ctk.CTkEntry(self.root, width=270, height=20, font=(BUTTON_FONT, 9),
                                              placeholder_text="Select Folder for Excel output",
                                              border_width=1, corner_radius=3)
        self.output_path_entry.place(x=50, y=60)
        self.output_path_button = ctk.CTkButton(self.root, text="📂", command=self.browse_output_path,
                                                font=(BUTTON_FONT, 9), fg_color="#217346",
                                                width=25, height=10)
        self.output_path_button.place(x=20, y=60)

        # Include Subfolders Checkbox
        self.include_subfolders_var = ctk.IntVar()
        self.include_subfolders_checkbox = ctk.CTkCheckBox(self.root, text="Include Subfolders?",
                                                           variable=self.include_subfolders_var,border_width=1,
                                                           command=self.toggle_include_subfolders,
                                                           font=(BUTTON_FONT, 9),checkbox_width=17, checkbox_height=17)
        self.include_subfolders_checkbox.place(x=192, y=34)

        # Extract Button
        self.extract_button = ctk.CTkButton(self.root, text="Extract", font=("Arial Black", 14),
                                            corner_radius=10, width=75, height=30, command=self.start_extraction, anchor="center")
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

        # DPI Option Menu
        self.dpi_var = ctk.IntVar(value=150)
        self.dpi_menu = ctk.CTkOptionMenu(self.root, values=["50","75", "150", "300", "450", "600"],
                                          command=self.dpi_callback, font=("Verdana Bold", 7),
                                          variable=self.dpi_var, width=43, height=14)
        self.dpi_menu.place(x=372, y=30)
        self.dpi_label = ctk.CTkLabel(self.root, text="DPI:", text_color="gray59", font=("Verdana Bold", 8),bg_color="transparent",height=5)
        self.dpi_label.place(x=348, y=32)

        # Import, Export, and Clear Areas Buttons
        self.import_button = ctk.CTkButton(self.root, text="⬇️ Import", command=self.import_rectangles,
                                           font=(BUTTON_FONT, 9), width=60, height=10)
        self.import_button.place(x=670, y=15)

        self.export_button = ctk.CTkButton(self.root, text="⬆️ Export", command=self.export_rectangles,
                                           font=(BUTTON_FONT, 9), width=60, height=10)
        self.export_button.place(x=670, y=40)

        self.clear_areas_button = ctk.CTkButton(self.root, text="🗑 Clear", command=self.clear_all_areas,
                                                font=(BUTTON_FONT, 9), width=60, height=10, fg_color = "red4")
        self.clear_areas_button.place(x=670, y=65)

        # Option Menu for Other Features
        self.optionmenu_var = StringVar(value="Other Features")
        self.optionmenu = ctk.CTkOptionMenu(self.root, values=list(OPTION_ACTIONS.keys()),
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

    def set_mode_area(self):
        self.pdf_viewer.selection_mode = "area"
        print("🟥 Switched to Area Mode")
        self.highlight_mode_button(self.mode_area_btn)
        self.reset_mode_button(self.mode_revision_btn)

    def set_mode_revision(self):
        self.pdf_viewer.selection_mode = "revision"
        print("🟩 Switched to Revision History Mode")
        self.highlight_mode_button(self.mode_revision_btn)
        self.reset_mode_button(self.mode_area_btn)

    def highlight_mode_button(self, button):
        """Visually highlight the active mode button."""
        button.configure(fg_color="#3949AB", text_color="white", border_color="white", border_width=2)

    def reset_mode_button(self, button):
        """Reset the visual style of inactive buttons."""
        button.configure(fg_color="gray25", text_color="white", border_width=0)

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

    def dpi_callback(self, dpi_value):
        self.ocr_settings['dpi_value'] = int(dpi_value)
        print(f"DPI set to: {dpi_value}")

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
        """Initiates the extraction process with a progress bar and total files count."""

        # Check if areas are defined
        if not self.pdf_viewer.areas:
            messagebox.showerror("Extraction Error", "No areas defined. Please select areas before extracting.")
            return

        # Validate PDF folder path
        self.pdf_folder = self.pdf_folder_entry.get()
        if not self.pdf_folder or not os.path.isdir(self.pdf_folder):
            messagebox.showerror("Invalid Folder",
                                 "The specified PDF folder does not exist. Please select a valid folder.")
            return

        # Validate Excel output path
        self.output_excel_path = self.output_path_entry.get()
        output_dir = os.path.dirname(self.output_excel_path)  # Extract folder path from full file path
        if not self.output_excel_path or not os.path.isdir(output_dir):
            messagebox.showerror("Invalid Output Path","The specified output path is invalid. Please select a valid folder.")
            return

        # Close any open PDF before extraction
        self.pdf_viewer.close_pdf()

        # Record start time
        self.start_time = time.time()

        # Create the progress window
        self.progress_window = ctk.CTkToplevel(self.root)
        self.progress_window.title("Progress")
        self.progress_window.geometry("300x120")

        # Make the progress window stay on top
        self.progress_window.transient(self.root)  # Set as a child of the root window
        self.progress_window.grab_set()  # Modal window
        self.progress_window.attributes('-topmost', True)  # Keep it on top

        # Add a progress label
        self.progress_label = ctk.CTkLabel(self.progress_window, text="Processing PDFs...")
        self.progress_label.pack(pady=5)

        # Add a total files label
        self.total_files_label = ctk.CTkLabel(self.progress_window, text="Total files: 0")
        self.total_files_label.pack(pady=5)

        # Add a progress bar to the window
        self.progress_var = ctk.DoubleVar(value=0)
        self.progress_bar = ctk.CTkProgressBar(self.progress_window, variable=self.progress_var,
                                               orientation="horizontal", width=250)
        self.progress_bar.pack(pady=10)

        # Set up shared counters (use multiprocessing.Value for direct shared memory access)
        progress_counter = multiprocessing.Value('i', 0)  # ✅ REAL shared memory
        total_files = multiprocessing.Value('i', 0)  # ✅ REAL shared memory

        # Use manager only for shared strings
        manager = multiprocessing.Manager()
        final_output_path = manager.Value("s", "")  # ✅ OK as proxy, we don’t call get_lock on this

        # Start extraction in a new Process
        # Get selected revision pattern key from dropdown
        selected_pattern_display = self.revision_pattern_var.get()
        selected_pattern_key = self.revision_dropdown_map[selected_pattern_display]
        selected_revision_regex = REVISION_PATTERNS[selected_pattern_key]["pattern"]

        extractor = TextExtractor(
            pdf_folder=self.pdf_folder,
            output_excel_path=self.output_excel_path,
            areas=self.pdf_viewer.areas,
            ocr_settings=self.ocr_settings,
            include_subfolders=self.include_subfolders,
            revision_regex=selected_revision_regex
        )

        # ✅ Pass the revision area from the viewer to the extractor
        extractor.revision_area = self.pdf_viewer.revision_area

        extraction_process = multiprocessing.Process(target=extractor.start_extraction,
                                                     args=(progress_counter, total_files, final_output_path))
        extraction_process.start()

        # ✅ Store the reference to fetch the filename later
        self.final_output_path = final_output_path

        # Monitor progress
        self.root.after(100, self.update_progress, progress_counter, total_files, extraction_process)

    def update_progress(self, progress_counter, total_files, extraction_process):
        """Updates the progress bar and total files count during extraction."""

        if total_files.value > 0:
            processed_files = progress_counter.value
            self.total_files_label.configure(text=f"Processed: {processed_files}/{total_files.value}")
            current_progress = processed_files / total_files.value
            self.progress_var.set(current_progress)

        # Check if the extraction process is alive
        if extraction_process.is_alive():
            self.root.after(100, self.update_progress, progress_counter, total_files, extraction_process)
        else:
            # ✅ Ensure the progress bar is complete and close the progress window
            self.progress_var.set(1)
            self.progress_window.destroy()

            # ✅ Show completion message
            end_time = time.time()
            elapsed_time = end_time - self.start_time
            formatted_time = time.strftime("%H:%M:%S", time.gmtime(elapsed_time))

            response = messagebox.askyesno(
                "Extraction Complete",
                f"PDF extraction completed successfully in {formatted_time}.\nWould you like to open the Excel file?"
            )

            # ✅ Ensure the extraction process is finished before accessing filename
            extraction_process.join()

            # ✅ Fetch the final output file name from multiprocessing.Value
            final_output_file = self.final_output_path.value.strip()

            if final_output_file:
                print(f"DEBUG: Found final_output_path -> {final_output_file}")
            else:
                print("DEBUG: No final_output_path found. Using default -> {self.output_excel_path}")
                final_output_file = self.output_excel_path  # Fallback

            # ✅ Open the correct output file
            if response and final_output_file and os.path.exists(final_output_file):
                try:
                    print(f"DEBUG: Opening {final_output_file}...")
                    os.startfile(final_output_file)
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open the Excel file: {e}")

    def optionmenu_callback(self, choice):
        """Execute the corresponding function based on the selected option."""
        action = OPTION_ACTIONS.get(choice)
        if action:
            action()  # Call the function
        else:
            messagebox.showerror("Error", f"No action found for {choice}")

    def on_window_resize(self, event=None):
        """Handles window resizing, ensuring only the canvas and scrollbars resize while keeping all buttons fixed."""
        new_width = self.root.winfo_width()
        new_height = self.root.winfo_height()

        if hasattr(self, "prev_width") and hasattr(self, "prev_height"):
            if new_width == self.prev_width and new_height == self.prev_height:
                return  # Skip unnecessary updates

        self.prev_width = new_width
        self.prev_height = new_height

        try:
            # ✅ Resize only the canvas
            canvas_width = new_width   # Adjust to fit within window width
            canvas_height = max(200, new_height - 160)  # Ensure it keeps a reasonable height

            self.pdf_viewer.resize_canvas(canvas_width, canvas_height)

            # ✅ Adjust scrollbars to match the resized canvas
            self.pdf_viewer.v_scrollbar.configure(height=canvas_height)
            self.pdf_viewer.h_scrollbar.configure(width=canvas_width-25)
            self.pdf_viewer.v_scrollbar.place_configure(x=canvas_width - self.pdf_viewer.v_scrollbar.winfo_width(), y=100)
            self.pdf_viewer.h_scrollbar.place_configure(x=10, y=canvas_height + 107)

            # ✅ Move version label **only horizontally**, not vertically
            self.version_label.place_configure(x=new_width - 140, y=canvas_height + 127)

            self.zoom_frame.place_configure(x=30, y=canvas_height + 127)
            self.zoom_out_label.pack(side="left", padx=(5, 2))
            self.zoom_slider.pack(side="left")
            self.zoom_in_label.pack(side="left", padx=(2, 5))



        except Exception as e:
            print(f"Error resizing widgets: {e}")

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



