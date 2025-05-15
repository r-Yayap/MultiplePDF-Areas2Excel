# gui.py
import multiprocessing
import os
import re
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
from ttkwidgets import CheckboxTreeview
from tkinter import ttk
from tkinterdnd2 import TkinterDnD, DND_ALL



class CTkDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class XtractorGUI:
    def __init__(self, root):
        self.root = root
        self.pdf_viewer = PDFViewer(self, self.root)  # Pass GUI instance and root window

        self.root.update_idletasks()


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
        # Create a tabbed panel on the left
        self.tab_view = ctk.CTkTabview(self.root, width=SIDEBAR_WIDTH)
        self.tab_view.pack(side="left", fill="y", padx=SIDEBAR_PADDING, pady=10)

        self.root.update_idletasks()  # Ensure sidebar dimensions are accurate

        # Zoom Slider
        self.zoom_var = ctk.DoubleVar(value=self.pdf_viewer.current_zoom)  # Initialize with the current zoom level
        self.zoom_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.zoom_out_label = ctk.CTkLabel(self.zoom_frame, text="➖", font=(BUTTON_FONT, 14))
        self.zoom_slider = ctk.CTkSlider(self.zoom_frame, from_=0.1, to=4, variable=self.zoom_var,
                                         command=self.update_zoom, width=170)
        self.zoom_in_label = ctk.CTkLabel(self.zoom_frame, text="➕", font=(BUTTON_FONT, 14))

        self.zoom_out_label.pack(side="left", padx=(5, 2))
        self.zoom_slider.pack(side="left")
        self.zoom_in_label.pack(side="left", padx=(2, 5))

        # Floating recent PDF and close PDF buttons (top-right)
        self.recent_pdf_button = ctk.CTkButton(self.root, text="↩", command=self.open_recent_pdf,
                                               font=(BUTTON_FONT, 10), width=24, height=24)
        self.recent_pdf_button.pack(pady=2)

        self.close_pdf_button = ctk.CTkButton(self.root, text="⨉", command=self.close_pdf,
                                              font=(BUTTON_FONT, 10), fg_color="red4",width=24, height=24)
        self.close_pdf_button.pack(pady=2)


        # Create each tab
        tab_files = self.tab_view.add("📁 Files")
        tab_rectangles = self.tab_view.add("🔲 Rectangles")
        tab_extract = self.tab_view.add("🚀 Extract")
        tab_tools = self.tab_view.add("🧰 Tools")

        # ======================= 📁 FILES TAB =======================

        # PDF Folder Drop Zone
        self.pdf_folder_button = ctk.CTkButton(tab_files,
                                               text="\n➕\n\nDrop Folder or Click to Browse",
                                               command=self.browse_pdf_folder,
                                               fg_color="transparent", border_width=2, width=240, height=80,
                                               hover_color="#444", text_color="white")
        self.pdf_folder_button.pack(pady=(10, 5))
        self.pdf_folder_button.drop_target_register(DND_ALL)
        self.pdf_folder_button.dnd_bind('<<Drop>>', self.drop_pdf_folder)

        self.pdf_folder_entry = ctk.CTkEntry(self.root, width=240, height=24, font=(BUTTON_FONT, 9),
                                             placeholder_text="Select Folder with PDFs", border_width=1,
                                             corner_radius=3)
        self.pdf_folder_entry.place(x=-200, y=-223) #hide haha
        # Files Tree Label
        self.files_tree_label = ctk.CTkLabel(tab_files, text="📄 Files in Folder", font=(BUTTON_FONT, 10))
        self.files_tree_label.pack(pady=(10, 0))

        # Create Frame to hold the tree
        self.files_tree_frame = ctk.CTkFrame(tab_files, height=140, width=240)
        self.files_tree_frame.pack(pady=(2, 5), fill="both", expand=True)

        # 🧱 Create scrollable frame ONCE
        self.files_tree_scrollframe = ctk.CTkScrollableFrame(
            self.files_tree_frame,
            orientation="horizontal",
            width=240,
            height=150,
            fg_color="transparent"
        )
        self.files_tree_scrollframe.pack(fill="both", expand=True)

        # Inner container for Treeview
        self.files_tree_container = ctk.CTkFrame(self.files_tree_scrollframe, fg_color="transparent")
        self.files_tree_container.pack(side="left", fill="both", expand=True)

        # Placeholder for Treeview widget
        self.files_tree_widget = None
        # Counter Label (created only once)
        self.pdf_counter_label = ctk.CTkLabel(self.files_tree_frame, text="Selected PDFs: 0", font=(BUTTON_FONT, 9))
        self.pdf_counter_label.pack(pady=(5, 0), anchor="center")

        # ======================= 🔲 RECTANGLES TAB =======================


        # 🔲 Frame for Mode Buttons (Area / Revision)
        mode_frame = ctk.CTkFrame(tab_rectangles, width=240, height=24, fg_color="transparent")
        mode_frame.pack_propagate(False)
        mode_frame.pack(pady=(5, 5))

        self.mode_area_btn = ctk.CTkButton(mode_frame, text="🟥 Area", command=self.set_mode_area,
                                           font=(BUTTON_FONT, 9), width=115, height=24)
        self.mode_area_btn.pack(side="left", padx=5)

        self.mode_revision_btn = ctk.CTkButton(mode_frame, text="🟩 Revision Table", command=self.set_mode_revision,
                                               font=(BUTTON_FONT, 9), width=115, height=24)
        self.mode_revision_btn.pack(side="left", padx=5)

        #revision pattern
        pattern_options = [f"{k} — {', '.join(v['examples'])}" for k, v in REVISION_PATTERNS.items()]
        self.revision_dropdown_map = {f"{k} — {', '.join(v['examples'])}": k for k, v in REVISION_PATTERNS.items()}
        self.revision_pattern_var = StringVar(value=pattern_options[-1])
        self.revision_pattern_menu = ctk.CTkOptionMenu(tab_rectangles,
                                                       font=("Verdana", 9),
                                                       values=pattern_options,
                                                       variable=self.revision_pattern_var,
                                                       width=240, height=24,
                                                       fg_color="#5A6C89", button_color="#5A6C89", text_color="white")
        self.revision_pattern_menu.pack(pady=(5, 5))


        # ⬇️ Frame for Import/Export Buttons
        action_frame = ctk.CTkFrame(tab_rectangles, width=240, height=24, fg_color="transparent")
        action_frame.pack_propagate(False)
        action_frame.pack(pady=(25, 5))

        self.import_button = ctk.CTkButton(action_frame, text="⬇️ Import Areas", command=self.import_rectangles,
                                           font=(BUTTON_FONT, 9), width=115, height=24)
        self.import_button.pack(side="left", padx=5)

        self.export_button = ctk.CTkButton(action_frame, text="⬆️ Export Areas", command=self.export_rectangles,
                                           font=(BUTTON_FONT, 9), width=115, height=24)
        self.export_button.pack(side="left", padx=5)

        # 🗑 Clear Button (Full width)
        self.clear_areas_button = ctk.CTkButton(tab_rectangles, text="🗑 Clear All Areas", command=self.clear_all_areas,
                                                font=(BUTTON_FONT, 9), width=240, height=24, fg_color="red4")
        self.clear_areas_button.pack(pady=(5, 15))

        # Treeview inside a sub-frame
        self.areas_frame = ctk.CTkFrame(tab_rectangles, height=120, width=240)
        self.areas_frame.pack(pady=5, fill="both")
        self.areas_tree = EditableTreeview(self, self.areas_frame,
                                           columns=("Title", "x0", "y0", "x1", "y1"),
                                           show="headings", height=4)
        self.areas_tree.heading("Title", text="Title")
        self.areas_tree.column("Title", width=60, anchor="center")
        for col in ("x0", "y0", "x1", "y1"):
            self.areas_tree.heading(col, text=col)
            self.areas_tree.column(col, width=40, anchor="center")
        self.areas_tree.pack(side="left", fill="both", expand=True)

        # ======================= 🚀 EXTRACT TAB =======================

        # 📦 Frame for OCR & DPI Settings
        extract_frame = ctk.CTkFrame(tab_extract, width=240, fg_color="transparent")
        extract_frame.pack(pady=(25, 5))

        # ─────────────── OCR Setting ───────────────
        ocr_row = ctk.CTkFrame(extract_frame, fg_color="transparent")
        ocr_row.pack(pady=(0, 5), fill="x")

        ocr_label = ctk.CTkLabel(ocr_row, text="OCR Mode:", font=(BUTTON_FONT, 9), width=80, anchor="w")
        ocr_label.pack(side="left", padx=(0, 5))

        self.ocr_menu_var = StringVar(value="Off")
        self.ocr_menu = ctk.CTkOptionMenu(ocr_row,
                                          values=["Off", "Text-first", "OCR-All", "Text1st+Image-beta"],
                                          command=self.ocr_menu_callback,
                                          font=("Verdana Bold", 9),
                                          variable=self.ocr_menu_var,
                                          width=140, height=24)
        self.ocr_menu.pack(side="left")

        # ─────────────── DPI Setting ───────────────
        dpi_row = ctk.CTkFrame(extract_frame, fg_color="transparent")
        dpi_row.pack(pady=(0, 5), fill="x")

        dpi_label = ctk.CTkLabel(dpi_row, text="DPI:", font=(BUTTON_FONT, 9), width=80, anchor="w")
        dpi_label.pack(side="left", padx=(0, 5))

        self.dpi_var = ctk.IntVar(value=150)
        self.dpi_menu = ctk.CTkOptionMenu(dpi_row,
                                          values=["50", "75", "150", "300", "450", "600"],
                                          command=self.dpi_callback,
                                          font=("Verdana Bold", 9),
                                          variable=self.dpi_var,
                                          width=140, height=24)
        self.dpi_menu.pack(side="left")

        self.output_path_entry = ctk.CTkEntry(tab_extract, width=240, height=24, font=(BUTTON_FONT, 9),
                                              placeholder_text="Select Excel Output Path", border_width=1,
                                              corner_radius=3)
        self.output_path_entry.pack(pady=(10, 2))
        self.output_path_button = ctk.CTkButton(tab_extract, text="📂 Browse Output Path", command=self.browse_output_path,
                                                font=(BUTTON_FONT, 9), width=240, height=24)
        self.output_path_button.pack(pady=2)

        self.extract_button = ctk.CTkButton(tab_extract, text="🚀 Extract Now", font=("Arial Black", 13),
                                            corner_radius=10, width=240, height=30, command=self.start_extraction)
        self.extract_button.pack(pady=30)

        # ======================= 🧰 TOOLS TAB =======================
        # Create a frame for all tools
        tool_frame = ctk.CTkFrame(tab_tools)
        tool_frame.pack(pady=10, fill="both", expand=True)

        for label, tool in tool_definitions.items():
            btn = ctk.CTkButton(tool_frame, text=label, width=240, height=28, font=(BUTTON_FONT, 10),
                                command=tool["action"])
            btn.pack(pady=(10, 2))

            help_btn = ctk.CTkButton(tool_frame, text="How to use?", width=240, height=20, font=(BUTTON_FONT, 9),
                                     fg_color="gray20", hover_color="gray30",
                                     command=lambda t=tool: self.show_tool_instructions(t["instructions"]))
            help_btn.pack(pady=(0, 5))

        self.version_label = ctk.CTkLabel(tab_tools, text=VERSION_TEXT, fg_color="transparent",
                                          text_color="gray59", font=(BUTTON_FONT, 9))
        self.version_label.pack(pady=10, anchor="w")
        self.version_label.bind("<Button-1>", self.display_version_info)

        self.root.after(300, self.place_zoom_and_version_controls)

    def show_tool_instructions(self, text):
        window = ctk.CTkToplevel(self.root)
        window.title("Tool Instructions")
        window.geometry("500x400")
        text_box = ctk.CTkTextbox(window, wrap="word")
        text_box.insert("end", text)
        text_box.configure(state="disabled")
        text_box.pack(padx=10, pady=10, fill="both", expand=True)
        window.grab_set()

    def place_zoom_and_version_controls(self):
        sidebar_width = self.tab_view.winfo_width() + 20
        window_height = self.root.winfo_height()

        self.zoom_frame.place(x=sidebar_width + 10, y=window_height - 65)
        # Place recent + close buttons at top-right
        self.recent_pdf_button.place(x=sidebar_width + 0, y=23)
        self.close_pdf_button.place(x=sidebar_width + 30, y=23)

    def setup_bindings(self):
        self.pdf_folder_entry.bind("<KeyRelease>", self.update_pdf_folder)
        self.output_path_entry.bind("<KeyRelease>", self.update_output_path)
        self.root.bind("<Configure>", self.on_window_resize)

    def setup_tooltips(self):
        create_tooltip(self.ocr_menu, "OCR options - select an OCR mode for text extraction")
        create_tooltip(self.dpi_menu, "DPI resolution")
        create_tooltip(self.pdf_folder_entry, "Select the main folder containing PDF files")
        create_tooltip(self.output_path_entry, "Select folder for the Excel output")
        create_tooltip(self.extract_button, "Start the extraction process")
        create_tooltip(self.import_button, "Import a saved template of selected areas")
        create_tooltip(self.export_button, "Export the selected areas as a template")
        create_tooltip(self.clear_areas_button, "Clear all selected areas")
        create_tooltip(self.revision_pattern_menu, "Choose the revision format pattern")


    def build_folder_tree(self):
        # ✅ Only destroy the previous Treeview, not the scrollable frame or container
        if self.files_tree_widget:
            self.files_tree_widget.destroy()
            self.files_tree_widget = None

        style = ttk.Style()
        style.configure("Treeview", font=("Segoe UI", 7))
        style.configure("Treeview.Heading", font=("Arial", 8, "bold"))

        # ✅ New Treeview inside pre-created container
        self.files_tree_widget = CheckboxTreeview(self.files_tree_container, show="tree", height=10)
        self.files_tree_widget.pack(side="left", fill="both", expand=True)
        self.files_tree_widget.column("#0", width=800, stretch=False)

        # ✅ Add vertical scrollbar (only once per build)
        v_scroll = ttk.Scrollbar(self.files_tree_container, orient="vertical", command=self.files_tree_widget.yview)
        self.files_tree_widget.configure(yscrollcommand=v_scroll.set)
        v_scroll.pack(side="right", fill="y")

        # ✅ Tree structure
        root_node = self.files_tree_widget.insert("", "end", text=os.path.basename(self.pdf_folder), tags=("checked",))

        def insert_children(parent, folder_path):
            try:
                for entry in sorted(os.listdir(folder_path)):
                    full_path = os.path.join(folder_path, entry)
                    if os.path.isdir(full_path):
                        if self.has_pdf(full_path):
                            node = self.files_tree_widget.insert(parent, "end", text=entry, tags=("checked",))
                            insert_children(node, full_path)
                    elif entry.lower().endswith(".pdf"):
                        self.files_tree_widget.insert(parent, "end", text=entry, values=[full_path], tags=("checked",))
            except Exception as e:
                print(f"Error accessing {folder_path}: {e}")

        insert_children(root_node, self.pdf_folder)


        self.files_tree_widget.bind("<<TreeviewSelect>>", lambda e: self.update_pdf_counter())
        self.files_tree_widget.bind("<ButtonRelease-1>", lambda e: self.root.after(100, self.update_pdf_counter))

        self.update_pdf_counter()

    def recursive_set_check_state(self, item_id):
        item = self.files_tree_widget.item(item_id)
        values = item.get("values", [])
        if values:
            full_path = os.path.abspath(values[0])
            tag = "checked" if full_path in self.dropped_pdf_set else "unchecked"
            self.files_tree_widget.item(item_id, tags=(tag,))
        else:
            # Folder node – recurse
            for child_id in self.files_tree_widget.get_children(item_id):
                self.recursive_set_check_state(child_id)

    def drop_pdf_folder(self, event):
        raw_data = event.data.strip()
        raw_items = re.findall(r'{(.*?)}', raw_data) or [raw_data.strip()]
        cleaned_paths = [os.path.abspath(p.strip('"')) for p in raw_items]

        dropped_pdfs = [p for p in cleaned_paths if os.path.isfile(p) and p.lower().endswith(".pdf")]
        dropped_folders = [p for p in cleaned_paths if os.path.isdir(p)]

        # Case 1: PDFs dropped directly
        if dropped_pdfs:
            self.dropped_pdf_set = set(dropped_pdfs)
            common_root = os.path.commonpath(dropped_pdfs)
            self.pdf_folder = common_root if os.path.isdir(common_root) else os.path.dirname(dropped_pdfs[0])
            self.pdf_folder_entry.delete(0, ctk.END)
            self.pdf_folder_entry.insert(0, self.pdf_folder)
            self.build_folder_tree()
            for iid in self.files_tree_widget.get_children(""):
                self.recursive_set_check_state(iid)
            self.update_pdf_counter()

            if not self.pdf_viewer.pdf_document:
                self.pdf_viewer.display_pdf(dropped_pdfs[0])
                self.recent_pdf_path = dropped_pdfs[0]
                print(f"Loaded dropped PDF: {dropped_pdfs[0]}")
            return

        # Case 2: A single folder was dropped
        if len(dropped_folders) == 1:
            self.pdf_folder = dropped_folders[0]
            self.pdf_folder_entry.delete(0, ctk.END)
            self.pdf_folder_entry.insert(0, self.pdf_folder)
            self.dropped_pdf_set = set()  # reset
            self.build_folder_tree()
            self.update_pdf_counter()
            return

        messagebox.showerror("Invalid Drop", "Please drop a PDF file or a folder.")

    def drop_sample_pdf(self, event):
        path = event.data.strip().replace("{", "").replace("}", "")
        if os.path.isfile(path) and path.lower().endswith(".pdf"):
            self.pdf_viewer.display_pdf(path)
            self.recent_pdf_path = path
            print(f"Dropped Sample PDF: {path}")
        else:
            messagebox.showerror("Invalid Drop", "Please drop a valid PDF file.")

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

    def update_pdf_counter(self):
        if not self.files_tree_widget:
            return

        def is_pdf(iid):
            item = self.files_tree_widget.item(iid)
            text = item.get("text", "")
            return text.lower().endswith(".pdf")

        checked = [iid for iid in self.files_tree_widget.get_checked() if is_pdf(iid)]
        count = len(checked)
        self.pdf_counter_label.configure(text=f"Selected PDFs: {count}")

    def has_pdf(self, folder_path):
        """Recursively checks if the folder or its subfolders contain any .pdf files."""
        try:
            for entry in os.listdir(folder_path):
                full_path = os.path.join(folder_path, entry)
                if os.path.isdir(full_path):
                    if self.has_pdf(full_path):
                        return True
                elif entry.lower().endswith(".pdf"):
                    return True
        except Exception as e:
            print(f"Error scanning {folder_path}: {e}")
        return False

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
        self.build_folder_tree()

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

        # 🧠 Extract only checked PDFs
        checked = self.files_tree_widget.get_checked()
        selected_paths = []

        for iid in checked:
            item = self.files_tree_widget.item(iid)
            if item and "values" in item and item["values"]:
                path = item["values"][0]
                if path.lower().endswith(".pdf"):
                    selected_paths.append(path)

        if not selected_paths:
            messagebox.showerror("No Files Selected", "Please check at least one PDF to extract.")
            return

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
            revision_regex=selected_revision_regex
        )
        extractor.revision_area = self.pdf_viewer.revision_area

        # ✅ Pass the revision area from the viewer to the extractor
        extractor.revision_area = self.pdf_viewer.revision_area

        # Pass selected files to start_extraction
        extraction_process = multiprocessing.Process(
            target=extractor.start_extraction,
            args=(progress_counter, total_files, final_output_path, selected_paths)  # pass files directly
        )

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
        new_width = self.root.winfo_width()
        new_height = self.root.winfo_height()

        if hasattr(self, "prev_width") and hasattr(self, "prev_height"):
            if new_width == self.prev_width and new_height == self.prev_height:
                return

        self.prev_width = new_width
        self.prev_height = new_height

        try:
            # Estimate the width of the left tabview panel

            sidebar_width = self.tab_view.winfo_width() + 20

            self.pdf_viewer.resize_canvas(self.root.winfo_width(), self.root.winfo_height(),x_offset=CANVAS_LEFT_MARGIN)

            self.pdf_viewer.update_rectangles()

            # Zoom and version controls
            self.zoom_frame.place_configure(x=sidebar_width + 10, y=new_height - 65)


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



