import threading
import time
import os
from datetime import datetime
import pymupdf as fitz
import customtkinter as ctk
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.drawing.image import Image as ExcelImage
from tkinter import messagebox
from openpyxl.utils import get_column_letter
from utils import adjust_coordinates_for_rotation  # Import the functio
from utils import *

class TextExtractor:
    def __init__(self, root, pdf_folder, output_excel_path, areas, ocr_settings, include_subfolders):
        self.root = root
        self.pdf_folder = pdf_folder
        self.output_excel_path = output_excel_path
        self.areas = areas
        self.ocr_settings = ocr_settings
        self.include_subfolders = include_subfolders
        self.wb = Workbook()
        self.ws = self.wb.active
        self.total_files = self.count_pdf_files()

        # Find tessdata path once and store it for OCR
        self.tessdata_folder = find_tessdata()

    def count_pdf_files(self):
        # Count all PDF files for progress tracking
        count = 0
        for root_folder, subfolders, files in os.walk(self.pdf_folder):
            if not self.include_subfolders:
                subfolders.clear()
            count += sum(1 for file in files if file.lower().endswith('.pdf'))
        return count

    def start_extraction_thread(self):
        extraction_thread = threading.Thread(target=self.extract_text)
        extraction_thread.start()

    def extract_text(self):
        self.close_pdf()
        start_time = time.time()

        if not self.areas:
            messagebox.showwarning("Error", "No areas defined. Please draw rectangles.")
            return

        if not os.path.isdir(self.pdf_folder):
            messagebox.showerror("Invalid Folder", "The specified PDF folder does not exist.")
            return
        if not self.output_excel_path:
            messagebox.showerror("Error", "Please enter a valid output path.")
            return

        self.ws = self.wb.active
        self.setup_excel_sheet()
        progress_window, progress_label, progress = self.setup_progress_window()

        processed_files = 0
        temp_image_paths = []

        for root_folder, subfolders, files in os.walk(self.pdf_folder):
            if not self.include_subfolders:
                subfolders.clear()

            for pdf_filename in files:
                if pdf_filename.lower().endswith('.pdf'):
                    pdf_path = os.path.join(root_folder, pdf_filename)
                    self.process_pdf_file(pdf_path, root_folder, temp_image_paths)  # Only pass necessary parameters
                    processed_files += 1

                    # Update progress based on total file count
                    progress.set(processed_files / self.total_files)
                    progress_label.configure(text=f"Processing PDF {processed_files}/{self.total_files}")
                    progress_label.update_idletasks()

        self.finalize_excel(temp_image_paths)

        end_time = time.time()
        elapsed_time = end_time - start_time
        self.display_elapsed_time(elapsed_time)
        progress_window.destroy()

    def setup_excel_sheet(self):
        # Extract titles from `self.areas` for headers
        headers = ["Size (Bytes)", "Date Last Modified", "Folder", "Filename"] + \
                  [area["title"] if "title" in area else f"Area {i + 1}" for i, area in enumerate(self.areas)]
        self.ws.append(headers)

    def setup_progress_window(self):
        # Set up the progress window for displaying progress
        progress_window = ctk.CTkToplevel(self.root)
        progress_window.title("Extraction in Progress...")
        progress_window.geometry("300x90")
        progress_window.attributes('-topmost', True)
        progress_label = ctk.CTkLabel(progress_window, text="Loading PDFs...")
        progress_label.pack()
        progress = ctk.CTkProgressBar(progress_window, orientation="horizontal", mode="determinate",
                                      progress_color='limegreen', width=150, height=15)
        progress.pack()
        return progress_window, progress_label, progress

    def process_pdf_file(self, pdf_path, root_folder, temp_image_paths):
        try:
            file_size = os.path.getsize(pdf_path)
            last_modified_date = datetime.fromtimestamp(os.path.getmtime(pdf_path)).strftime('%Y-%m-%d %H:%M:%S')
            pdf_document = fitz.open(pdf_path)

            row_values = [
                file_size,
                last_modified_date,
                os.path.relpath(root_folder, self.pdf_folder),
                pdf_path.split(os.sep)[-1]
            ]

            for page_number in range(pdf_document.page_count):
                page = pdf_document[page_number]
                sort_val = False if page.rotation == 0 else False

                for i, area in enumerate(self.areas, start=1):
                    coordinates = area["coordinates"]
                    self.extract_text_from_area(page, coordinates, row_values, temp_image_paths, pdf_path,
                                                page_number, i, sort_val)

            # Append extracted data to Excel
            self.ws.append(row_values)

            # Add a hyperlink to the filename cell
            pdf_filename_cell = self.ws.cell(row=self.ws.max_row, column=4)
            pdf_filename_cell.value = pdf_path.split(os.sep)[-1]
            pdf_filename_cell.hyperlink = Hyperlink(target=pdf_path, ref=pdf_filename_cell.coordinate)
            pdf_filename_cell.font = Font(color="0000FF")  # Set font color to blue for hyperlink

        except Exception as e:
            print(f"Error processing {pdf_path}: {e}")
            self.ws.append([os.path.relpath(root_folder, self.pdf_folder), pdf_path, "Error"] + [""] * len(self.areas))

    def extract_text_from_area(self, page, area_coordinates, row_values, temp_image_paths, pdf_path, page_number,
                               area_index, sort_val):
        pdf_width, pdf_height = page.rect.width, page.rect.height
        adjusted_coordinates = adjust_coordinates_for_rotation(area_coordinates, page.rotation, pdf_height, pdf_width)
        text_area = ""

        try:
            # OCR Mode: Off - Only normal extraction
            if self.ocr_settings["enable_ocr"] == "Off":
                text_area = page.get_text("text", clip=adjusted_coordinates, sort=sort_val)

            # OCR Mode: Text-first - OCR if text is blank
            elif self.ocr_settings["enable_ocr"] == "Text-first":
                text_area = page.get_text("text", clip=adjusted_coordinates, sort=sort_val)
                if not text_area.strip():
                    text_area, img_path = self.apply_ocr(page, area_coordinates, pdf_path, page_number, area_index)
                    temp_image_paths.append(img_path)

            # OCR Mode: OCR-All - Always use OCR
            elif self.ocr_settings["enable_ocr"] == "OCR-All":
                text_area, img_path = self.apply_ocr(page, area_coordinates, pdf_path, page_number, area_index)
                temp_image_paths.append(img_path)

            # OCR Mode: Text1st+Image-beta - OCR if text is blank, and include image
            elif self.ocr_settings["enable_ocr"] == "Text1st+Image-beta":
                text_area = page.get_text("text", clip=adjusted_coordinates, sort=sort_val)
                pix = page.get_pixmap(clip=area_coordinates, dpi=self.ocr_settings.get("dpi_value", 150))
                if not text_area.strip():
                    text_area, img_path = self.apply_ocr(page, area_coordinates, pdf_path, page_number, area_index)
                else:
                    img_path = f"{pdf_path.split(os.sep)[-1]}_page{page_number + 1}_area{area_index}.png"
                    pix.save(img_path)
                temp_image_paths.append(img_path)

                # Embed image in Excel if available
                img = ExcelImage(img_path)
                img_cell = self.ws.cell(row=self.ws.max_row, column=len(row_values) + 1)
                img_anchor = f"{get_column_letter(img_cell.column)}{img_cell.row + 1}"
                img.anchor = img_anchor
                self.ws.add_image(img)

            row_values.append(text_area)
            print(f"Page {page_number + 1}, Area {area_index} - Sample Extracted Text: {text_area}")

        except Exception as e:
            print(f"Error extracting text from area {area_index} in {pdf_path}, Page {page_number + 1}: {e}")

    def apply_ocr(self, page, area_coordinates, pdf_path, page_number, area_index):
        # Ensure tessdata_folder is available
        if not self.tessdata_folder:
            print("Tessdata folder not found. OCR cannot proceed.")
            return "", ""

        # Apply OCR using the tessdata_folder path
        pix = page.get_pixmap(clip=area_coordinates, dpi=self.ocr_settings.get("dpi_value", 150))
        pdfdata = pix.pdfocr_tobytes(language="eng", tessdata=self.tessdata_folder)  # Use tessdata_folder here
        clipdoc = fitz.open("pdf", pdfdata)
        text_area = "_OCR_" + clipdoc[0].get_text()

        # Save image to a temporary path for embedding in Excel
        img_path = f"{pdf_path.split(os.sep)[-1]}_page{page_number + 1}_area{area_index}.png"
        pix.save(img_path)
        return text_area, img_path

    def finalize_excel(self, temp_image_paths):
        try:
            for row in self.ws.iter_rows(min_row=2, max_row=self.ws.max_row):
                for cell in row:
                    if "_OCR_" in str(cell.value):
                        cell.value = cell.value.replace("_OCR_", "")
                        cell.font = Font(color="FF3300")
            self.wb.save(self.output_excel_path)

            for img_path in temp_image_paths:
                try:
                    os.remove(img_path)
                except OSError as e:
                    print(f"Error: {img_path} : {e.strerror}")

        except PermissionError:
            timestamp = time.strftime("%Y%m%d%H%M%S")
            timestamped_output_path = f"{os.path.splitext(self.output_excel_path)[0]}_{timestamp}.xlsx"
            self.wb.save(timestamped_output_path)
            print(f"A copy has been created: {timestamped_output_path}")

    def display_elapsed_time(self, elapsed_time):
        open_file = messagebox.askyesno("Open Excel File", f"Elapsed Time: {elapsed_time:.2f} seconds\n\nDo you want to open the Excel file now?")
        if open_file:
            os.startfile(self.output_excel_path)

    def close_pdf(self):
        # Close any open PDF documents if necessary
        if hasattr(self, 'pdf_document') and self.pdf_document:
            self.pdf_document.close()
