#extractor.py

import multiprocessing
import os
import re
import shutil
import getpass
import pymupdf as fitz
import sys
from datetime import datetime
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from utils import adjust_coordinates_for_rotation, find_tessdata


# Define patterns
REVISION_REGEX = re.compile(r"^[A-Z]{1,2}\d{1,2}[a-zA-Z]?$", re.IGNORECASE)
DATE_REGEX = re.compile(r"""12
    (?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}) |               # 01/01/24 or 1-1-2025
    (?:\d{1,2}\s*[-]?\s*[A-Za-z]{3,9}\s*[-]?\s*\d{2,4})  # 3-Apr-2025 or 3 April 25
""", re.VERBOSE)
DESC_KEYWORDS = ["issued for", "issue", "submission", "schematic", "detailed", "concept", "design", "construction", "revised", "resubmission"]



class TextExtractor:
    def __init__(self, pdf_folder, output_excel_path, areas, ocr_settings, include_subfolders, revision_regex=None):
        self.header_column_map = None
        self.final_output_path = None
        self.pdf_folder = pdf_folder
        self.output_excel_path = output_excel_path
        self.areas = areas
        self.ocr_settings = ocr_settings
        self.include_subfolders = include_subfolders
        self.revision_regex = re.compile(revision_regex, re.IGNORECASE) if revision_regex else REVISION_REGEX
        self.tessdata_folder = find_tessdata() if ocr_settings["enable_ocr"] != "Off" else None

        # Initialize headers with fixed metadata columns
        self.headers = ["Size (Bytes)", "Date Last Modified", "Folder", "Filename", "Page No"]

        # Dictionary to store unique header assignments per area
        self.unique_headers_mapping = {}  # Maps each rectangle index to a unique column header

        header_count = {}  # To count occurrences of each title

        for i, area in enumerate(self.areas):
            title = area.get("title", f"Area {i + 1}")  # Default title if not manually set

            # Ensure unique title for each rectangle
            if title not in header_count:
                header_count[title] = 1
                unique_title = title
            else:
                header_count[title] += 1
                unique_title = f"{title} ({header_count[title]})"

            self.headers.append(unique_title)  # Add unique header to headers list
            self.unique_headers_mapping[i] = unique_title  # Assign rectangle index to its specific column

        # Define application directory (where the script is located)


        if getattr(sys, 'frozen', False):  # Check if running as PyInstaller EXE
            app_directory = os.path.dirname(sys.executable)  # Use the real EXE location
        else:
            app_directory = os.path.dirname(os.path.abspath(__file__))  # Running as script

        # Create a main "temp" folder beside the application
        main_temp_folder = os.path.join(app_directory, "temp")

        # Create a unique subfolder for each session
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        username = getpass.getuser()  # Get the current Windows username
        self.temp_image_folder = os.path.join(main_temp_folder, f"{username}-{timestamp}")

        # Revision Area
        self.revision_area = None  # Will be set externally if needed
        self.revision_data_mapping = {}  # (filename, page_number) → list of revision strings

    def detect_column_indices(self, sample_rows, max_rows=3):
        col_scores = {}

        for row in sample_rows[:max_rows]:
            for col_idx, cell in enumerate(row):
                text = str(cell).strip().lower()
                if not text:
                    continue
                if col_idx not in col_scores:
                    col_scores[col_idx] = {"rev": 0, "date": 0, "desc": 0}

                if self.revision_regex.fullmatch(text.upper()):
                    col_scores[col_idx]["rev"] += 1
                elif DATE_REGEX.search(text):
                    col_scores[col_idx]["date"] += 1
                elif any(kw in text for kw in DESC_KEYWORDS):
                    col_scores[col_idx]["desc"] += 1

        rev_idx = max(col_scores.items(), key=lambda x: x[1]["rev"], default=(None, {}))[0]
        desc_idx = max(col_scores.items(), key=lambda x: x[1]["desc"], default=(None, {}))[0]
        date_idx = max(col_scores.items(), key=lambda x: x[1]["date"], default=(None, {}))[0]

        print(f"🔍 Column Indices Detected → Rev: {rev_idx}, Desc: {desc_idx}, Date: {date_idx}")
        return rev_idx, desc_idx, date_idx

    def parse_revision_row(self, row, rev_idx, desc_idx, date_idx):
        rev = row[rev_idx].strip() if rev_idx is not None and rev_idx < len(row) else None
        desc = row[desc_idx].strip() if desc_idx is not None and desc_idx < len(row) else None
        date = row[date_idx].strip() if date_idx is not None and date_idx < len(row) else None

        # Validate formats
        if rev and not self.revision_regex.fullmatch(rev.upper()):
            rev = None
        if date and not DATE_REGEX.search(date):
            date = None

        return rev, desc, date

    def extract_revision_history_from_page_obj(self, page, revision_coordinates):
        try:

            if not revision_coordinates or len(revision_coordinates) != 4:
                return []

            # Skip if area is too small or outside page
            clip_rect = fitz.Rect(revision_coordinates)
            if clip_rect.is_empty or clip_rect.get_area() < 100:  # Skip tiny boxes
                print(f"⚠️ Revision area too small or empty, skipping.")
                return []

            tables = page.find_tables(clip=clip_rect)
            print(f"📋 Running revision extraction")
            if not tables or not tables.tables:
                print("❌ No tables found.")
                return []

            table = tables.tables[0]
            data = table.extract()
            if not data or len(data) < 2:
                print("❌ Not enough rows to detect header or parse data.")
                return []

            data = data[::-1]
            print("📋 Table Detected:")
            for i, row in enumerate(data):
                print(f"🔹 Row {i}: {row}")

            rev_idx, desc_idx, date_idx = self.detect_column_indices(data[1:])
            extracted = []
            non_empty_rows = (row for row in data if any(row))
            for row in non_empty_rows:
                rev, desc, date = self.parse_revision_row(row, rev_idx, desc_idx, date_idx)
                if rev and desc and date:
                    print(f"✅ Parsed → {rev} | {desc} | {date}")
                    extracted.append(f"{rev} | {desc} | {date}")
                else:
                    print(f"⛔ Incomplete fields → Skipped")

            return extracted
        except Exception as e:
            print(f"❌ Error during revision extraction: {e}")
            return []

    def clean_text(self, text):
        """Cleans text by replacing newlines, stripping, and removing illegal characters."""
        replacement_char = '■'  # Character to replace prohibited control characters

        # Step 1: Replace newline and carriage return characters with a space
        text = text.replace('\n', ' ').replace('\r', ' ')

        # Step 2: Strip leading and trailing whitespace
        text = text.strip()

        # Step 3: Replace prohibited control characters with a replacement character
        text = re.sub(r'[\x00-\x1F\x7F-\x9F]', replacement_char, text)

        # Step 4: Remove extra spaces between words
        return re.sub(r'\s+', ' ', text)

    def start_extraction(self, progress_list, total_files, final_output_path):
        self.final_output_path = final_output_path  # ✅ Store reference for later use

        """Starts the extraction process using multiprocessing with progress tracking."""
        try:
            with multiprocessing.Pool() as pool:
                # Gather all PDF files in the specified folder
                pdf_files = self.get_pdf_files()
                total_files.value = len(pdf_files)  # Set total files count

                # Process each PDF file using multiprocessing with progress tracking
                results = pool.starmap(self.process_single_pdf,
                                       [(pdf_path, progress_list, total_files) for pdf_path in pdf_files])

                # Consolidate results into an Excel file after extraction is complete
                self.consolidate_results(results, final_output_path)
        except Exception as e:
            print(f"Error during extraction: {e}")
            return None

    def process_single_pdf(self, pdf_path, progress_list, total_files):
        """Processes a single PDF file, extracting text and images as necessary."""
        try:
            # Ensure the file exists and is not empty
            if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
                print(f"Skipping missing or empty file: {pdf_path}")
                result_rows = [["Error", "File not found or empty", "", pdf_path, "",
                                "FileNotFoundError: File not found or empty"]]
                progress_list.append(result_rows)
                return result_rows

            # Open the PDF file
            pdf_document = fitz.open(pdf_path)
            file_size = os.path.getsize(pdf_path)
            last_modified_date = datetime.fromtimestamp(os.path.getmtime(pdf_path)).strftime('%Y-%m-%d %H:%M:%S')
            folder = os.path.relpath(os.path.dirname(pdf_path), self.pdf_folder)
            filename = os.path.basename(pdf_path)

            result_rows = []  # Collect each page's data here

            # Process each page safely
            for page_number in range(pdf_document.page_count):
                try:
                    page = pdf_document[page_number]
                    page.remove_rotation()
                    extracted_areas = []

                    for area_index, area in enumerate(self.areas):
                        coordinates = area["coordinates"]
                        text_area, img_path = self.extract_text_from_area(page, coordinates, pdf_path, page_number,
                                                                          area_index)

                        # If text_area is empty, treat it as blank rather than "Error"
                        extracted_areas.append((text_area if text_area != "Error" else "", img_path or ""))

                    revision_data = []
                    if self.revision_area:
                        revision_data = self.extract_revision_history_from_page_obj(page, self.revision_area["coordinates"])


                    # Append all relevant data for this page as a row in results
                    result_rows.append(
                        [file_size, last_modified_date, folder, filename, page_number + 1, extracted_areas,
                         revision_data]
                    )

                except Exception as e:
                    print(f"Error processing page {page_number + 1} in {pdf_path}: {e}")
                    result_rows.append(
                        [file_size, last_modified_date, folder, filename, page_number + 1, "Error processing page",
                         str(e)])

            pdf_document.close()

            # Add each page to results
            progress_list.append(result_rows)
            return result_rows

        except Exception as e:
            print(f"Error processing {pdf_path}: {e}")
            # Add an error placeholder for this file in results
            result_rows = [
                [file_size, last_modified_date, folder, filename, "Error", f"File Processing Error: {str(e)}"]]
            progress_list.append(result_rows)
            return result_rows

    def get_pdf_files(self):
        """Gathers all PDF files within the specified folder."""
        pdf_files = []
        for root_folder, subfolders, files in os.walk(self.pdf_folder):
            if not self.include_subfolders:
                subfolders.clear()
            pdf_files.extend(
                [os.path.join(root_folder, f) for f in files if f.lower().endswith('.pdf')]
            )
        return pdf_files

    def extract_text_from_area(self, page, area_coordinates, pdf_path, page_number, area_index):
        """Extracts text or image content from a specified area in a PDF page."""
        adjusted_coordinates = adjust_coordinates_for_rotation(
            area_coordinates, page.rotation, page.rect.height, page.rect.width
        )
        text_area = ""
        img_path = None

        try:
            # Text extraction without OCR
            if self.ocr_settings["enable_ocr"] == "Off":
                text_area = page.get_text("text", clip=adjusted_coordinates)

            # OCR only if there's no text
            elif self.ocr_settings["enable_ocr"] == "Text-first":
                text_area = page.get_text("text", clip=adjusted_coordinates)
                if not text_area.strip():  # Perform OCR only if text_area is empty
                    text_area, _ = self.apply_ocr(page, area_coordinates, pdf_path, page_number, area_index)

            # OCR for all areas, no images saved
            elif self.ocr_settings["enable_ocr"] == "OCR-All":
                text_area, _ = self.apply_ocr(page, area_coordinates, pdf_path, page_number, area_index)

            # OCR with image saving regardless of extracted text for Text1st+Image-beta
            elif self.ocr_settings["enable_ocr"] == "Text1st+Image-beta":
                text_area = page.get_text("text", clip=adjusted_coordinates)
                pix = page.get_pixmap(clip=area_coordinates, dpi=self.ocr_settings.get("dpi_value", 150))

                if not os.path.exists(self.temp_image_folder):
                    os.makedirs(self.temp_image_folder, exist_ok=True)

                img_path = os.path.join(self.temp_image_folder, f"{os.path.basename(pdf_path)}_page{page_number + 1}_area{area_index}.png")
                pix.save(img_path)

                # Apply OCR if no text found
                if not text_area.strip():
                    text_area, _ = self.apply_ocr(page, area_coordinates, pdf_path, page_number, area_index)

            # Clean the extracted text
            text_area = self.clean_text(text_area)

            # Return empty string if text_area is blank after cleaning
            return (text_area if text_area.strip() else "", img_path)

        except fitz.EmptyFileError as e:
            print(f"Error: {pdf_path} is empty or corrupted: {e}")
            return f"EmptyFileError: {str(e)}", None

        except fitz.FileNotFoundError as e:
            print(f"Error: {pdf_path} not found: {e}")
            return f"FileNotFoundError: {str(e)}", None

        except RuntimeError as e:
            print(f"Runtime error on page {page_number + 1} in {pdf_path}: {e}")
            return f"RuntimeError: {str(e)}", None

        except Exception as e:
            print(f"Unexpected error extracting text from area {area_index} in {pdf_path}, Page {page_number + 1}: {e}")
            return f"UnexpectedError: {str(e)}", None

    def apply_ocr(self, page, coordinates, pdf_path, page_number, area_index):
        """Applies OCR on a specified area and returns the extracted text and image path."""
        if not self.tessdata_folder:
            print("Tessdata folder not found. OCR cannot proceed.")
            return "", ""

        if not os.path.exists(self.temp_image_folder):
             os.makedirs(self.temp_image_folder, exist_ok=True)

        pix = page.get_pixmap(clip=coordinates, dpi=self.ocr_settings.get("dpi_value", 150))
        pdfdata = pix.pdfocr_tobytes(language="eng", tessdata=self.tessdata_folder)
        clipdoc = fitz.open("pdf", pdfdata)
        text_area = "_OCR_" + clipdoc[0].get_text()
        img_path = os.path.join(self.temp_image_folder,f"{os.path.basename(pdf_path)}_page{page_number + 1}_area{area_index}.png")
        pix.save(img_path)
        return text_area, img_path

    def extract_revision_history_from_page(self, pdf_path, page, filename, page_number):
        try:
            if (
                    self.revision_area and
                    isinstance(self.revision_area, dict) and
                    "coordinates" in self.revision_area and
                    isinstance(self.revision_area["coordinates"], list) and
                    len(self.revision_area["coordinates"]) == 4
            ):

                revision_data = self.extract_revision_history_from_page_obj(page, self.revision_area["coordinates"])
                self.revision_data_mapping[(filename, page_number + 1)] = revision_data
            else:
                print(f"⚠️ Invalid revision area format for {filename} Page {page_number + 1}: {self.revision_area}")
        except Exception as e:
            print(f"⚠️ Revision extraction failed for {filename} Page {page_number + 1}: {e}")

    def consolidate_results(self, results, final_output_path):
        """Consolidates all extracted results into an Excel file with each page as a row."""
        wb = Workbook()
        ws = wb.active

        # Determine the maximum number of revision entries
        max_revision_columns = max(
            (len(row[6]) for result_pages in results for row in result_pages if
             isinstance(row, list) and len(row) == 7),
            default=0
        )

        # Extend headers dynamically with Rev1, Rev2, ...
        rev_headers = [f"Rev{i + 1}" for i in range(max_revision_columns)]
        self.headers.extend(rev_headers)

        self.header_column_map = {title: idx + 1 for idx, title in enumerate(self.headers)}

        ws.append(self.headers)  # Add headers to the first row

        # Precompute revision column indices to avoid repeated lookups
        rev_col_indices = {
            i: self.header_column_map[header]
            for i, header in enumerate(rev_headers)
        }

        for result_pages in results:
            for row_data in result_pages:
                try:
                    if not isinstance(row_data, list) or len(row_data) != 7:
                        print(f"⚠️ Skipping malformed row_data: {row_data}")
                        continue

                    # Unpack metadata and extracted area data for the page
                    file_size, last_modified, folder, filename, page_number, extracted_areas, revision_data = row_data

                    # Prepare row for insertion
                    row_index = ws.max_row + 1
                    metadata = [file_size, last_modified, folder, filename, page_number]

                    # Write metadata to columns 1-5
                    for col, data in enumerate(metadata, start=1):
                        cell = ws.cell(row=row_index, column=col)
                        cell.value = data
                        if col == 4:  # Set filename as a hyperlink in column 4
                            absolute_path = os.path.abspath(os.path.join(self.pdf_folder, folder, filename))
                            cell.hyperlink = absolute_path
                            cell.style = "Hyperlink"

                    # Write extracted areas to columns starting from 6
                    for index, area in enumerate(self.areas):  # Use the updated `self.areas` list

                        # Get the unique title assigned to this rectangle
                        column_title = self.unique_headers_mapping.get(index, f"Area {index + 1}")
                        col_index = self.header_column_map[column_title]

                        # Retrieve text and image path from extracted areas
                        text, img_path = extracted_areas[index] if isinstance(extracted_areas, list) else (
                        extracted_areas, None)

                        # Place text in the designated column, checking for OCR and cleaning
                        text_cell = ws.cell(row=row_index, column=col_index)
                        cleaned_text = self.clean_text(text.replace("_OCR_", "")) if text != "Error" else "Error"
                        text_cell.value = cleaned_text

                        # Highlight OCR text in red if "_OCR_" was detected
                        if "_OCR_" in text:
                            text_cell.font = Font(color="FF3300")

                        # If there is an associated image, add it to the cell
                        if img_path and img_path != "Error":
                            img = ExcelImage(img_path)
                            img.anchor = f"{get_column_letter(col_index)}{row_index}"
                            ws.add_image(img)

                    # Append revisions if available
                    for i, revision_text in enumerate(revision_data):
                        if i in rev_col_indices:
                            col_index = rev_col_indices[i]
                            ws.cell(row=row_index, column=col_index).value = revision_text

                except Exception as e:
                    print(f"❌ Unexpected error consolidating data: {e}")
                    print(f"   ↳ row_data was: {row_data}")
                    ws.append(["Error", "Consolidation failed", str(e)] + [""] * len(self.areas))

        # Generate a unique filename if the output file already exists
        output_filename = self.output_excel_path
        if os.path.exists(output_filename):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name, file_ext = os.path.splitext(output_filename)
            output_filename = f"{file_name}_{timestamp}{file_ext}"

        # Save to the new filename
        try:
            wb.save(output_filename)
            wb.close()
            final_output_path.value = output_filename  # ✅ Save to shared multiprocessing Value

            print(f"Consolidated results saved to {output_filename}")
        except Exception as e:
            print(f"Error saving Excel file: {e}")

        # Cleanup temporary images
        if os.path.exists(self.temp_image_folder):
            print(f"Temp Images will be deleted: {self.temp_image_folder}")
            shutil.rmtree(self.temp_image_folder)


