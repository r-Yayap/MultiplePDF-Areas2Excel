#extractor.py

import shutil
import os
import multiprocessing
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
import pymupdf as fitz
from utils import adjust_coordinates_for_rotation, find_tessdata


class TextExtractor:
    def __init__(self, pdf_folder, output_excel_path, areas, ocr_settings, include_subfolders):
        self.pdf_folder = pdf_folder
        self.output_excel_path = output_excel_path
        self.areas = areas
        self.ocr_settings = ocr_settings
        self.include_subfolders = include_subfolders
        self.tessdata_folder = find_tessdata() if ocr_settings["enable_ocr"] != "Off" else None
        self.temp_image_folder = "temp_images"
        self.headers = ["Size (Bytes)", "Date Last Modified", "Folder", "Filename"] + \
                       [f"{area['title']}" if "title" in area else f"Area {i + 1}" for i, area in enumerate(self.areas)]

        if not os.path.exists(self.temp_image_folder):
            os.makedirs(self.temp_image_folder)

    def start_extraction(self, progress_list, total_files):
        """Starts the extraction process and sends progress updates."""
        pdf_files = self.get_pdf_files()
        total_files.value = len(pdf_files)  # Set total files in the main process

        with multiprocessing.Pool() as pool:
            results = pool.starmap(self.process_single_pdf, [(pdf_path, progress_list) for pdf_path in pdf_files])

        self.consolidate_results(results)
        # Signal completion by filling the list
        progress_list.append(1)

    def process_single_pdf(self, pdf_path, progress_list):
        """Processes a single PDF file, extracting text and images as necessary."""
        try:
            pdf_document = fitz.open(pdf_path)
            file_size = os.path.getsize(pdf_path)
            last_modified_date = datetime.fromtimestamp(os.path.getmtime(pdf_path)).strftime('%Y-%m-%d %H:%M:%S')
            row_values = [
                file_size, last_modified_date,
                os.path.relpath(os.path.dirname(pdf_path), self.pdf_folder),
                os.path.basename(pdf_path)
            ]
            extracted_areas = []

            for page_number in range(pdf_document.page_count):
                page = pdf_document[page_number]
                for area_index, area in enumerate(self.areas):
                    coordinates = area["coordinates"]
                    text_area, img_path = self.extract_text_from_area(page, coordinates, pdf_path, page_number,
                                                                      area_index)
                    extracted_areas.append((text_area or "", img_path))

            pdf_document.close()
            row_values.extend(extracted_areas)

            # Increment progress by adding a placeholder value to the list
            progress_list.append(1)
            return row_values

        except Exception as e:
            print(f"Error processing {pdf_path}: {e}")
            return [file_size, last_modified_date, os.path.relpath(os.path.dirname(pdf_path), self.pdf_folder),
                    os.path.basename(pdf_path)] + ["Error"] * len(self.areas)


        except Exception as e:
            print(f"Error processing {pdf_path}: {e}")
            return [file_size, last_modified_date, os.path.relpath(os.path.dirname(pdf_path), self.pdf_folder),
                    os.path.basename(pdf_path)] + ["Error"] * len(self.areas)

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
            if self.ocr_settings["enable_ocr"] == "Off":
                text_area = page.get_text("text", clip=adjusted_coordinates)
            elif self.ocr_settings["enable_ocr"] == "Text-first":
                text_area = page.get_text("text", clip=adjusted_coordinates)
                if not text_area.strip():
                    text_area, img_path = self.apply_ocr(page, area_coordinates, pdf_path, page_number, area_index)
            elif self.ocr_settings["enable_ocr"] == "OCR-All":
                text_area, img_path = self.apply_ocr(page, area_coordinates, pdf_path, page_number, area_index)
            elif self.ocr_settings["enable_ocr"] == "Text1st+Image-beta":
                text_area = page.get_text("text", clip=adjusted_coordinates)
                pix = page.get_pixmap(clip=area_coordinates, dpi=self.ocr_settings.get("dpi_value", 150))
                img_path = os.path.join(self.temp_image_folder,
                                        f"{os.path.basename(pdf_path)}_page{page_number + 1}_area{area_index}.png")
                pix.save(img_path)
                if not text_area.strip():
                    text_area, img_path = self.apply_ocr(page, area_coordinates, pdf_path, page_number, area_index)

            text_area = text_area.replace('\n', ' ')

        except Exception as e:
            print(f"Error extracting text from area {area_index} in {pdf_path}, Page {page_number + 1}: {e}")

        return text_area, img_path

    def apply_ocr(self, page, coordinates, pdf_path, page_number, area_index):
        """Applies OCR on a specified area and returns the extracted text and image path."""
        if not self.tessdata_folder:
            print("Tessdata folder not found. OCR cannot proceed.")
            return "", ""

        pix = page.get_pixmap(clip=coordinates, dpi=self.ocr_settings.get("dpi_value", 150))
        pdfdata = pix.pdfocr_tobytes(language="eng", tessdata=self.tessdata_folder)
        clipdoc = fitz.open("pdf", pdfdata)
        text_area = "_OCR_" + clipdoc[0].get_text()
        img_path = os.path.join(self.temp_image_folder,
                                f"{os.path.basename(pdf_path)}_page{page_number + 1}_area{area_index}.png")
        pix.save(img_path)
        return text_area, img_path

    def consolidate_results(self, results):
        """Consolidates all extracted results into an Excel file."""
        wb = Workbook()
        ws = wb.active
        ws.append(self.headers)

        for row_data in results:
            row_index = ws.max_row + 1
            basic_row_data = row_data[:4]
            extracted_areas = row_data[4:]

            for col, data in enumerate(basic_row_data, start=1):
                cell = ws.cell(row=row_index, column=col)
                cell.value = data

                if col == 4:  # Link the filename cell
                    cell.hyperlink = Hyperlink(target=data, ref=cell.coordinate)
                    cell.font = Font(color="0000FF")  # Set hyperlink color to blue

            for i, area in enumerate(extracted_areas, start=5):
                # Check if area has both text and image path correctly
                if isinstance(area, tuple) and len(area) == 2:
                    text_area, img_path = area
                else:
                    text_area, img_path = area, None  # If no image path, set to None

                text_cell = ws.cell(row=row_index, column=i)
                text_cell.value = text_area.replace("_OCR_", "")
                if "_OCR_" in text_area:
                    text_cell.font = Font(color="FF3300")  # Highlight OCR text

                if img_path:
                    img = ExcelImage(img_path)
                    img.anchor = f"{get_column_letter(i)}{row_index}"
                    ws.add_image(img)

        wb.save(self.output_excel_path)
        print(f"Consolidated results saved to {self.output_excel_path}")

        # Cleanup temporary images
        if os.path.exists(self.temp_image_folder):
            shutil.rmtree(self.temp_image_folder)
