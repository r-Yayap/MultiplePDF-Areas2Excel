from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
import os
import shutil
from datetime import datetime
import concurrent.futures
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
                       [f"{area['title']}" if "title" in area else f"Area {i+1}" for i, area in enumerate(self.areas)]

        # Create the image directory if it does not exist
        if not os.path.exists(self.temp_image_folder):
            os.makedirs(self.temp_image_folder)

    def start_extraction(self):
        with concurrent.futures.ProcessPoolExecutor() as executor:
            pdf_files = self.get_pdf_files()
            futures = {executor.submit(self.process_single_pdf, pdf_path): pdf_path for pdf_path in pdf_files}
            results = []
            for future in concurrent.futures.as_completed(futures):
                try:
                    result = future.result()
                    results.append(result)
                except Exception as e:
                    print(f"Error processing {futures[future]}: {e}")
        self.consolidate_results(results)

    def get_pdf_files(self):
        pdf_files = []
        for root_folder, subfolders, files in os.walk(self.pdf_folder):
            if not self.include_subfolders:
                subfolders.clear()
            pdf_files.extend(
                [os.path.join(root_folder, f) for f in files if f.lower().endswith('.pdf')]
            )
        return pdf_files

    def process_single_pdf(self, pdf_path):
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
                    text_area, img_path = self.extract_text_from_area(page, coordinates, pdf_path, page_number, area_index)
                    extracted_areas.append((text_area or "", img_path))

            pdf_document.close()
            row_values.extend(extracted_areas)
            return row_values

        except Exception as e:
            print(f"Error processing {pdf_path}: {e}")
            return [file_size, last_modified_date, os.path.relpath(os.path.dirname(pdf_path), self.pdf_folder),
                    os.path.basename(pdf_path)] + ["Error"] * len(self.areas)

    def extract_text_from_area(self, page, area_coordinates, pdf_path, page_number, area_index):
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
                img_path = os.path.join(self.temp_image_folder, f"{os.path.basename(pdf_path)}_page{page_number + 1}_area{area_index}.png")
                pix.save(img_path)
                if not text_area.strip():
                    text_area, img_path = self.apply_ocr(page, area_coordinates, pdf_path, page_number, area_index)

            text_area = text_area.replace('\n', ' ')

        except Exception as e:
            print(f"Error extracting text from area {area_index} in {pdf_path}, Page {page_number + 1}: {e}")

        return text_area, img_path

    def apply_ocr(self, page, coordinates, pdf_path, page_number, area_index):
        if not self.tessdata_folder:
            print("Tessdata folder not found. OCR cannot proceed.")
            return "", ""

        pix = page.get_pixmap(clip=coordinates, dpi=self.ocr_settings.get("dpi_value", 150))
        pdfdata = pix.pdfocr_tobytes(language="eng", tessdata=self.tessdata_folder)
        clipdoc = fitz.open("pdf", pdfdata)
        text_area = "_OCR_" + clipdoc[0].get_text()
        img_path = os.path.join(self.temp_image_folder, f"{os.path.basename(pdf_path)}_page{page_number + 1}_area{area_index}.png")
        pix.save(img_path)
        return text_area, img_path

    def consolidate_results(self, results):
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

            for i, (text_area, img_path) in enumerate(extracted_areas, start=5):
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

        # Cleanup
        if os.path.exists(self.temp_image_folder):
            shutil.rmtree(self.temp_image_folder)
