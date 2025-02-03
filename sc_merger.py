import os
import re
import tkinter as tk
from difflib import SequenceMatcher
from tkinter import filedialog, messagebox
from datetime import datetime
import functools

import customtkinter as ctk
import pandas as pd
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.utils import get_column_letter
import unicodedata

# Import tkinterdnd2 for drag-and-drop
from tkinterdnd2 import TkinterDnD, DND_ALL


# Helper to clear a docx paragraph's contents
def clear_paragraph(paragraph):
    p = paragraph._element
    for child in list(p):
        p.remove(child)

# Helper to auto-select a header based on keywords.
def auto_select_header(headers, keywords):
    for header in headers:
        lower_header = header.lower()
        for kw in keywords:
            if kw in lower_header:
                return header
    return headers[0] if headers else ""


class ExcelMerger:
    """Handles Excel merging logic, including conditional formatting, hyperlink handling,
    and title rich-text highlighting (using tokenization and dynamic programming)."""

    @staticmethod
    def merge_3_excels(excel1_path, excel2_path, excel3_path,
                       ref_column1, ref_column2, ref_column3,
                       output_path,
                       title_column1=None, title_column2=None, title_column3=None,
                       compare_excel3=False):
        # Read Excel files
        df1 = pd.read_excel(excel1_path, engine='openpyxl', dtype=str).fillna("")
        df2 = pd.read_excel(excel2_path, engine='openpyxl', dtype=str).fillna("")
        df3 = pd.read_excel(excel3_path, engine='openpyxl', dtype=str).fillna("")

        # Ensure the reference columns exist
        for df, ref_col in zip([df1, df2, df3],
                               [ref_column1, ref_column2, ref_column3]):
            if ref_col not in df.columns:
                raise KeyError(f"Reference column '{ref_col}' not found in one of the Excel files.")

        # For hyperlink purposes, add original row indices to df1
        df1 = ExcelMerger.add_original_row_index_to_dataframe(df1, excel1_path)
        # Extract hyperlinks from Excel1
        hyperlinks = ExcelMerger._extract_hyperlinks(excel1_path)

        # Compute occurrence counts
        df1['refno_count'] = df1.groupby(ref_column1).cumcount()
        df2['refno_count'] = df2.groupby(ref_column2).cumcount()
        df3['refno_count'] = df3.groupby(ref_column3).cumcount()

        # Rename reference columns to fixed names
        df1 = df1.rename(columns={ref_column1: 'number_1'})
        df2 = df2.rename(columns={ref_column2: 'number_2'})
        df3 = df3.rename(columns={ref_column3: 'number_3'})

        # Create a common key column from the reference values
        df1['common_ref'] = df1['number_1']
        df2['common_ref'] = df2['number_2']
        df3['common_ref'] = df3['number_3']

        # Rename title columns if provided
        if title_column1:
            df1 = df1.rename(columns={title_column1: 'title_excel1'})
        if title_column2:
            df2 = df2.rename(columns={title_column2: 'title_excel2'})
        if title_column3:
            df3 = df3.rename(columns={title_column3: 'title_excel3'})

        # Merge df1 and df2 on common_ref and refno_count
        merged_df = pd.merge(
            df1, df2,
            on=['common_ref', 'refno_count'],
            how='outer',
            suffixes=('_1', '_2')
        ).fillna("")

        # Merge the result with df3
        merged_df = pd.merge(
            merged_df, df3,
            on=['common_ref', 'refno_count'],
            how='outer',
            suffixes=("", "_3")
        ).drop(columns=['refno_count']).fillna("")

        if 'original_row_index' in merged_df.columns:
            merged_df['original_row_index'] = pd.to_numeric(
                merged_df['original_row_index'], errors='coerce').fillna(0).astype(int)

        print("Merged DataFrame columns (3 files):")
        print(merged_df.columns.tolist())

        temp_file_path = ExcelMerger._save_merged_to_excel(merged_df, output_path)
        ExcelMerger._apply_formatting_and_hyperlinks(temp_file_path, hyperlinks, merged_df)

        # *** IMPORTANT: First, reorder columns ***
        ExcelMerger.reorder_columns(temp_file_path)

        # *** Then apply title highlighting ***
        # Highlight differences between title_excel1 and title_excel2 (update both columns)
        if title_column1 and title_column2:
            ExcelMerger.apply_title_highlighting(
                temp_file_path, merged_df, 'title_excel1', 'title_excel2',
                reorder=False, update_baseline=True
            )
        # Highlight differences between title_excel1 and title_excel3, but update only title_excel3
        if title_column1 and title_column3 and compare_excel3:
            ExcelMerger.apply_title_highlighting(
                temp_file_path, merged_df, 'title_excel1', 'title_excel3',
                reorder=False, update_baseline=False
            )

        return temp_file_path, merged_df

    @staticmethod
    def merge_excels(excel1_path, excel2_path, ref_column1, ref_column2, output_path,
                     title_column1=None, title_column2=None):
        # Two-file merge code (same as before)
        excel1 = pd.read_excel(excel1_path, engine='openpyxl', dtype=str).fillna("")
        excel2 = pd.read_excel(excel2_path, engine='openpyxl', dtype=str).fillna("")
        if ref_column1 not in excel1.columns or ref_column2 not in excel2.columns:
            raise KeyError("Reference columns not found in one or both Excel files.")
        excel1 = ExcelMerger.add_original_row_index_to_dataframe(excel1, excel1_path)
        hyperlinks = ExcelMerger._extract_hyperlinks(excel1_path)
        excel1['refno_count'] = excel1.groupby(ref_column1).cumcount()
        excel2['refno_count'] = excel2.groupby(ref_column2).cumcount()
        excel1 = excel1.rename(columns={ref_column1: 'number_1'})
        excel2 = excel2.rename(columns={ref_column2: 'number_2'})
        excel1['common_ref'] = excel1['number_1']
        excel2['common_ref'] = excel2['number_2']
        if title_column1:
            excel1 = excel1.rename(columns={title_column1: 'title_excel1'})
        if title_column2:
            excel2 = excel2.rename(columns={title_column2: 'title_excel2'})

        merged_df = pd.merge(
            excel1, excel2,
            on=['common_ref', 'refno_count'],
            how='outer',
            suffixes=('_1', '_2')
        ).drop(columns=['refno_count']).fillna("")

        merged_df['original_row_index'] = pd.to_numeric(
            merged_df['original_row_index'], errors='coerce').fillna(0).astype(int)
        print("Merged DataFrame columns:")
        print(merged_df.columns.tolist())
        temp_file_path = ExcelMerger._save_merged_to_excel(merged_df, output_path)
        ExcelMerger._apply_formatting_and_hyperlinks(temp_file_path, hyperlinks, merged_df)
        if title_column1 and title_column2:
            ExcelMerger.apply_title_highlighting(temp_file_path, merged_df, 'title_excel1', 'title_excel2')
        return temp_file_path, merged_df

    @staticmethod
    def reorder_columns(file_path):
        wb = load_workbook(file_path)
        ws = wb.active

        # Delete the original_row_index column.
        # (Assume the header for that column is still "original_row_index".)
        orig_index_col_idx = None
        for cell in ws[1]:
            if cell.value == "original_row_index":
                orig_index_col_idx = cell.column
                break
        if orig_index_col_idx is not None:
            ws.delete_cols(orig_index_col_idx, 1)
            print(f"Dropped 'original_row_index' column at position {orig_index_col_idx}.")

        # Now, move the common_ref column to the first column.
        common_ref_col_idx = None
        for cell in ws[1]:
            if cell.value == "common_ref":
                common_ref_col_idx = cell.column
                break
        if common_ref_col_idx is not None and common_ref_col_idx != 1:
            # Save the values from the current common_ref column.
            common_ref_values = [ws.cell(row=r, column=common_ref_col_idx).value
                                 for r in range(1, ws.max_row + 1)]
            ws.delete_cols(common_ref_col_idx, 1)
            ws.insert_cols(1)
            for r, value in enumerate(common_ref_values, start=1):
                ws.cell(row=r, column=1).value = value
            print(f"Moved 'common_ref' column from position {common_ref_col_idx} to column 1.")

        wb.save(file_path)
        wb.close()
        print("Column reordering complete.")

    @staticmethod
    def _extract_hyperlinks(file_path):
        print("Extracting hyperlinks from:", file_path)
        hyperlinks = {}
        wb = load_workbook(file_path, data_only=False)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            row_number = row[0].row
            hyperlinks[row_number] = {}
            for cell in row:
                if cell.hyperlink:
                    col_idx = cell.column
                    print(f"Hyperlink found at row {row_number}, col {col_idx}: {cell.hyperlink.target}")
                    hyperlinks[row_number][col_idx] = cell.hyperlink.target
        print("Extracted Hyperlinks:", hyperlinks)
        return hyperlinks

    @staticmethod
    def add_original_row_index_to_dataframe(df, file_path):
        print(f"Adding original row indices from file: {file_path}")
        wb = load_workbook(file_path, data_only=False)
        ws = wb.active
        original_row_indices = []
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            row_values = [cell.value for cell in row]
            if any(row_values):
                original_row_indices.append(row[0].row)
        if len(original_row_indices) != len(df):
            raise ValueError(
                f"Mismatch between extracted row indices ({len(original_row_indices)}) and DataFrame rows ({len(df)}). Ensure no extra blank rows in Excel."
            )
        df['original_row_index'] = original_row_indices
        print(f"Assigned original row indices: {original_row_indices}")
        return df

    @staticmethod
    def _save_merged_to_excel(df, output_path):
        temp_file_path = output_path if not os.path.isdir(output_path) else os.path.join(output_path, 'merged_result_temp.xlsx')
        df.to_excel(temp_file_path, index=False, header=True)
        return temp_file_path

    @staticmethod
    def _apply_formatting_and_hyperlinks(file_path, hyperlinks, merged_df):
        wb = load_workbook(file_path)
        ws = wb.active
        print("Applying formatting and hyperlinks to:", file_path)

        duplicate_font = Font(bold=True, color="FF3300")

        # ------------------------------
        # Branch based on whether we're merging 3 or 2 Excel files.
        # ------------------------------
        if 'number_3' in merged_df.columns:
            # --- 3-file merging logic ---
            refno1_col_idx = merged_df.columns.get_loc('number_1') + 1
            refno2_col_idx = merged_df.columns.get_loc('number_2') + 1
            refno3_col_idx = merged_df.columns.get_loc('number_3') + 1

            # Define fill styles for each presence pattern.
            fill_styles = {
                (True, False, False): PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid"),
                (False, True, False): PatternFill(start_color="FFCC66", end_color="FFCC66", fill_type="solid"),
                (False, False, True): PatternFill(start_color="AACCFF", end_color="AACCFF", fill_type="solid"),
                (True, True, False): PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid"),
                (True, False, True): PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid"),
                (False, True, True): PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid"),
                # For rows with all drawing numbers or none, we won't apply any fill.
                (True, True, True): None,
                (False, False, False): None
            }

            # For duplicate highlighting, determine duplicates in each column.
            refno1_duplicates = merged_df['number_1'][merged_df['number_1'].duplicated(keep=False)].tolist()
            refno2_duplicates = merged_df['number_2'][merged_df['number_2'].duplicated(keep=False)].tolist()
            refno3_duplicates = merged_df['number_3'][merged_df['number_3'].duplicated(keep=False)].tolist()

            # Loop over each row (account for header row in Excel, hence i+2).
            for row_idx in range(2, len(merged_df) + 2):
                num1 = ws.cell(row=row_idx, column=refno1_col_idx).value
                num2 = ws.cell(row=row_idx, column=refno2_col_idx).value
                num3 = ws.cell(row=row_idx, column=refno3_col_idx).value

                presence = (bool(num1), bool(num2), bool(num3))
                # Only if the row is incomplete (i.e. not all present or all missing)
                if presence not in [(True, True, True), (False, False, False)]:
                    fill = fill_styles.get(presence)
                    if fill is not None:
                        # Instead of filling the blank cells, fill the cells that hold a drawing number.
                        if num1:
                            ws.cell(row=row_idx, column=refno1_col_idx).fill = fill
                        if num2:
                            ws.cell(row=row_idx, column=refno2_col_idx).fill = fill
                        if num3:
                            ws.cell(row=row_idx, column=refno3_col_idx).fill = fill

                # Apply duplicate formatting (if the cell's value appears more than once).
                if num1 in refno1_duplicates:
                    ws.cell(row=row_idx, column=refno1_col_idx).font = duplicate_font
                if num2 in refno2_duplicates:
                    ws.cell(row=row_idx, column=refno2_col_idx).font = duplicate_font
                if num3 in refno3_duplicates:
                    ws.cell(row=row_idx, column=refno3_col_idx).font = duplicate_font

            # --- Add an Instance column for 3-file merging ---
            instance_mapping = {
                (True, False, False): "PDF Only",
                (False, True, False): "number 2",
                (False, False, True): "number 3",
                (True, True, False): "4: PDF and number_2",
                (True, False, True): "5: PDF and number_3",
                (False, True, True): "No PDF but found on number_2 and number_3",
                (True, True, True): "",
                (False, False, False): "None"
            }
            # Create a new column header for the instance info.
            instance_col_idx = ws.max_column + 1
            ws.cell(row=1, column=instance_col_idx, value="Case")
            for row_idx in range(2, len(merged_df) + 2):
                # Read the cell values again to compute the presence tuple.
                num1 = ws.cell(row=row_idx, column=refno1_col_idx).value
                num2 = ws.cell(row=row_idx, column=refno2_col_idx).value
                num3 = ws.cell(row=row_idx, column=refno3_col_idx).value
                presence = (bool(num1), bool(num2), bool(num3))
                instance_text = instance_mapping.get(presence, "Unknown")
                ws.cell(row=row_idx, column=instance_col_idx, value=instance_text)

        else:
            # --- 2-file merging logic ---
            refno1_col_idx = merged_df.columns.get_loc('number_1') + 1
            refno2_col_idx = merged_df.columns.get_loc('number_2') + 1

            fill_styles_2 = {
                (True, False): PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid"),
                (False, True): PatternFill(start_color="FFCC66", end_color="FFCC66", fill_type="solid"),
                (True, True): None,
                (False, False): None
            }

            refno1_duplicates = merged_df['number_1'][merged_df['number_1'].duplicated(keep=False)].tolist()
            refno2_duplicates = merged_df['number_2'][merged_df['number_2'].duplicated(keep=False)].tolist()

            for row_idx in range(2, len(merged_df) + 2):
                num1 = ws.cell(row=row_idx, column=refno1_col_idx).value
                num2 = ws.cell(row=row_idx, column=refno2_col_idx).value
                presence = (bool(num1), bool(num2))
                if presence not in [(True, True), (False, False)]:
                    fill = fill_styles_2.get(presence)
                    if fill is not None:
                        if num1:
                            ws.cell(row=row_idx, column=refno1_col_idx).fill = fill
                        if num2:
                            ws.cell(row=row_idx, column=refno2_col_idx).fill = fill

                if num1 in refno1_duplicates:
                    ws.cell(row=row_idx, column=refno1_col_idx).font = duplicate_font
                if num2 in refno2_duplicates:
                    ws.cell(row=row_idx, column=refno2_col_idx).font = duplicate_font

            # --- Add an Instance column for 2-file merging ---
            instance_mapping_2 = {
                (True, False): "Instance 1: Only number_1",
                (False, True): "Instance 2: Only number_2",
                (True, True): "Complete",
                (False, False): "None"
            }
            instance_col_idx = ws.max_column + 1
            ws.cell(row=1, column=instance_col_idx, value="Instance")
            for row_idx in range(2, len(merged_df) + 2):
                num1 = ws.cell(row=row_idx, column=refno1_col_idx).value
                num2 = ws.cell(row=row_idx, column=refno2_col_idx).value
                presence = (bool(num1), bool(num2))
                instance_text = instance_mapping_2.get(presence, "Unknown")
                ws.cell(row=row_idx, column=instance_col_idx, value=instance_text)

        # ------------------------------
        # Process hyperlinks (unchanged from your original code)
        for original_row_index, columns in hyperlinks.items():
            new_row = merged_df[merged_df['original_row_index'] == original_row_index]
            if new_row.empty:
                print(f"No matching row for original_row_index: {original_row_index}")
                continue
            new_row_idx = new_row.index[0] + 2
            for col_idx, hyperlink in columns.items():
                try:
                    ws.cell(row=new_row_idx, column=col_idx).hyperlink = hyperlink
                    ws.cell(row=new_row_idx, column=col_idx).style = "Hyperlink"
                    print(f"Applied hyperlink at row {new_row_idx}, col {col_idx}, link: {hyperlink}")
                except Exception as e:
                    print(f"Error applying hyperlink for row {new_row_idx}, col {col_idx}: {e}")

        wb.save(file_path)
        wb.close()

    @staticmethod
    def tokenize_with_indices(text):
        tokens = []
        for match in re.finditer(r'[^\s]+', text):
            tokens.append((match.group(), match.start()))
        return tokens

    @staticmethod
    def dp_align_tokens(tokens1, tokens2):
        token_strs1 = [t[0] for t in tokens1]
        token_strs2 = [t[0] for t in tokens2]
        n, m = len(token_strs1), len(token_strs2)
        dp = [[0] * (m + 1) for _ in range(n + 1)]
        backtrack = [[None] * (m + 1) for _ in range(n + 1)]
        for i in range(1, n + 1):
            dp[i][0] = i
            backtrack[i][0] = 'UP'
        for j in range(1, m + 1):
            dp[0][j] = j
            backtrack[0][j] = 'LEFT'
        for i in range(1, n + 1):
            for j in range(1, m + 1):
                similarity = SequenceMatcher(None, token_strs1[i - 1], token_strs2[j - 1]).ratio()
                if token_strs1[i - 1] == token_strs2[j - 1]:
                    replace_cost = dp[i - 1][j - 1]
                    flag = "EXACT"
                elif token_strs1[i - 1].lower() == token_strs2[j - 1].lower():
                    replace_cost = dp[i - 1][j - 1] + 0.5
                    flag = "CASE_ONLY"
                elif similarity >= 0.8:
                    replace_cost = dp[i - 1][j - 1] + (1 - similarity)
                    flag = "CHAR_LEVEL"
                elif similarity >= 0.4:
                    replace_cost = dp[i - 1][j - 1] + 2
                    flag = "CHAR_LEVEL"
                else:
                    replace_cost = float('inf')
                delete_cost = dp[i - 1][j] + 1
                insert_cost = dp[i][j - 1] + 1
                dp[i][j] = min(replace_cost, delete_cost, insert_cost)
                if dp[i][j] == replace_cost:
                    backtrack[i][j] = 'DIAG'
                elif dp[i][j] == delete_cost:
                    backtrack[i][j] = 'UP'
                else:
                    backtrack[i][j] = 'LEFT'
        aligned_tokens1, aligned_tokens2, flags = [], [], []
        i, j = n, m
        while i > 0 or j > 0:
            if i > 0 and j > 0 and backtrack[i][j] == 'DIAG':
                aligned_tokens1.append(tokens1[i - 1])
                aligned_tokens2.append(tokens2[j - 1])
                if token_strs1[i - 1] == token_strs2[j - 1]:
                    flags.append("EXACT")
                elif token_strs1[i - 1].lower() == token_strs2[j - 1].lower():
                    flags.append("CASE_ONLY")
                else:
                    flags.append("CHAR_LEVEL")
                i -= 1
                j -= 1
            elif i > 0 and (j == 0 or backtrack[i][j] == 'UP'):
                aligned_tokens1.append(tokens1[i - 1])
                aligned_tokens2.append((None, None))
                flags.append("MISSING_2")
                i -= 1
            elif j > 0 and (i == 0 or backtrack[i][j] == 'LEFT'):
                aligned_tokens1.append((None, None))
                aligned_tokens2.append(tokens2[j - 1])
                flags.append("MISSING_1")
                j -= 1
        return aligned_tokens1[::-1], aligned_tokens2[::-1], flags[::-1]

    @staticmethod
    def create_rich_text(original_text, aligned_tokens, flags):
        rich_text = CellRichText()
        last_index = 0
        for (token, idx), flag in zip(aligned_tokens, flags):
            if token is None:
                continue
            if idx > last_index:
                rich_text.append(original_text[last_index:idx])
            inline_font = InlineFont(rFont="Calibri", sz=11)
            if flag == "EXACT":
                pass
            elif flag == "CASE_ONLY":
                inline_font.color = "808080"
            elif flag == "CHAR_LEVEL":
                inline_font.color = "FFA500"
            elif flag in ["MISSING_1", "MISSING_2"]:
                inline_font.color = "FF0000"
            rich_text.append(TextBlock(inline_font, token))
            last_index = idx + len(token)
        if last_index < len(original_text):
            rich_text.append(original_text[last_index:])
        return rich_text

    @staticmethod
    def get_ws_column_index(ws, header_name):
        header_name = header_name.strip().lower()
        print(f"Looking for header '{header_name}' in row 1:")
        for cell in ws[1]:
            if isinstance(cell.value, str):
                print(f"  Found header '{cell.value.strip().lower()}' at column {cell.column}")
                if cell.value.strip().lower() == header_name:
                    return cell.column  # openpyxl uses 1-based indexing
        return None

    @staticmethod
    def apply_title_highlighting(file_path, merged_df, title_col1, title_col2, reorder=True, update_baseline=True):
        wb = load_workbook(file_path, rich_text=True)
        ws = wb.active

        # Get current column indices by scanning the header row
        col_idx1 = ExcelMerger.get_ws_column_index(ws, title_col1)
        col_idx2 = ExcelMerger.get_ws_column_index(ws, title_col2)
        if col_idx1 is None or col_idx2 is None:
            print(f"Could not find header {title_col1} or {title_col2} in the worksheet.")
            wb.close()
            return

        print(f"Applying title highlighting on columns '{title_col1}' and '{title_col2}':")
        for i, row in merged_df.iterrows():
            excel_row = i + 2  # account for header row
            baseline_text = str(row.get(title_col1, ""))
            other_text = str(row.get(title_col2, ""))
            tokens1 = ExcelMerger.tokenize_with_indices(baseline_text)
            tokens2 = ExcelMerger.tokenize_with_indices(other_text)
            aligned_tokens1, aligned_tokens2, flags = ExcelMerger.dp_align_tokens(tokens1, tokens2)
            rich_text1 = ExcelMerger.create_rich_text(baseline_text, aligned_tokens1, flags)
            rich_text2 = ExcelMerger.create_rich_text(other_text, aligned_tokens2, flags)
            # Update baseline column only if update_baseline is True.
            if update_baseline:
                ws.cell(row=excel_row, column=col_idx1).value = rich_text1
            # Always update the second column.
            ws.cell(row=excel_row, column=col_idx2).value = rich_text2

        if reorder:
            ExcelMerger.reorder_columns(file_path)

        wb.save(file_path)
        wb.close()
        print("Title highlighting applied and workbook saved:", file_path)


# class TitleComparison:
#     """Handles title comparison logic and generates a Word report."""
#
#     @staticmethod
#     def create_report(df, title_column1, title_column2, output_path, include_excel3=False):
#         """
#         Create a Word document highlighting differences between title columns.
#         The first column will be "common_ref", followed by "title_excel1", "title_excel2",
#         and, if include_excel3 is True, "title_excel3". In this report, title_excel1 is used
#         as the baseline; title_excel2 and (optionally) title_excel3 will be compared to title_excel1
#         and highlighted.
#         """
#         # Expected fixed column names in the merged DataFrame:
#         headers = ["common_ref", "title_excel1", "title_excel2"]
#         if include_excel3:
#             headers.append("title_excel3")
#         print("Creating report using headers:", headers)
#         doc = Document()
#         doc.add_heading('Title Differences Report', level=1)
#         # For summary, count mismatches between title_excel1 and title_excel2 only.
#         mismatches = df['title_excel1'] != df['title_excel2']
#         TitleComparison._add_summary(doc, len(df), len(df[mismatches]))
#         num_cols = len(headers)
#         table = doc.add_table(rows=1, cols=num_cols)
#         table.style = 'Table Grid'
#         header_cells = table.rows[0].cells
#         header_cells[0].text = "Common Ref"
#         header_cells[1].text = "title_excel1"
#         header_cells[2].text = "title_excel2"
#         if include_excel3:
#             header_cells[3].text = "title_excel3"
#         for cell in header_cells:
#             for paragraph in cell.paragraphs:
#                 for run in paragraph.runs:
#                     run.font.name = 'Helvetica'
#                     run.font.size = Pt(9)
#                     run.bold = True
#         # Populate rows
#         for _, row in df.iterrows():
#             row_cells = table.add_row().cells
#             row_cells[0].text = str(row.get("common_ref", ""))
#             # Baseline title (Excel1) is added without highlighting.
#             row_cells[1].text = str(row.get("title_excel1", ""))
#             # For title_excel2, compare it against baseline (title_excel1)
#             baseline = str(row.get("title_excel1", ""))
#             text2 = str(row.get("title_excel2", ""))
#             para2 = row_cells[2].paragraphs[0]
#             clear_paragraph(para2)
#             tokens_base = TitleComparison.tokenize_with_indices(baseline)
#             tokens_2 = TitleComparison.tokenize_with_indices(text2)
#             aligned_tokens2, _, flags2 = TitleComparison.dp_align_tokens(tokens_2, tokens_base)
#             TitleComparison.reconstruct_text_with_flags(para2, aligned_tokens2, tokens_base, flags2)
#             if include_excel3:
#                 text3 = str(row.get("title_excel3", ""))
#                 para3 = row_cells[3].paragraphs[0]
#                 clear_paragraph(para3)
#                 tokens_3 = TitleComparison.tokenize_with_indices(text3)
#                 aligned_tokens3, _, flags3 = TitleComparison.dp_align_tokens(tokens_3, tokens_base)
#                 TitleComparison.reconstruct_text_with_flags(para3, aligned_tokens3, tokens_base, flags3)
#         doc.save(output_path)
#         print("Report saved at:", output_path)
#
#     @staticmethod
#     def _add_summary(doc, total_titles, mismatched_count):
#         doc.add_heading('Summary of Title Comparison', level=2)
#         total_para = doc.add_paragraph(f"Total Titles Compared: {total_titles}")
#         total_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
#         total_run = total_para.runs[0]
#         total_run.font.name = 'Arial'
#         total_run.font.size = Pt(10)
#         mismatch_para = doc.add_paragraph(f"Titles with Differences: {mismatched_count}")
#         mismatch_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
#         mismatch_run = mismatch_para.runs[0]
#         mismatch_run.font.name = 'Arial'
#         mismatch_run.font.size = Pt(10)
#
#     @staticmethod
#     def tokenize_with_indices(text):
#         tokens = []
#         # Use a regex that matches either whitespace or non-whitespace sequences.
#         for match in re.finditer(r'\s+|\S+', text):
#             tokens.append((match.group(), match.start()))
#         print("DEBUG (Report): Tokenized Text with Indices:", tokens)
#         return tokens
#
#     @staticmethod
#     def dp_align_tokens(tokens1, tokens2):
#         token_strs1 = [t[0] for t in tokens1]
#         token_strs2 = [t[0] for t in tokens2]
#         n, m = len(token_strs1), len(token_strs2)
#         dp = [[0] * (m + 1) for _ in range(n + 1)]
#         backtrack = [[None] * (m + 1) for _ in range(n + 1)]
#         for i in range(1, n + 1):
#             dp[i][0] = i
#             backtrack[i][0] = 'UP'
#         for j in range(1, m + 1):
#             dp[0][j] = j
#             backtrack[0][j] = 'LEFT'
#         for i in range(1, n + 1):
#             for j in range(1, m + 1):
#                 similarity = SequenceMatcher(None, token_strs1[i - 1], token_strs2[j - 1]).ratio()
#                 if token_strs1[i - 1] == token_strs2[j - 1]:
#                     replace_cost = dp[i - 1][j - 1]
#                     flag = "EXACT"
#                 elif token_strs1[i - 1].lower() == token_strs2[j - 1].lower():
#                     replace_cost = dp[i - 1][j - 1] + 0.5
#                     flag = "CASE_ONLY"
#                 elif similarity >= 0.8:
#                     replace_cost = dp[i - 1][j - 1] + (1 - similarity)
#                     flag = "CHAR_LEVEL"
#                 elif similarity >= 0.4:
#                     replace_cost = dp[i - 1][j - 1] + 2
#                     flag = "CHAR_LEVEL"
#                 else:
#                     replace_cost = float('inf')
#                 delete_cost = dp[i - 1][j] + 1
#                 insert_cost = dp[i][j - 1] + 1
#                 dp[i][j] = min(replace_cost, delete_cost, insert_cost)
#                 if dp[i][j] == replace_cost:
#                     backtrack[i][j] = 'DIAG'
#                 elif dp[i][j] == delete_cost:
#                     backtrack[i][j] = 'UP'
#                 else:
#                     backtrack[i][j] = 'LEFT'
#         aligned_tokens1, aligned_tokens2, flags = [], [], []
#         i, j = n, m
#         while i > 0 or j > 0:
#             if i > 0 and j > 0 and backtrack[i][j] == 'DIAG':
#                 aligned_tokens1.append(tokens1[i - 1])
#                 aligned_tokens2.append(tokens2[j - 1])
#                 if token_strs1[i - 1] == token_strs2[j - 1]:
#                     flags.append("EXACT")
#                 elif token_strs1[i - 1].lower() == token_strs2[j - 1].lower():
#                     flags.append("CASE_ONLY")
#                 else:
#                     flags.append("CHAR_LEVEL")
#                 i -= 1
#                 j -= 1
#             elif i > 0 and (j == 0 or backtrack[i][j] == 'UP'):
#                 aligned_tokens1.append(tokens1[i - 1])
#                 aligned_tokens2.append((None, None))
#                 flags.append("MISSING_2")
#                 i -= 1
#             elif j > 0 and (i == 0 or backtrack[i][j] == 'LEFT'):
#                 aligned_tokens1.append((None, None))
#                 aligned_tokens2.append(tokens2[j - 1])
#                 flags.append("MISSING_1")
#                 j -= 1
#         return aligned_tokens1[::-1], aligned_tokens2[::-1], flags[::-1]
#
#     @staticmethod
#     def reconstruct_text_with_flags(paragraph, tokens1, tokens2, flags):
#         for (token1, index1), (token2, index2), flag in zip(tokens1, tokens2, flags):
#             if token1 is None:
#                 continue
#             if index1 > 0:
#                 paragraph.add_run(" ")
#             if flag == "EXACT":
#                 paragraph.add_run(token1)
#             elif flag == "CASE_ONLY":
#                 run = paragraph.add_run(token1)
#                 run.font.color.rgb = RGBColor(128, 128, 128)
#             elif flag == "CHAR_LEVEL":
#                 matcher = SequenceMatcher(None, token1, token2)
#                 for op, i1, i2, j1, j2 in matcher.get_opcodes():
#                     text_part = token1[i1:i2]
#                     if op == "equal":
#                         paragraph.add_run(text_part)
#                     else:
#                         diff_run = paragraph.add_run(text_part)
#                         diff_run.font.color.rgb = RGBColor(255, 165, 0)
#             elif flag in ["MISSING_1", "MISSING_2"]:
#                 run = paragraph.add_run(token1)
#                 run.font.color.rgb = RGBColor(255, 0, 0)
#         return paragraph
#

# Custom main window class that supports drag-and-drop.

class CTkDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


class MergerGUI:
    """Handles the GUI for the Excel merger with drag-and-drop file selection and theme toggle."""

    def __init__(self, master=None):

        #set_custom_theme("dark")  # or "light" if you prefer the light theme

        ctk.set_appearance_mode("dark")  # Set dark mode at startup


        # Use our custom CTkDnD main window for drag-and-drop support.
        self.mergerApp = CTkDnD() if master is None else ctk.CTkToplevel(master)
        self.mergerApp.title("[No Name Yet]")

        # File paths for three Excel files and the output file.
        self.excel1_path = tk.StringVar()
        self.excel2_path = tk.StringVar()
        self.excel3_path = tk.StringVar()
        self.output_path = tk.StringVar()

        # Header selections for each file (via drop-down lists).
        self.ref_column1 = tk.StringVar()
        self.title_column1 = tk.StringVar()
        self.ref_column2 = tk.StringVar()
        self.title_column2 = tk.StringVar()
        self.ref_column3 = tk.StringVar()
        self.title_column3 = tk.StringVar()

        # Boolean variables for report options and comparing Excel3 title.
        self.generate_report = tk.BooleanVar(value=False)
        self.generate_word_report = tk.BooleanVar(value=False)
        self.compare_excel3_title = tk.BooleanVar(value=False)

        # Boolean variable for theme mode; True = dark mode.
        self.theme_mode = tk.BooleanVar(value=True)

        self.excel1_headers = []
        self.excel2_headers = []
        self.excel3_headers = []

        self._build_gui()

    def _build_gui(self):
        self.mergerApp.grid_rowconfigure(0, weight=1)
        self.mergerApp.grid_columnconfigure(0, weight=1)
        self.mergerApp.grid_columnconfigure(1, weight=1)
        self.mergerApp.grid_columnconfigure(2, weight=1)

        # Create three frames for Excel 1, 2, and 3.
        self.excel1_frame = ctk.CTkFrame(self.mergerApp)
        self.excel1_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.excel2_frame = ctk.CTkFrame(self.mergerApp)
        self.excel2_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.excel3_frame = ctk.CTkFrame(self.mergerApp)
        self.excel3_frame.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")

        for frame in (self.excel1_frame, self.excel2_frame, self.excel3_frame):
            frame.grid_rowconfigure(0, weight=0)
            frame.grid_columnconfigure(0, weight=0)
            frame.grid_columnconfigure(1, weight=1)

        self._build_excel1_section(self.excel1_frame)
        self._build_excel2_section(self.excel2_frame)
        self._build_excel3_section(self.excel3_frame)

        # Create controls frame.
        self.controls_frame = ctk.CTkFrame(self.mergerApp)
        self.controls_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        self._build_controls(self.controls_frame)

    def _build_excel1_section(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12
        self.excel1_button = ctk.CTkButton(parent_frame,
            text="\n➕\n\nSelect File or\nDrag & Drop Here",
            command=self._browse_excel1,
            border_width=3,
            fg_color="transparent",
            hover_color=("#D6D6D6", "#505050"),  # Light and dark hover color
            text_color=("#333333", "#FFFFFF"),
            corner_radius=10,
            width=200,
            height=150)
        self.excel1_button.grid(row=0, column=0, columnspan=2, padx=33, pady=33, sticky="ew")
        # Enable drag and drop on Excel1 button.
        self.excel1_button.drop_target_register(DND_ALL)
        self.excel1_button.dnd_bind('<<Drop>>', self.drop_excel1)
        ctk.CTkLabel(parent_frame, text="Reference Column:", font=(font_name, font_size)).grid(
            row=1, column=0, padx=5, pady=2, sticky="e")
        self.ref_option_menu1 = ctk.CTkOptionMenu(parent_frame, variable=self.ref_column1, values=[])
        self.ref_option_menu1.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkCheckBox(parent_frame, text="Compare Title",
                        variable=self.generate_report, command=self._toggle_title_entries).grid(
            row=2, column=0, columnspan=2, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(parent_frame, text="Drawing Title:", font=(font_name, font_size)).grid(
            row=3, column=0, padx=5, pady=2, sticky="e")
        self.title_option_menu1 = ctk.CTkOptionMenu(parent_frame, variable=self.title_column1, values=[], state="disabled")
        self.title_option_menu1.grid(row=3, column=1, padx=5, pady=2, sticky="ew")

    def _build_excel2_section(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12
        self.excel2_button = ctk.CTkButton(parent_frame,
            text="\n➕\n\nSelect File or\nDrag & Drop Here",
            command=self._browse_excel2,
            border_width=3,
            fg_color="transparent",
            hover_color=("#D6D6D6", "#505050"),  # Light and dark hover color
            text_color=("#333333", "#FFFFFF"),
            corner_radius=10,
            width=200,
            height=150)
        self.excel2_button.grid(row=0, column=0, columnspan=2, padx=33, pady=33, sticky="ew")
        self.excel2_button.drop_target_register(DND_ALL)
        self.excel2_button.dnd_bind('<<Drop>>', self.drop_excel2)
        ctk.CTkLabel(parent_frame, text="Reference Column:", font=(font_name, font_size)).grid(
            row=1, column=0, padx=5, pady=2, sticky="e")
        self.ref_option_menu2 = ctk.CTkOptionMenu(parent_frame, variable=self.ref_column2, values=[])
        self.ref_option_menu2.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        # Spacer row for alignment.
        ctk.CTkLabel(parent_frame, text="", font=(font_name, font_size)).grid(
            row=2, column=0, padx=5, pady=2, sticky="e")
        ctk.CTkLabel(parent_frame, text="Drawing Title:", font=(font_name, font_size)).grid(
            row=3, column=0, padx=5, pady=2, sticky="e")
        self.title_option_menu2 = ctk.CTkOptionMenu(parent_frame, variable=self.title_column2, values=[], state="disabled")
        self.title_option_menu2.grid(row=3, column=1, padx=5, pady=2, sticky="ew")

    def _build_excel3_section(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12
        self.excel3_button = ctk.CTkButton(parent_frame,
            text="\n➕\n\nSelect File or\nDrag & Drop Here",
            command=self._browse_excel3,
            border_width=3,
            fg_color="transparent",
            hover_color=("#D6D6D6", "#505050"),  # Light and dark hover color
            text_color=("#333333", "#FFFFFF"),
            corner_radius=10,
            width=200,
            height=150)
        self.excel3_button.grid(row=0, column=0, columnspan=2, padx=33, pady=33, sticky="ew")
        self.excel3_button.drop_target_register(DND_ALL)
        self.excel3_button.dnd_bind('<<Drop>>', self.drop_excel3)
        ctk.CTkLabel(parent_frame, text="Reference Column:", font=(font_name, font_size)).grid(
            row=1, column=0, padx=5, pady=2, sticky="e")
        self.ref_option_menu3 = ctk.CTkOptionMenu(parent_frame, variable=self.ref_column3, values=[])
        self.ref_option_menu3.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        # Place the compare title checkbox above the title dropdown.
        ctk.CTkCheckBox(parent_frame, text="Compare Title",
                        variable=self.compare_excel3_title, command=self._toggle_title_entries).grid(
            row=2, column=0, columnspan=2, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(parent_frame, text="Drawing Title:", font=(font_name, font_size)).grid(
            row=3, column=0, padx=5, pady=2, sticky="e")
        self.title_option_menu3 = ctk.CTkOptionMenu(parent_frame, variable=self.title_column3, values=[], state="disabled")
        self.title_option_menu3.grid(row=3, column=1, padx=5, pady=2, sticky="ew")

    def _build_controls(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12
        ctk.CTkLabel(parent_frame, text="Output Path:", font=(font_name, font_size)).grid(
            row=0, column=0, padx=5, pady=2, sticky="e")
        ctk.CTkEntry(parent_frame, textvariable=self.output_path, width=300).grid(
            row=0, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkButton(parent_frame, text="Use Excel 1 Path", command=self._use_excel1_path).grid(
            row=0, column=2, padx=5, pady=2)
        # Theme toggle switch (no label) placed at the bottom-right of the controls frame.
        self.theme_switch = ctk.CTkSwitch(parent_frame, text="", variable=self.theme_mode,
                                          command=self.toggle_theme, switch_width=20, switch_height=10)
        # Use place to anchor it at the bottom-right corner of the parent frame.
        self.theme_switch.place(relx=1.0, rely=1.0, anchor="se")
        ctk.CTkButton(parent_frame, text="Start Merge", command=self._start_merge).grid(
            row=2, column=0, columnspan=3, pady=10)
        parent_frame.grid_columnconfigure(1, weight=1)

    def toggle_theme(self):
        if self.theme_mode.get():
            ctk.set_appearance_mode("dark")
            print("Theme set to dark mode")
        else:
            ctk.set_appearance_mode("light")
            print("Theme set to light mode")

    # --- Drag and Drop Handlers ---
    def drop_excel1(self, event):
        file_path = event.data.replace("{", "").replace("}", "")
        self.excel1_path.set(file_path)
        self._load_excel1_headers(file_path)
        import os
        filename = os.path.basename(file_path)
        self.excel1_button.configure(text=filename, fg_color="#217346")

    def drop_excel2(self, event):
        file_path = event.data.replace("{", "").replace("}", "")
        self.excel2_path.set(file_path)
        self._load_excel2_headers(file_path)
        import os
        filename = os.path.basename(file_path)
        self.excel2_button.configure(text=filename, fg_color="#217346")

    def drop_excel3(self, event):
        file_path = event.data.replace("{", "").replace("}", "")
        self.excel3_path.set(file_path)
        self._load_excel3_headers(file_path)
        import os
        filename = os.path.basename(file_path)
        self.excel3_button.configure(text=filename, fg_color="#217346")

    def _browse_excel1(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")],
                                               title="Select Excel File 1")
        if file_path:
            self.excel1_path.set(file_path)
            self._load_excel1_headers(file_path)
            import os
            filename = os.path.basename(file_path)
            self.excel1_button.configure(text=filename, fg_color="#217346")

    def _browse_excel2(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")],
                                               title="Select Excel File 2")
        if file_path:
            self.excel2_path.set(file_path)
            self._load_excel2_headers(file_path)
            import os
            filename = os.path.basename(file_path)
            self.excel2_button.configure(text=filename, fg_color="#217346")

    def _browse_excel3(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")],
                                               title="Select Excel File 3")
        if file_path:
            self.excel3_path.set(file_path)
            self._load_excel3_headers(file_path)
            import os
            filename = os.path.basename(file_path)
            self.excel3_button.configure(text=filename, fg_color="#217346")

    def _load_excel1_headers(self, file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
            headers = list(df.columns)
            self.excel1_headers = headers
            self.ref_option_menu1.configure(values=headers)
            self.title_option_menu1.configure(values=headers)
            self.ref_column1.set(auto_select_header(headers, ["drawing", "sheet", "ref", "number"]))
            self.title_column1.set(auto_select_header(headers, ["title"]))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load headers from Excel File 1: {e}")

    def _load_excel2_headers(self, file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
            headers = list(df.columns)
            self.excel2_headers = headers
            self.ref_option_menu2.configure(values=headers)
            self.title_option_menu2.configure(values=headers)
            self.ref_column2.set(auto_select_header(headers, ["drawing", "sheet", "ref", "number"]))
            self.title_column2.set(auto_select_header(headers, ["title"]))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load headers from Excel File 2: {e}")

    def _load_excel3_headers(self, file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
            headers = list(df.columns)
            self.excel3_headers = headers
            self.ref_option_menu3.configure(values=headers)
            self.title_option_menu3.configure(values=headers)
            self.ref_column3.set(auto_select_header(headers, ["drawing", "sheet", "ref", "number"]))
            self.title_column3.set(auto_select_header(headers, ["title"]))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load headers from Excel File 3: {e}")

    def _use_excel1_path(self):
        excel1 = self.excel1_path.get()
        if excel1:
            directory, file_name = os.path.split(excel1)
            name, ext = os.path.splitext(file_name)
            self.output_path.set(os.path.join(directory, f"{name}_merged{ext}"))

    def _toggle_title_entries(self):
        state_12 = "normal" if self.generate_report.get() else "disabled"
        self.title_option_menu1.configure(state=state_12)
        self.title_option_menu2.configure(state=state_12)
        state_3 = "normal" if self.compare_excel3_title.get() else "disabled"
        self.title_option_menu3.configure(state=state_3)

    def _start_merge(self):
        excel1_path = self.excel1_path.get()
        excel2_path = self.excel2_path.get()
        excel3_path = self.excel3_path.get()
        ref_column1 = self.ref_column1.get()
        ref_column2 = self.ref_column2.get()
        ref_column3 = self.ref_column3.get() if self.excel3_path.get().strip() else None

        output_path = self.output_path.get().strip()
        if not output_path:
            messagebox.showerror("Error", "Please provide a valid output path.")
            return
        if os.path.exists(output_path):
            directory, file_name = os.path.split(output_path)
            base, ext = os.path.splitext(file_name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(directory, f"{base}_{timestamp}{ext}")

        try:
            if self.generate_report.get():
                title_col1 = self.title_column1.get()
                title_col2 = self.title_column2.get()
                title_col3 = self.title_column3.get() if self.excel3_path.get().strip() else None
            else:
                title_col1 = title_col2 = title_col3 = None

            if excel3_path.strip():
                merged_file_path, merged_df = ExcelMerger.merge_3_excels(
                    excel1_path, excel2_path, excel3_path,
                    ref_column1, ref_column2, ref_column3,
                    output_path,
                    title_column1=title_col1, title_column2=title_col2, title_column3=title_col3,
                    compare_excel3=self.compare_excel3_title.get()
                )
            else:
                merged_file_path, merged_df = ExcelMerger.merge_excels(
                    excel1_path, excel2_path,
                    ref_column1, ref_column2,
                    output_path,
                    title_column1=title_col1, title_column2=title_col2
                )
            messagebox.showinfo("Success", f"Merged file saved at {merged_file_path}")

            ### TitleComparison is disabled
            # if self.generate_report.get() and self.generate_word_report.get():
            #     report_path = os.path.splitext(merged_file_path)[0] + "-TitleComparison.docx"
            #     include_excel3 = self.compare_excel3_title.get() and bool(excel3_path.strip())
            #     TitleComparison.create_report(merged_df, 'title_excel1', 'title_excel2', report_path, include_excel3)
            #     messagebox.showinfo("Success", f"Title comparison report saved at {report_path}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run(self):
        if isinstance(self.mergerApp, ctk.CTk):
            self.mergerApp.mainloop()

if __name__ == "__main__":
    app = MergerGUI()
    app.run()
