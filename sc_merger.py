import os
import re
import tkinter as tk
from difflib import SequenceMatcher
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

import unicodedata


class ExcelMerger:
    """Handles Excel merging logic, including conditional formatting, hyperlink handling,
    and title rich-text highlighting (using tokenization and dynamic programming)."""

    @staticmethod
    def merge_excels(excel1_path, excel2_path, ref_column1, ref_column2, output_path,
                     title_column1=None, title_column2=None):
        """
        Merge two Excel files while retaining hyperlinks and applying formatting.
        The selected headers will be renamed:
          - The reference column in Excel 1 becomes "number_1" and in Excel 2 "number_2"
          - The title column in Excel 1 becomes "title_excel1" and in Excel 2 "title_excel2"
        Optionally, apply title rich-text highlighting if title_column1 and title_column2 are provided.
        """
        # Read Excel files
        excel1 = pd.read_excel(excel1_path, engine='openpyxl', dtype=str).fillna("")
        excel2 = pd.read_excel(excel2_path, engine='openpyxl', dtype=str).fillna("")

        # Ensure the reference columns exist
        if ref_column1 not in excel1.columns or ref_column2 not in excel2.columns:
            raise KeyError("Reference columns not found in one or both Excel files.")

        # Add original_row_index to track rows (based on non-empty rows in excel1)
        excel1 = ExcelMerger.add_original_row_index_to_dataframe(excel1, excel1_path)

        # Extract hyperlinks from the original file (from Excel1)
        hyperlinks = ExcelMerger._extract_hyperlinks(excel1_path)

        # Prepare data for merging:
        # First, compute a count for duplicate reference values
        excel1['refno_count'] = excel1.groupby(ref_column1).cumcount()
        excel2['refno_count'] = excel2.groupby(ref_column2).cumcount()

        # Rename the reference columns to our desired fixed names:
        excel1 = excel1.rename(columns={ref_column1: 'number_1'})
        excel2 = excel2.rename(columns={ref_column2: 'number_2'})

        # Optionally rename the title columns if provided.
        if title_column1:
            excel1 = excel1.rename(columns={title_column1: 'title_excel1'})
        if title_column2:
            excel2 = excel2.rename(columns={title_column2: 'title_excel2'})

        # Now merge the DataFrames. (Since we’ve renamed the keys, we merge on 'number_1' and 'number_2'.)
        merged_df = pd.merge(
            excel1, excel2,
            left_on=['number_1', 'refno_count'],
            right_on=['number_2', 'refno_count'],
            how='outer'
        ).drop(columns=['refno_count']).fillna("")

        merged_df['original_row_index'] = pd.to_numeric(
            merged_df['original_row_index'], errors='coerce').fillna(0).astype(int)

        # Debug: print merged columns to verify renaming.
        print("Merged DataFrame columns:")
        print(merged_df.columns.tolist())

        # Save the merged file without extra formatting.
        temp_file_path = ExcelMerger._save_merged_to_excel(merged_df, output_path)

        # Apply formatting and reapply hyperlinks.
        ExcelMerger._apply_formatting_and_hyperlinks(temp_file_path, hyperlinks, merged_df)

        # If title columns were provided, apply rich-text title highlighting in Excel.
        # (Now we know the fixed names are "title_excel1" and "title_excel2".)
        if title_column1 and title_column2:
            ExcelMerger.apply_title_highlighting(temp_file_path, merged_df, 'title_excel1', 'title_excel2')

        return temp_file_path, merged_df

    @staticmethod
    def _extract_hyperlinks(file_path):
        print("Extracting hyperlinks from:", file_path)
        hyperlinks = {}
        wb = load_workbook(file_path, data_only=False)
        ws = wb.active

        # Loop through all rows and columns (skip header)
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
                f"Mismatch between extracted row indices ({len(original_row_indices)}) "
                f"and DataFrame rows ({len(df)}). Ensure no extra blank rows in Excel."
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

        fill_missing_refno1 = PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid")
        fill_missing_refno2 = PatternFill(start_color="FFCC66", end_color="FFCC66", fill_type="solid")
        duplicate_font = Font(bold=True, color="FF3300")

        # Our reference keys are now 'number_1' and 'number_2'
        refno1_col_idx = merged_df.columns.get_loc('number_1') + 1
        refno2_col_idx = merged_df.columns.get_loc('number_2') + 1

        refno1_duplicates = merged_df['number_1'][merged_df['number_1'].duplicated(keep=False)].tolist()
        refno2_duplicates = merged_df['number_2'][merged_df['number_2'].duplicated(keep=False)].tolist()

        for row_idx in range(2, len(merged_df) + 2):
            refno1_value = ws.cell(row=row_idx, column=refno1_col_idx).value
            refno2_value = ws.cell(row=row_idx, column=refno2_col_idx).value

            if refno1_value and not refno2_value:
                ws.cell(row=row_idx, column=refno1_col_idx).fill = fill_missing_refno1
            if refno2_value and not refno1_value:
                ws.cell(row=row_idx, column=refno2_col_idx).fill = fill_missing_refno2

            if refno1_value in refno1_duplicates:
                ws.cell(row=row_idx, column=refno1_col_idx).font = duplicate_font
            if refno2_value in refno2_duplicates:
                ws.cell(row=row_idx, column=refno2_col_idx).font = duplicate_font

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
                inline_font.color = "808080"  # Gray
            elif flag == "CHAR_LEVEL":
                inline_font.color = "FFA500"  # Orange
            elif flag in ["MISSING_1", "MISSING_2"]:
                inline_font.color = "FF0000"  # Red
            rich_text.append(TextBlock(inline_font, token))
            last_index = idx + len(token)
        if last_index < len(original_text):
            rich_text.append(original_text[last_index:])
        return rich_text

    @staticmethod
    def apply_title_highlighting(file_path, merged_df, title_col1, title_col2):
        wb = load_workbook(file_path, rich_text=True)
        ws = wb.active

        # Get Excel column indices (1-based)
        col_idx1 = merged_df.columns.get_loc(title_col1) + 1
        col_idx2 = merged_df.columns.get_loc(title_col2) + 1

        # Debug: print out which columns (and their indices) are used for title highlighting.
        print("Applying title highlighting:")
        print(f"Title column 1: {title_col1} (Column index: {col_idx1})")
        print(f"Title column 2: {title_col2} (Column index: {col_idx2})")

        for i, row in merged_df.iterrows():
            excel_row = i + 2  # Account for header row in Excel
            title1 = str(row.get(title_col1, ""))
            title2 = str(row.get(title_col2, ""))

            # Debug: print the titles being compared on each row
            print(f"Row {excel_row}: Comparing Title1: '{title1}' with Title2: '{title2}'")

            tokens1 = ExcelMerger.tokenize_with_indices(title1)
            tokens2 = ExcelMerger.tokenize_with_indices(title2)
            aligned_tokens1, aligned_tokens2, flags = ExcelMerger.dp_align_tokens(tokens1, tokens2)

            rich_text1 = ExcelMerger.create_rich_text(title1, aligned_tokens1, flags)
            rich_text2 = ExcelMerger.create_rich_text(title2, aligned_tokens2, flags)

            ws.cell(row=excel_row, column=col_idx1).value = rich_text1
            ws.cell(row=excel_row, column=col_idx2).value = rich_text2

        wb.save(file_path)
        wb.close()
        print("Title highlighting applied and workbook saved:", file_path)


class TitleComparison:
    """Handles title comparison logic and generates a Word report."""

    @staticmethod
    def create_report(df, title_column1, title_column2, output_path):
        # With our new renaming, we expect the merged DataFrame to have
        # 'title_excel1' and 'title_excel2' for the title columns.
        print("Creating report using columns: title_excel1 and title_excel2")
        doc = Document()
        doc.add_heading('Title Differences Report', level=1)

        mismatches = df['title_excel1'] != df['title_excel2']
        TitleComparison._add_summary(doc, len(df), len(df[mismatches]))

        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = "Drawing Number"
        header_cells[1].text = "title_excel1"
        header_cells[2].text = "title_excel2"

        for cell in header_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Helvetica'
                    run.font.size = Pt(9)
                    run.bold = True

        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = row.get('number_1', "")
            para1 = row_cells[1].paragraphs[0]
            para2 = row_cells[2].paragraphs[0]
            TitleComparison._highlight_differences(para1, para2, row['title_excel1'], row['title_excel2'])

        doc.save(output_path)
        print("Report saved at:", output_path)

    @staticmethod
    def _add_summary(doc, total_titles, mismatched_count):
        doc.add_heading('Summary of Title Comparison', level=2)
        total_para = doc.add_paragraph(f"Total Titles Compared: {total_titles}")
        total_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        total_run = total_para.runs[0]
        total_run.font.name = 'Arial'
        total_run.font.size = Pt(10)

        mismatch_para = doc.add_paragraph(f"Titles with Differences: {mismatched_count}")
        mismatch_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        mismatch_run = mismatch_para.runs[0]
        mismatch_run.font.name = 'Arial'
        mismatch_run.font.size = Pt(10)

    @staticmethod
    def tokenize_with_indices(text):
        tokens = []
        for match in re.finditer(r'[^\s]+', text):
            tokens.append((match.group(), match.start()))
        print("DEBUG (Report): Tokenized Text with Indices:", tokens)
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
            if backtrack[i][j] == 'DIAG':
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
            elif backtrack[i][j] == 'UP':
                aligned_tokens1.append(tokens1[i - 1])
                aligned_tokens2.append((None, None))
                flags.append("MISSING_2")
                i -= 1
            elif backtrack[i][j] == 'LEFT':
                aligned_tokens1.append((None, None))
                aligned_tokens2.append(tokens2[j - 1])
                flags.append("MISSING_1")
                j -= 1
        return aligned_tokens1[::-1], aligned_tokens2[::-1], flags[::-1]

    @staticmethod
    def _highlight_differences(paragraph1, paragraph2, text1, text2):
        tokens1 = TitleComparison.tokenize_with_indices(text1)
        tokens2 = TitleComparison.tokenize_with_indices(text2)
        aligned_tokens1, aligned_tokens2, flags = TitleComparison.dp_align_tokens(tokens1, tokens2)
        TitleComparison.reconstruct_text_with_flags(paragraph1, aligned_tokens1, aligned_tokens2, flags)
        TitleComparison.reconstruct_text_with_flags(paragraph2, aligned_tokens2, aligned_tokens1, flags)

    @staticmethod
    def reconstruct_text_with_flags(paragraph, tokens1, tokens2, flags):
        for (token1, index1), (token2, index2), flag in zip(tokens1, tokens2, flags):
            if token1 is None:
                continue
            if index1 > 0:
                paragraph.add_run(" ")
            if flag == "EXACT":
                paragraph.add_run(token1)
            elif flag == "CASE_ONLY":
                run = paragraph.add_run(token1)
                run.font.color.rgb = RGBColor(128, 128, 128)
            elif flag == "CHAR_LEVEL":
                matcher = SequenceMatcher(None, token1, token2)
                for op, i1, i2, j1, j2 in matcher.get_opcodes():
                    text_part = token1[i1:i2]
                    if op == "equal":
                        paragraph.add_run(text_part)
                    else:
                        diff_run = paragraph.add_run(text_part)
                        diff_run.font.color.rgb = RGBColor(255, 165, 0)
            elif flag in ["MISSING_1", "MISSING_2"]:
                run = paragraph.add_run(token1)
                run.font.color.rgb = RGBColor(255, 0, 0)
        return paragraph


class MergerGUI:
    """Handles the GUI for the Excel merger."""

    def __init__(self, master=None):
        self.mergerApp = ctk.CTkToplevel(master) if master else ctk.CTk()
        self.mergerApp.title("Merger Tool")

        self.excel1_path = tk.StringVar()
        self.excel2_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.ref_column1 = tk.StringVar()
        self.title_column1 = tk.StringVar()
        self.ref_column2 = tk.StringVar()
        self.title_column2 = tk.StringVar()

        self.generate_report = tk.BooleanVar(value=False)

        self.excel1_headers = []
        self.excel2_headers = []

        self._build_gui()

    def _build_gui(self):
        self.mergerApp.grid_rowconfigure(0, weight=1)
        self.mergerApp.grid_columnconfigure(0, weight=1)
        self.mergerApp.grid_columnconfigure(1, weight=1)

        self.excel1_frame = ctk.CTkFrame(self.mergerApp)
        self.excel1_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.excel2_frame = ctk.CTkFrame(self.mergerApp)
        self.excel2_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        for frame in (self.excel1_frame, self.excel2_frame):
            frame.grid_rowconfigure(0, weight=0)
            frame.grid_columnconfigure(0, weight=0)
            frame.grid_columnconfigure(1, weight=1)

        self._build_excel1_section(self.excel1_frame)
        self._build_excel2_section(self.excel2_frame)

        self.controls_frame = ctk.CTkFrame(self.mergerApp)
        self.controls_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        self._build_controls(self.controls_frame)

    def _build_excel1_section(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12

        self.excel1_button = ctk.CTkButton(parent_frame, text="Select Excel File 1",
                                           command=self._browse_excel1, width=200, height=40)
        self.excel1_button.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(parent_frame, text="Reference Column:", font=(font_name, font_size)).grid(
            row=1, column=0, padx=5, pady=2, sticky="e")
        self.ref_option_menu1 = ctk.CTkOptionMenu(parent_frame, variable=self.ref_column1, values=[])
        self.ref_option_menu1.grid(row=1, column=1, padx=5, pady=2, sticky="ew")

        ctk.CTkCheckBox(parent_frame, text="Generate Title Comparison Report",
                        variable=self.generate_report, command=self._toggle_title_entries).grid(
            row=2, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        ctk.CTkLabel(parent_frame, text="Drawing Title:", font=(font_name, font_size)).grid(
            row=3, column=0, padx=5, pady=2, sticky="e")
        self.title_option_menu1 = ctk.CTkOptionMenu(parent_frame, variable=self.title_column1, values=[], state="disabled")
        self.title_option_menu1.grid(row=3, column=1, padx=5, pady=2, sticky="ew")

    def _build_excel2_section(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12

        self.excel2_button = ctk.CTkButton(parent_frame, text="Select Excel File 2",
                                           command=self._browse_excel2, width=200, height=40)
        self.excel2_button.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(parent_frame, text="Reference Column:", font=(font_name, font_size)).grid(
            row=1, column=0, padx=5, pady=2, sticky="e")
        self.ref_option_menu2 = ctk.CTkOptionMenu(parent_frame, variable=self.ref_column2, values=[])
        self.ref_option_menu2.grid(row=1, column=1, padx=5, pady=2, sticky="ew")

        ctk.CTkLabel(parent_frame, text="", font=(font_name, font_size)).grid(
            row=2, column=0, padx=5, pady=2, sticky="e")

        ctk.CTkLabel(parent_frame, text="Drawing Title:", font=(font_name, font_size)).grid(
            row=3, column=0, padx=5, pady=2, sticky="e")
        self.title_option_menu2 = ctk.CTkOptionMenu(parent_frame, variable=self.title_column2, values=[], state="disabled")
        self.title_option_menu2.grid(row=3, column=1, padx=5, pady=2, sticky="ew")

    def _build_controls(self, parent_frame):
        font_name = "Helvetica"
        font_size = 12

        ctk.CTkLabel(parent_frame, text="Output Path:", font=(font_name, font_size)).grid(
            row=0, column=0, padx=5, pady=2, sticky="e")
        ctk.CTkEntry(parent_frame, textvariable=self.output_path, width=300).grid(
            row=0, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkButton(parent_frame, text="Use Excel 1 Path", command=self._use_excel1_path).grid(
            row=0, column=2, padx=5, pady=2)

        ctk.CTkButton(parent_frame, text="Start Merge", command=self._start_merge).grid(
            row=1, column=0, columnspan=3, pady=10)
        parent_frame.grid_columnconfigure(1, weight=1)

    def _browse_excel1(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")],
                                               title="Select Excel File 1")
        if file_path:
            self.excel1_path.set(file_path)
            self._load_excel1_headers(file_path)
            self.excel1_button.configure(fg_color="#217346")  # Excel green

    def _browse_excel2(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")],
                                               title="Select Excel File 2")
        if file_path:
            self.excel2_path.set(file_path)
            self._load_excel2_headers(file_path)
            self.excel2_button.configure(fg_color="#217346")  # Excel green

    def _load_excel1_headers(self, file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
            headers = list(df.columns)
            self.excel1_headers = headers
            self.ref_option_menu1.configure(values=headers)
            self.title_option_menu1.configure(values=headers)
            if headers:
                self.ref_column1.set(headers[0])
                self.title_column1.set(headers[0])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load headers from Excel File 1: {e}")

    def _load_excel2_headers(self, file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
            headers = list(df.columns)
            self.excel2_headers = headers
            self.ref_option_menu2.configure(values=headers)
            self.title_option_menu2.configure(values=headers)
            if headers:
                self.ref_column2.set(headers[0])
                self.title_column2.set(headers[0])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load headers from Excel File 2: {e}")

    def _use_excel1_path(self):
        excel1 = self.excel1_path.get()
        if excel1:
            directory, file_name = os.path.split(excel1)
            name, ext = os.path.splitext(file_name)
            self.output_path.set(os.path.join(directory, f"{name}_merged{ext}"))

    def _toggle_title_entries(self):
        state = "normal" if self.generate_report.get() else "disabled"
        self.title_option_menu1.configure(state=state)
        self.title_option_menu2.configure(state=state)

    def _start_merge(self):
        excel1_path = self.excel1_path.get()
        excel2_path = self.excel2_path.get()
        ref_column1 = self.ref_column1.get()
        ref_column2 = self.ref_column2.get()
        output_path = self.output_path.get() or excel1_path

        try:
            if self.generate_report.get():
                title_col1 = self.title_column1.get()
                title_col2 = self.title_column2.get()
            else:
                title_col1 = title_col2 = None

            merged_file_path, merged_df = ExcelMerger.merge_excels(
                excel1_path, excel2_path, ref_column1, ref_column2, output_path,
                title_column1=title_col1, title_column2=title_col2
            )
            messagebox.showinfo("Success", f"Merged file saved at {merged_file_path}")

            if self.generate_report.get():
                report_path = os.path.splitext(merged_file_path)[0] + "-TitleComparison.docx"
                TitleComparison.create_report(merged_df, title_col1, title_col2, report_path)
                messagebox.showinfo("Success", f"Title comparison report saved at {report_path}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run(self):
        if isinstance(self.mergerApp, ctk.CTk):
            self.mergerApp.mainloop()


if __name__ == "__main__":
    app = MergerGUI()
    app.run()
