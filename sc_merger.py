import os
import tkinter as tk

from tkinter import filedialog, messagebox
from difflib import SequenceMatcher

import customtkinter as ctk
import pandas as pd
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from docx.shared import RGBColor
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

import re
import unicodedata


class ExcelMerger:
    """Handles Excel merging logic, including conditional formatting and hyperlink handling."""

    @staticmethod
    def merge_excels(excel1_path, excel2_path, ref_column1, ref_column2, output_path):
        """
        Merge two Excel files while retaining hyperlinks and applying formatting.
        """
        # Read Excel files
        excel1 = pd.read_excel(excel1_path, engine='openpyxl', dtype=str).fillna("")
        excel2 = pd.read_excel(excel2_path, engine='openpyxl', dtype=str).fillna("")

        # Ensure the reference columns exist
        if ref_column1 not in excel1.columns or ref_column2 not in excel2.columns:
            raise KeyError("Reference columns not found in one or both Excel files.")

        # Add `original_row_index` to track rows using the new method
        excel1 = ExcelMerger.add_original_row_index_to_dataframe(excel1, excel1_path)

        # Extract hyperlinks from the original file
        hyperlinks = ExcelMerger._extract_hyperlinks(excel1_path)

        # Prepare data for merging
        excel1['refno_count'] = excel1.groupby(ref_column1).cumcount()
        excel2['refno_count'] = excel2.groupby(ref_column2).cumcount()
        excel1 = excel1.rename(columns={ref_column1: 'refno1'})
        excel2 = excel2.rename(columns={ref_column2: 'refno2'})

        # Merge data and handle missing `original_row_index`
        merged_df = pd.merge(
            excel1, excel2,
            left_on=['refno1', 'refno_count'],
            right_on=['refno2', 'refno_count'],
            how='outer'
        ).drop(columns=['refno_count']).fillna("")
        merged_df['original_row_index'] = pd.to_numeric(merged_df['original_row_index'], errors='coerce').fillna(
            0).astype(int)

        # Save the merged file without formatting
        temp_file_path = ExcelMerger._save_merged_to_excel(merged_df, output_path)

        # Apply formatting and hyperlinks
        ExcelMerger._apply_formatting_and_hyperlinks(temp_file_path, hyperlinks, merged_df)

        return temp_file_path, merged_df

    @staticmethod
    def _extract_hyperlinks(file_path):
        """
        Extract hyperlinks from the original Excel file and map them using row numbers.
        """
        print("Extracting hyperlinks from:", file_path)
        hyperlinks = {}
        wb = load_workbook(file_path, data_only=False)
        ws = wb.active

        # Loop through all rows and columns to find hyperlinks
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):  # Start at row 2 to skip headers
            row_number = row[0].row  # Get the row number directly from the cell
            hyperlinks[row_number] = {}
            for cell in row:
                if cell.hyperlink:
                    col_idx = cell.column  # Numeric column index
                    print(f"Hyperlink found at row {row_number}, col {col_idx}: {cell.hyperlink.target}")
                    hyperlinks[row_number][col_idx] = cell.hyperlink.target

        print("Extracted Hyperlinks:", hyperlinks)
        return hyperlinks

    @staticmethod
    def add_original_row_index_to_dataframe(df, file_path):
        """
        Add `original_row_index` to the DataFrame by matching DataFrame rows
        with non-empty rows in the Excel file.
        """
        from openpyxl import load_workbook

        print(f"Adding original row indices from file: {file_path}")

        # Load the workbook and active worksheet
        wb = load_workbook(file_path, data_only=False)
        ws = wb.active

        # Prepare a list to store original row indices
        original_row_indices = []

        # Iterate through rows in the worksheet, skipping the header row
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):  # Row 2 onwards
            # Construct a tuple of cell values for the row
            row_values = [cell.value for cell in row]
            # Check if the row is non-empty (contains at least one non-blank cell)
            if any(row_values):
                original_row_indices.append(row[0].row)  # Store the actual Excel row index

        # Match the extracted indices with the DataFrame rows
        if len(original_row_indices) != len(df):
            raise ValueError(
                f"Mismatch between extracted row indices ({len(original_row_indices)}) "
                f"and DataFrame rows ({len(df)}). Ensure no extra blank rows in Excel."
            )

        # Assign the original row indices to the DataFrame
        df['original_row_index'] = original_row_indices

        print(f"Assigned original row indices: {original_row_indices}")
        return df

    @staticmethod
    def _save_merged_to_excel(df, output_path):
        """
        Save the merged DataFrame to an Excel file.
        """
        temp_file_path = output_path if not os.path.isdir(output_path) else os.path.join(output_path, 'merged_result_temp.xlsx')
        df.to_excel(temp_file_path, index=False, header=True)
        return temp_file_path

    @staticmethod
    def _apply_formatting_and_hyperlinks(file_path, hyperlinks, merged_df):
        """
        Apply conditional formatting for mismatches and duplicates,
        and reapply hyperlinks using the `hyperlinks` dictionary.
        """
        wb = load_workbook(file_path)
        ws = wb.active
        print("Applying formatting and hyperlinks to:", file_path)

        # Define styles for formatting
        fill_missing_refno1 = PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid")
        fill_missing_refno2 = PatternFill(start_color="FFCC66", end_color="FFCC66", fill_type="solid")
        duplicate_font = Font(bold=True, color="FF3300")

        # Get column indices for refno1 and refno2
        refno1_col_idx = merged_df.columns.get_loc('refno1') + 1
        refno2_col_idx = merged_df.columns.get_loc('refno2') + 1

        # Identify duplicates in refno1 and refno2
        refno1_duplicates = merged_df['refno1'][merged_df['refno1'].duplicated(keep=False)].tolist()
        refno2_duplicates = merged_df['refno2'][merged_df['refno2'].duplicated(keep=False)].tolist()

        # Apply mismatch and duplicate formatting
        for row_idx in range(2, len(merged_df) + 2):  # Start at row 2 for Excel
            refno1_value = ws.cell(row=row_idx, column=refno1_col_idx).value
            refno2_value = ws.cell(row=row_idx, column=refno2_col_idx).value

            # Highlight missing refno1 or refno2
            if refno1_value and not refno2_value:
                ws.cell(row=row_idx, column=refno1_col_idx).fill = fill_missing_refno1
            if refno2_value and not refno1_value:
                ws.cell(row=row_idx, column=refno2_col_idx).fill = fill_missing_refno2

            # Apply duplicate formatting
            if refno1_value in refno1_duplicates:
                ws.cell(row=row_idx, column=refno1_col_idx).font = duplicate_font
            if refno2_value in refno2_duplicates:
                ws.cell(row=row_idx, column=refno2_col_idx).font = duplicate_font

        # Apply hyperlinks
        for original_row_index, columns in hyperlinks.items():
            # Find the new row in the merged DataFrame
            new_row = merged_df[merged_df['original_row_index'] == original_row_index]

            # If no matching row is found, skip
            if new_row.empty:
                print(f"No matching row for original_row_index: {original_row_index}")
                continue

            # Get the new Excel row index (1-based)
            new_row_idx = new_row.index[0] + 2  # DataFrame index is zero-based; Excel rows start at 2

            # Apply hyperlinks to the recorded column positions
            for col_idx, hyperlink in columns.items():
                try:
                    ws.cell(row=new_row_idx, column=col_idx).hyperlink = hyperlink
                    ws.cell(row=new_row_idx, column=col_idx).style = "Hyperlink"
                    print(f"Applied hyperlink at new_row_idx: {new_row_idx}, col_idx: {col_idx}, link: {hyperlink}")
                except Exception as e:
                    print(f"Error applying hyperlink for row {new_row_idx}, column {col_idx}: {e}")

        # Save the workbook with updated formatting and hyperlinks
        wb.save(file_path)


class TitleComparison:
    """Handles title comparison logic and generates a Word report."""

    @staticmethod
    def create_report(df, title_column1, title_column2, output_path):
        """
        Create a Word document highlighting differences between two title columns.
        Include all rows, even those with no differences.
        """
        doc = Document()
        doc.add_heading('Title Differences Report', level=1)

        # Add summary
        TitleComparison._add_summary(doc, len(df), len(df[df[title_column1] != df[title_column2]]))

        # Create a table
        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = "Drawing Number"
        header_cells[1].text = title_column1
        header_cells[2].text = title_column2

        for cell in header_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Verdana'
                    run.font.size = Pt(10)
                    run.bold = True

        # Populate rows
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = row.get('refno1', "")

            # Get paragraphs for both title columns
            para1 = row_cells[1].paragraphs[0]
            para2 = row_cells[2].paragraphs[0]

            # Highlight differences
            TitleComparison._highlight_differences(para1, para2, row[title_column1], row[title_column2])

        # Save the document
        doc.save(output_path)

    @staticmethod
    def _add_summary(doc, total_titles, mismatched_count):
        """
        Add a summary of the title comparison to the Word document with styled text.
        """
        doc.add_heading('Summary of Title Comparison', level=2)

        # Style for total titles paragraph
        total_para = doc.add_paragraph(f"Total Titles Compared: {total_titles}")
        total_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        total_run = total_para.runs[0]
        total_run.font.name = 'Arial'
        total_run.font.size = Pt(10)

        # Style for mismatched titles paragraph
        mismatch_para = doc.add_paragraph(f"Titles with Differences: {mismatched_count}")
        mismatch_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        mismatch_run = mismatch_para.runs[0]
        mismatch_run.font.name = 'Arial'
        mismatch_run.font.size = Pt(10)

    @staticmethod
    def _highlight_differences(paragraph1, paragraph2, text1, text2):
        """
        Highlight differences between two pieces of text, preserving spaces for rendering.
        """
        tokens1 = TitleComparison.tokenize_with_indices(text1)
        tokens2 = TitleComparison.tokenize_with_indices(text2)

        # Perform alignment
        aligned_tokens1, aligned_tokens2, flags = TitleComparison.dp_align_tokens(tokens1, tokens2)

        # Reconstruct the text in Word
        TitleComparison.reconstruct_text_with_flags(paragraph1, aligned_tokens1, flags)
        TitleComparison.reconstruct_text_with_flags(paragraph2, aligned_tokens2, flags)

    @staticmethod
    def tokenize_with_indices(text):
        """
        Tokenize text into words and punctuation, excluding spaces, and return indices.
        """
        tokens = []
        for match in re.finditer(r'[^\s]+', text):
            tokens.append((match.group(), match.start()))
        print("DEBUG: Tokenized Text with Indices:", tokens)
        return tokens

    @staticmethod
    def dp_align_tokens(tokens1, tokens2):
        """
        Align tokens from two lists using dynamic programming for optimal alignment.
        """
        # Extract token strings and their indices
        token_strs1 = [t[0] for t in tokens1]
        token_strs2 = [t[0] for t in tokens2]

        n, m = len(token_strs1), len(token_strs2)
        dp = [[0] * (m + 1) for _ in range(n + 1)]
        backtrack = [[None] * (m + 1) for _ in range(n + 1)]

        # Fill DP table
        for i in range(1, n + 1):
            dp[i][0] = i  # Cost of deletions
            backtrack[i][0] = 'UP'
        for j in range(1, m + 1):
            dp[0][j] = j  # Cost of insertions
            backtrack[0][j] = 'LEFT'

        for i in range(1, n + 1):
            for j in range(1, m + 1):
                similarity = SequenceMatcher(None, token_strs1[i - 1], token_strs2[j - 1]).ratio()

                # Calculate costs
                if similarity == 1.0:  # Exact match
                    replace_cost = dp[i - 1][j - 1]
                elif similarity >= 0.8:  # High similarity
                    replace_cost = dp[i - 1][j - 1] + (1 - similarity)
                elif similarity >= 0.4:  # Moderate similarity
                    replace_cost = dp[i - 1][j - 1] + 2
                else:  # Completely unrelated tokens
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

        # Traceback to align tokens
        aligned_tokens1, aligned_tokens2, flags = [], [], []
        i, j = n, m
        while i > 0 or j > 0:
            if backtrack[i][j] == 'DIAG':
                aligned_tokens1.append(tokens1[i - 1])
                aligned_tokens2.append(tokens2[j - 1])
                flags.append("EXACT" if token_strs1[i - 1] == token_strs2[j - 1] else "CHAR_LEVEL")
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
    def reconstruct_text_with_flags(paragraph, tokens, flags):
        """
        Reconstruct the text in the paragraph using original indices and apply flags for highlighting.
        """
        for (token, index), flag in zip(tokens, flags):
            if token is None:
                continue

            # Add a space before the token if required
            if index > 0:
                paragraph.add_run(" ")

            run = paragraph.add_run(token)

            if flag == "EXACT":
                continue  # No formatting needed
            elif flag == "CASE_ONLY":
                run.font.color.rgb = RGBColor(128, 128, 128)  # Gray
            elif flag == "CHAR_LEVEL":
                run.font.color.rgb = RGBColor(255, 165, 0)  # Orange
            elif flag == "MISSING_1" or flag == "MISSING_2":
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red


class MergerGUI:
    """Handles the GUI for the Excel merger."""

    def __init__(self, master=None):
        self.mergerApp = ctk.CTkToplevel(master) if master else ctk.CTk()
        self.mergerApp.title("Merger Tool")

        # Initialize tkinter variables
        self.excel1_path = tk.StringVar()
        self.excel2_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.ref_column1 = tk.StringVar(value="Area 1")
        self.ref_column2 = tk.StringVar(value="SHEET NO")
        self.title_column1 = tk.StringVar(value="Drawing Title")
        self.title_column2 = tk.StringVar(value="LOD Title")
        self.generate_report = tk.BooleanVar(value=False)

        self._build_gui()

    def _build_gui(self):
        """Build the GUI components."""
        # Frames for better layout management
        left_frame = ctk.CTkFrame(self.mergerApp)
        left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        right_frame = ctk.CTkFrame(self.mergerApp)
        right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        # Build components in respective frames
        self._build_file_selection(left_frame)
        self._build_reference_selection(right_frame)
        self._build_controls()

    def _build_file_selection(self, parent_frame):
        """Build file selection UI."""
        font_name = "Helvetica"
        font_size = 12

        # Excel 1
        ctk.CTkLabel(parent_frame, text="Excel File 1:", font=(font_name, font_size)).grid(row=0, column=0, padx=5, pady=5)
        ctk.CTkEntry(parent_frame, textvariable=self.excel1_path, width=200).grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkButton(parent_frame, text="Browse", command=self._browse_excel1).grid(row=0, column=2, padx=5, pady=5)

        # Excel 2
        ctk.CTkLabel(parent_frame, text="Excel File 2:", font=(font_name, font_size)).grid(row=1, column=0, padx=5, pady=5)
        ctk.CTkEntry(parent_frame, textvariable=self.excel2_path, width=200).grid(row=1, column=1, padx=5, pady=5)
        ctk.CTkButton(parent_frame, text="Browse", command=self._browse_excel2).grid(row=1, column=2, padx=5, pady=5)

        # Output Path
        ctk.CTkLabel(parent_frame, text="Output Path:", font=(font_name, font_size)).grid(row=2, column=0, padx=5, pady=5)
        ctk.CTkEntry(parent_frame, textvariable=self.output_path, width=200).grid(row=2, column=1, padx=5, pady=5)
        ctk.CTkButton(parent_frame, text="Use Excel 1 Path", command=self._use_excel1_path).grid(row=2, column=2, padx=5, pady=5)

    def _build_reference_selection(self, parent_frame):
        """Build reference column selection UI."""
        font_name = "Helvetica"
        font_size = 12

        # Reference Column 1
        ctk.CTkLabel(parent_frame, text="Reference Column (Excel 1):", font=(font_name, font_size)).grid(row=0,
                                                                                                         column=0,
                                                                                                         padx=5, pady=5)
        ctk.CTkEntry(parent_frame, textvariable=self.ref_column1, width=200).grid(row=0, column=1, padx=5, pady=5)

        # Reference Column 2
        ctk.CTkLabel(parent_frame, text="Reference Column (Excel 2):", font=(font_name, font_size)).grid(row=1,
                                                                                                         column=0,
                                                                                                         padx=5, pady=5)
        ctk.CTkEntry(parent_frame, textvariable=self.ref_column2, width=200).grid(row=1, column=1, padx=5, pady=5)

        # Generate Report Checkbox
        ctk.CTkCheckBox(parent_frame, text="Generate Title Comparison Report", variable=self.generate_report,
                        command=self._toggle_title_entries).grid(row=2, column=0, columnspan=2, padx=5, pady=5,
                                                                 sticky="w")

        # Title Columns (Initially disabled)
        self.title_entry1 = ctk.CTkEntry(parent_frame, textvariable=self.title_column1, width=200, state="disabled")
        self.title_entry1.grid(row=3, column=0, padx=5, pady=5)

        self.title_entry2 = ctk.CTkEntry(parent_frame, textvariable=self.title_column2, width=200, state="disabled")
        self.title_entry2.grid(row=3, column=1, padx=5, pady=5)

    def _build_controls(self):
        """Build the control buttons."""
        ctk.CTkButton(self.mergerApp, text="Start Merge", command=self._start_merge).grid(row=1, column=0, columnspan=2, pady=10)

    def _browse_excel1(self):
        """Browse for Excel File 1."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")], title="Select Excel File 1")
        if file_path:
            self.excel1_path.set(file_path)

    def _browse_excel2(self):
        """Browse for Excel File 2."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")], title="Select Excel File 2")
        if file_path:
            self.excel2_path.set(file_path)

    def _use_excel1_path(self):
        """Set output path to the same directory as Excel 1."""
        excel1 = self.excel1_path.get()
        if excel1:
            directory, file_name = os.path.split(excel1)
            name, ext = os.path.splitext(file_name)
            self.output_path.set(os.path.join(directory, f"{name}_merged{ext}"))

    def _toggle_title_entries(self):
        """Enable or disable title column entries based on the checkbox state."""
        state = "normal" if self.generate_report.get() else "disabled"
        # Directly configure the Entry widgets associated with title_column1 and title_column2
        self.title_entry1.configure(state=state)
        self.title_entry2.configure(state=state)

    def _start_merge(self):
        """Start the merge process."""
        excel1_path = self.excel1_path.get()
        excel2_path = self.excel2_path.get()
        ref_column1 = self.ref_column1.get()
        ref_column2 = self.ref_column2.get()
        output_path = self.output_path.get() or excel1_path  # Default to excel1_path if no output path is provided

        try:
            # Merge Excel files
            merged_file_path, merged_df = ExcelMerger.merge_excels(excel1_path, excel2_path, ref_column1, ref_column2, output_path)
            messagebox.showinfo("Success", f"Merged file saved at {merged_file_path}")

            # Generate title comparison report if required
            if self.generate_report.get():
                title_column1 = self.title_column1.get()
                title_column2 = self.title_column2.get()
                report_path = os.path.splitext(merged_file_path)[0] + "-TitleComparison.docx"
                TitleComparison.create_report(merged_df, title_column1, title_column2, report_path)
                messagebox.showinfo("Success", f"Title comparison report saved at {report_path}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run(self):
        """Run the GUI application."""
        if isinstance(self.mergerApp, ctk.CTk):  # Call mainloop only if it's the main window
            self.mergerApp.mainloop()


if __name__ == "__main__":
    app = MergerGUI()  # Standalone initialization
    app.run()