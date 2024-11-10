import os
import pandas as pd
import tkinter as tk
import customtkinter as ctk
from openpyxl import load_workbook
from tkinter import filedialog, messagebox


def merge_excels(excel1_path, excel2_path, ref_column1, ref_column2, output_path):
    """Merge two Excel files based on specified reference columns."""
    # Load Excel files
    excel1 = pd.read_excel(excel1_path, engine='openpyxl')
    excel2 = pd.read_excel(excel2_path, engine='openpyxl')

    # Keep track of hyperlinks in Excel1
    hyperlinks_excel1 = {}
    wb1 = load_workbook(excel1_path)
    ws1 = wb1.active

    # Loop through cells to store hyperlinks
    for row in ws1.iter_rows(min_row=2):
        for cell in row:
            if cell.hyperlink:
                hyperlinks_excel1[cell.row] = hyperlinks_excel1.get(cell.row, {})
                hyperlinks_excel1[cell.row][cell.column] = cell.hyperlink.target

    # Fill NaN and add cumulative count for alignment
    excel1[ref_column1] = excel1[ref_column1].fillna('BLANK_EXCEL1')
    excel2[ref_column2] = excel2[ref_column2].fillna('BLANK_EXCEL2')
    excel1['refno_count'] = excel1.groupby(ref_column1).cumcount()
    excel2['refno_count'] = excel2.groupby(ref_column2).cumcount()

    # Rename columns and merge
    excel1 = excel1.rename(columns={ref_column1: 'refno1'})
    excel2 = excel2.rename(columns={ref_column2: 'refno2'})
    merged = pd.merge(excel1, excel2, left_on=['refno1', 'refno_count'], right_on=['refno2', 'refno_count'],
                      how='outer').drop(columns=['refno_count'])

    # Replace placeholders back with NaN
    merged['refno1'] = merged['refno1'].replace('BLANK_EXCEL1', pd.NA)
    merged['refno2'] = merged['refno2'].replace('BLANK_EXCEL2', pd.NA)

    # Determine the file path based on output_path
    if os.path.isdir(output_path):
        temp_file_path = os.path.join(output_path, 'merged_result_temp.xlsx')
    else:
        temp_file_path = output_path

    # Save to Excel without hyperlinks
    merged.to_excel(temp_file_path, index=False)

    # Reopen merged file to re-add hyperlinks
    wb_merged = load_workbook(temp_file_path)
    ws_merged = wb_merged.active

    # Restore hyperlinks and apply hyperlink style
    for row_idx, link_dict in hyperlinks_excel1.items():
        for col_idx, link in link_dict.items():
            cell = ws_merged.cell(row=row_idx, column=col_idx)
            cell.hyperlink = link
            cell.style = "Hyperlink"  # Apply the built-in Hyperlink style

    # Save final file with hyperlinks restored
    wb_merged.save(temp_file_path)

    return temp_file_path


def open_file_dialog(title):
    """Open a file dialog to select an Excel file."""
    return filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")], title=title)


def start_merge(excel1_entry, excel2_entry, ref1_entry, ref2_entry, output_entry):
    """Start the merge process and inform the user of completion."""
    excel1_path = excel1_entry.get()
    excel2_path = excel2_entry.get()
    ref_column1 = ref1_entry.get()
    ref_column2 = ref2_entry.get()
    output_path = output_entry.get() or excel1_path  # Default to excel1_path if output_entry is empty

    merged_file_path = merge_excels(excel1_path, excel2_path, ref_column1, ref_column2, output_path)
    messagebox.showinfo("Success", f"Merged file saved as {merged_file_path}")



def browse_excel1(excel1_entry):
    """Browse and set the path for Excel 1."""
    file_path = open_file_dialog("Select Excel 1")
    excel1_entry.delete(0, tk.END)
    excel1_entry.insert(0, file_path)


def browse_excel2(excel2_entry):
    """Browse and set the path for Excel 2."""
    file_path = open_file_dialog("Select Excel 2")
    excel2_entry.delete(0, tk.END)
    excel2_entry.insert(0, file_path)


"""ENTRY POINT"""
def create_merger_gui():
    """Create the GUI for the Excel merger."""
    # Create the merger window with a unique name
    merger_window = ctk.CTk()
    merger_window.title("Merger Tool")

    font_name = "Helvetica"
    font_size = 12
    button_font = "Helvetica"

    # Create two frames for layout
    left_frame = ctk.CTkFrame(merger_window)
    left_frame.grid(row=0, column=0, padx=10, pady=10)

    right_frame = ctk.CTkFrame(merger_window)
    right_frame.grid(row=0, column=1, padx=10, pady=10)

    # Left Frame: Excel File 1 and 2 Entries & Output Path
    # Excel File 1 Entry
    excel1_browse_button = ctk.CTkButton(left_frame, text="Browse", command=lambda: browse_excel1(excel1_entry),
                                         font=(font_name, font_size))
    excel1_browse_button.grid(row=0, column=0, padx=10, pady=5)

    excel1_entry = ctk.CTkEntry(left_frame, width=200, font=(font_name, 9), placeholder_text="..Select Excel File 1")
    excel1_entry.grid(row=0, column=1, padx=10, pady=5)



    # Excel File 2 Entry
    excel2_browse_button = ctk.CTkButton(left_frame, text="Browse", command=lambda: browse_excel2(excel2_entry),
                                         font=(font_name, font_size))
    excel2_browse_button.grid(row=1, column=0, padx=10, pady=5)

    excel2_entry = ctk.CTkEntry(left_frame, width=200, font=(font_name, 9), placeholder_text="..Select Excel File 2")
    excel2_entry.grid(row=1, column=1, padx=10, pady=5)



    # Output Path Entry

    output_entry = ctk.CTkEntry(left_frame, width=200, font=(font_name, 9), placeholder_text="Output file path")
    output_entry.grid(row=2, column=1, padx=10, pady=5)

    # Button to set output path to Excel 1
    use_excel1_button = ctk.CTkButton(left_frame, text="Use Excel 1 Path",
                                      command=lambda: output_entry.insert(0, excel1_entry.get()),
                                      font=(font_name, font_size))
    use_excel1_button.grid(row=3, column=0, padx=5, pady=5)

    # Button to set output path to Excel 2
    use_excel2_button = ctk.CTkButton(left_frame, text="Use Excel 2 Path",
                                      command=lambda: output_entry.insert(0, excel2_entry.get()),
                                      font=(font_name, font_size))
    use_excel2_button.grid(row=3, column=1, padx=5, pady=5)

    # Right Frame: Reference Columns
    # Reference Column Entry
    ref1_label = ctk.CTkLabel(right_frame, text="Reference Column (Excel 1):", font=(font_name, font_size))
    ref1_label.grid(row=0, column=0, padx=10, pady=5, sticky="e")

    ref1_entry = ctk.CTkEntry(right_frame, width=200, font=(font_name, 9), placeholder_text="Column Name of Reference")
    ref1_entry.grid(row=0, column=1, padx=10, pady=5)
    ref1_entry.insert(0,"Area 1")

    ref2_label = ctk.CTkLabel(right_frame, text="Reference Column (Excel 2):", font=(font_name, font_size))
    ref2_label.grid(row=1, column=0, padx=10, pady=5, sticky="e")

    ref2_entry = ctk.CTkEntry(right_frame, width=200, font=(font_name, 9), placeholder_text="Column Name of Reference")
    ref2_entry.grid(row=1, column=1, padx=10, pady=5)
    ref2_entry.insert(0,"SHEET NO")

    # Case sensitivity label
    case_sensitive_label = ctk.CTkLabel(right_frame, text="Note: Reference columns are CASE SENSITIVE.",
                                        fg_color="transparent", text_color="gray59",
                                        padx=0, pady=0, anchor="nw", font=(button_font, 9.5))
    case_sensitive_label.grid(row=2, column=0, columnspan=2, padx=10, pady=5)

    # Start Merge Button
    start_button = ctk.CTkButton(merger_window, text="Start Merge",
                                 command=lambda: start_merge(excel1_entry, excel2_entry, ref1_entry, ref2_entry,
                                                             output_entry),
                                 font=(font_name, font_size))
    start_button.grid(row=1, column=0, columnspan=2, pady=10)

    # Run the merger GUI
    merger_window.mainloop()


if __name__ == "__main__":
    create_merger_gui()