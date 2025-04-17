import os
import pandas as pd
from tkinter import Tk, filedialog, messagebox
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.styles import PatternFill

# List of extensions to check
EXTS = ["rvt", "ifc", "nwc", "dwfx", "nwd", "xml", "dwg"]

def choose_directory(title="Select Directory"):
    root = Tk()
    root.withdraw()
    selected_path = filedialog.askdirectory(title=title)
    root.destroy()
    return selected_path

def choose_file_save_location(title="Save Excel File As"):
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title=title
    )
    root.destroy()
    return file_path

def list_files(start_path):
    # Initialize a dict for each extension
    file_dict = {ext.upper(): {} for ext in EXTS}

    for root_dir, dirs, files in os.walk(start_path):
        for filename in files:
            name, ext = os.path.splitext(filename)
            ext = ext.lstrip('.').lower()
            if ext in EXTS:
                full_path = os.path.join(root_dir, filename)
                file_dict[ext.upper()][full_path] = {
                    'path': full_path,
                    'name': name,
                    'size': os.path.getsize(full_path),
                    'modified': datetime.fromtimestamp(os.path.getmtime(full_path))
                }
    return file_dict

def prepare_data_for_export(file_dict, selected_folder):
    # Gather all unique base filenames
    all_names = set()
    for ext_map in file_dict.values():
        for info in ext_map.values():
            all_names.add(info['name'])

    rows = []
    for name in sorted(all_names):
        row = {'FileName': name}
        for ext in EXTS:
            key = ext.upper()
            matches = [info for info in file_dict[key].values() if info['name'] == name]
            if matches:
                row[key] = "✓"
                row[f"{key}_DupCount"] = len(matches)
            else:
                row[key] = ""
                row[f"{key}_DupCount"] = 0
        rows.append(row)

    return pd.DataFrame(rows)

def save_to_excel(df, excel_file):
    # Write DataFrame to Excel
    df.to_excel(excel_file, index=False)
    wb = load_workbook(excel_file)
    ws = wb.active

    # Define fills
    present_fill = PatternFill(fill_type="solid", start_color="ABF5DF", end_color="ABF5DF")
    empty_fill   = PatternFill(fill_type="solid", start_color="000000", end_color="000000")
    dup_fill     = PatternFill(fill_type="solid", start_color="FF0000", end_color="FF0000")

    # Apply fills per extension column based on duplicate count
    for ext in EXTS:
        col_idx      = df.columns.get_loc(ext.upper()) + 1
        dup_col_idx  = df.columns.get_loc(f"{ext.upper()}_DupCount") + 1

        for row in range(2, ws.max_row + 1):
            dup_count = ws.cell(row=row, column=dup_col_idx).value or 0
            cell = ws.cell(row=row, column=col_idx)

            if dup_count > 1:
                cell.fill = dup_fill
            elif cell.value == "✓":
                cell.fill = present_fill
            else:
                cell.fill = empty_fill

    wb.save(excel_file)
    wb.close()
    print(f"Directory listing exported to {excel_file}")

def pdf_dwg_counter():
    directory_path = choose_directory()
    if not directory_path:
        print("No directory selected.")
        return

    file_dict = list_files(directory_path)
    excel_file_path = choose_file_save_location()
    if not excel_file_path:
        print("No file location selected.")
        return

    df = prepare_data_for_export(file_dict, directory_path)
    save_to_excel(df, excel_file_path)

    if messagebox.askyesno("Open Excel File", "Do you want to open the Excel file now?"):
        os.startfile(excel_file_path)

if __name__ == "__main__":
    pdf_dwg_counter()
