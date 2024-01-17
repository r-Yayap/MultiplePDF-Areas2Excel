import os
import pandas as pd
from tkinter import Tk, filedialog, Button, Label

def list_files_in_directory(selected_folder):
    file_list = []
    for root, dirs, files in os.walk(selected_folder):
        for file in files:
            full_path = os.path.join(root, file)
            relative_path = os.path.relpath(full_path, selected_folder)
            folder_name = os.path.dirname(relative_path)
            filename, file_format = os.path.splitext(file)
            file_list.append({'Folder': folder_name, 'Filename': filename, 'Format': file_format[1:]})

    return file_list

def create_excel_file(file_list, output_excel_path):
    df = pd.DataFrame(file_list)
    df.to_excel(output_excel_path, index=False)

def select_input_folder():
    root = Tk()
    root.withdraw()  # Hide the main window

    input_folder = filedialog.askdirectory(title="Select Input Folder")
    return input_folder

def select_output_folder():
    root = Tk()
    root.withdraw()  # Hide the main window

    output_folder = filedialog.askdirectory(title="Select Output Folder")
    return output_folder

def generate_file_list_and_excel():
    input_folder = select_input_folder()
    output_folder = select_output_folder()

    if input_folder and output_folder:  # Check if folders are selected
        files = list_files_in_directory(input_folder)

        # Prompt the user for the output file name
        output_file_name = filedialog.asksaveasfilename(
            title="Save As",
            initialdir=output_folder,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            defaultextension=".xlsx"
        )

        if output_file_name:
            # Ensure the file has the correct extension
            if not output_file_name.lower().endswith('.xlsx'):
                output_file_name += '.xlsx'

            output_excel_path = os.path.join(output_folder, output_file_name)
            create_excel_file(files, output_excel_path)
            print(f"Excel file created at {output_excel_path}")