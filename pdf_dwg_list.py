import os
import pandas as pd
from tkinter import Tk, filedialog
from openpyxl import load_workbook


def choose_directory(title="Select Directory"):
    root = Tk()
    root.withdraw()  # Hide the main window
    selected_path = filedialog.askdirectory(title=title)
    root.destroy()  # Close the main window
    return selected_path


def choose_file_save_location(title="Save Excel File As"):
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title=title
    )
    root.destroy()  # Close the main window
    return file_path

def list_files(start_path):
    file_dict = {'PDF': {}, 'DWG': {}}
    for root, dirs, files in os.walk(start_path):
        for file in files:
            file_path = os.path.join(root, file)
            directory = os.path.relpath(root, start_path)
            file_name, file_extension = os.path.splitext(file)
            file_type = 'PDF' if file_extension.lower() == '.pdf' else 'DWG' if file_extension.lower() == '.dwg' else 'Other'

            # Exclude 'Other' file types
            if file_type in ['PDF', 'DWG']:
                # Use the full path as the key in the dictionary
                file_dict[file_type][file_path] = file_name

    return file_dict

def export_to_excel(file_dict, selected_folder, excel_file):
    pdf_dict = file_dict['PDF']
    dwg_dict = file_dict['DWG']

    # Combine PDF and DWG dictionaries
    combined_dict = {}

    # Helper function to add or update entries in the combined_dict
    def update_combined_dict(file_name, pdf_path, dwg_path):
        if file_name not in combined_dict:
            combined_dict[file_name] = {'File': file_name, 'PDF': None, 'DWG': None, 'FolderPDF': [], 'FolderDWG': [],
                                        'PDFHasDuplicate': None, 'DWGHasDuplicate': None}

        # Add PDF information
        if pdf_path:
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            combined_dict[file_name]['PDF'] = pdf_name

            # Check for duplicate entries and add square brackets to FolderPDF
            pdf_folder = os.path.relpath(os.path.dirname(pdf_path), selected_folder) if os.path.dirname(
                pdf_path) and selected_folder else ''
            combined_dict[file_name]['FolderPDF'].append(pdf_folder)
            combined_dict[file_name]['PDFHasDuplicate'] = len(combined_dict[file_name]['FolderPDF']) if len(
                combined_dict[file_name]['FolderPDF']) > 1 else None

        # Add DWG information
        if dwg_path:
            dwg_name = os.path.splitext(os.path.basename(dwg_path))[0]
            combined_dict[file_name]['DWG'] = dwg_name

            # Check for duplicate entries and add square brackets to FolderDWG
            dwg_folder = os.path.relpath(os.path.dirname(dwg_path), selected_folder) if os.path.dirname(
                dwg_path) and selected_folder else ''
            combined_dict[file_name]['FolderDWG'].append(dwg_folder)
            combined_dict[file_name]['DWGHasDuplicate'] = len(combined_dict[file_name]['FolderDWG']) if len(
                combined_dict[file_name]['FolderDWG']) > 1 else None

    # Process PDF files
    for pdf_path, pdf_name in pdf_dict.items():
        dwg_path = dwg_dict.get(pdf_path, '')
        file_name = os.path.splitext(os.path.basename(pdf_path))[0]
        update_combined_dict(file_name, pdf_path, dwg_path)

    # Process DWG files
    for dwg_path, dwg_name in dwg_dict.items():
        pdf_path = pdf_dict.get(dwg_path, '')
        file_name = os.path.splitext(os.path.basename(dwg_path))[0]
        update_combined_dict(file_name, pdf_path, dwg_path)

    # Check for duplicate entries and update 'HasDuplicate' column
    for file_name, data in combined_dict.items():
        pdf_duplicate = any(entry['PDF'] == data['PDF'] for entry in combined_dict.values() if entry != data)
        dwg_duplicate = any(entry['DWG'] == data['DWG'] for entry in combined_dict.values() if entry != data)
        combined_dict[file_name]['HasDuplicate'] = not (pdf_duplicate or dwg_duplicate)

    # Create a DataFrame from the combined dictionary
    df = pd.DataFrame.from_dict(combined_dict, orient='index').reset_index()

    # Sort the DataFrame alphabetically by File name
    df = df.sort_values(by='File')

    # Drop the first column created by reset_index()
    df = df.drop(columns=[df.columns[0], 'File', 'HasDuplicate'])  # Drop the first column and 'HasDuplicate'

    # Save the final DataFrame to Excel
    df.to_excel(excel_file, index=False)

    # Apply conditional formatting for duplicates
    wb = load_workbook(excel_file)
    ws = wb.active

    # Save the workbook
    wb.save(excel_file)

    print(f"Directory listing exported to {excel_file}")


def pdf_dwg_counter():
    # Choose the directory using a dialog box
    directory_path = choose_directory()

    if directory_path:
        # Get the dictionary of files in the directory
        file_dict = list_files(directory_path)

        # Choose where to save the Excel file
        excel_file_path = choose_file_save_location()

        if excel_file_path:
            # Export the dictionary to the specified Excel file
            export_to_excel(file_dict, directory_path, excel_file_path)
        else:
            print("No file location selected.")
    else:
        print("No directory selected.")

'''
End of Counter functions
'''
