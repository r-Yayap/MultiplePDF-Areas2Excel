import os
import pandas as pd
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog
from tkinter import ttk

class EditableTreeview(ttk.Treeview):
    def __init__(self, *args, **kwargs):
        ttk.Treeview.__init__(self, *args, **kwargs)
        self._entry = None
        self._col = None

        # Bind right-click to show context menu
        self.bind("<Button-3>", self.show_context_menu)
        

        # Create context menu
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Remove Row", command=self.remove_row)

    def on_double_click(self, event):
        item = self.focus()
        col = self.identify_column(event.x)
        if item and col and col != "#0":
            self._col = col
            cell_values = self.item(item, "values")
            if cell_values:
                col_index = int(col.split("#")[-1]) - 1
                cell_value = cell_values[col_index]
                self.edit_cell(item, col, cell_value)
                
    def show_context_menu(self, event):
        item = self.identify_row(event.y)
        if item:
            self.context_menu.post(event.x_brnroot, event.y_brnroot)
            
    def remove_row(self):
        item = self.focus()
        if item:
            self.delete(item)
            
    def on_focus_out(self, event):
        if self._entry is not None:
            self.stop_editing()

    def edit_cell(self, item, col, _):
        def on_ok():
            new_value = entry_var.get()
            if new_value:
                current_values = list(self.item(item, "values"))
                current_values[col_index] = new_value
                self.item(item, values=tuple(current_values))
            top.destroy()

        bbox = self.bbox(item, col)
        x, y, _, _ = bbox
        col_index = int(col.replace("#", "")) - 1  # Subtract 1 for 0-based indexing

        top = ctk.CTkToplevel(self)
        top.title("Edit Cell")

        entry_var = ctk.StringVar()
        entry_var.set(self.item(item, "values")[col_index])

        entry = ctk.CTkEntry(top, justify="center", textvariable=entry_var,
                             width=100, height=20, font=("Verdana", 9),
                             border_width = 1,
                             corner_radius = 3)

        entry.pack(pady=5)

        ok_button = ctk.CTkButton(top, text="OK", command=on_ok)
        ok_button.pack()

        top.geometry(f"+{x}+{y}")
        top.transient(self)  # Set the transient master to the treeview
        top.grab_set()  # Make the pop-up modal

        entry.focus_set()
        top.wait_window(top)  # Wait for the window to be closed

    def stop_editing(self, event=None):
        if self._entry is not None:
            new_value = self._entry.get()
            col = int(self._col.replace("#", ""))
            item = self.focus()

            if event and getattr(event, "keysym", "") == "Return" and item:
                current_values = self.item(item, "values")
                updated_values = [new_value if i == 0 else val for i, val in enumerate(current_values)]
                self.item(item, values=updated_values)

            self._entry.destroy()
            self._entry = None
            self._col = None


def bulk_rename_files(csv_file_path, brnroot_folder):
    df = pd.read_csv(csv_file_path, header=None, usecols=[0, 1])

    for index, row in df.iterrows():
        try:
            original_filename = str(row.iloc[0])  # First column
            new_filename = str(row.iloc[1])  # Second column

            for folder_path, _, filenames in os.walk(brnroot_folder):
                for filename in filenames:
                    if filename == original_filename:
                        original_filepath = os.path.join(folder_path, filename)
                        new_filepath = os.path.join(folder_path, new_filename)

                        if os.path.exists(original_filepath):
                            os.rename(original_filepath, new_filepath)
                            print(f"Renamed: {original_filepath} -> {new_filepath}")
                        else:
                            print(f"File not found: {original_filepath}")

        except Exception as e:
            print(f"Error: {e}")

def display_csv_in_treeview(treeview, csv_file_path):
    # Clear existing data in the Treeview
    for item in treeview.get_children():
        treeview.delete(item)

    # Load CSV file into a DataFrame without specifying column names
    df = pd.read_csv(csv_file_path, header=None, usecols=[0, 1])

    # Insert data into the Treeview
    for index, row in df.iterrows():
        treeview.insert("", "end", values=(row.iloc[0], row.iloc[1]))

def bulk_rename_gui():
    def browse_csv():
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        csv_entry.delete(0, tk.END)
        csv_entry.insert(0, file_path)

        # Update the displayed CSV data in the Treeview
        display_csv_in_treeview(treeview, file_path)

    def browse_folder():
        folder_path = filedialog.askdirectory()
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, folder_path)

    def start_bulk_rename():
        csv_path = csv_entry.get()
        folder_path = folder_entry.get()
        bulk_rename_files(csv_path, folder_path)

    # Create the main window
    brnroot = ctk.CTk()
    brnroot.title("Bulk File Rename")

    font_name = "Helvetica"
    font_size = 12
    
    # CSV File Entry
    csv_label = ctk.CTkLabel(brnroot, text="Filename CSV:", font=(font_name, font_size))
    csv_label.grid(row=0, column=0, padx=10, pady=5, sticky="e")

    csv_entry = ctk.CTkEntry(brnroot, width=200, height=20,font=(font_name,9),placeholder_text="..CSV containing old and new name",
                                border_width=1,
                                corner_radius=3)
    csv_entry.grid(row=0, column=1, padx=10, pady=5)

    csv_browse_button = ctk.CTkButton(brnroot, text="Browse", width = 80, command=browse_csv, font=(font_name, font_size))
    csv_browse_button.grid(row=0, column=2, padx=10, pady=5)

    # brnroot Folder Entry
    folder_label = ctk.CTkLabel(brnroot, text="Root Folder:", font=(font_name, font_size))
    folder_label.grid(row=1, column=0, padx=10, pady=5, sticky="e")

    folder_entry = ctk.CTkEntry(brnroot, width=200, height=20,font=(font_name,9),placeholder_text="..folder w/ files to rename",
                                border_width=1,
                                corner_radius=3)
    folder_entry.grid(row=1, column=1, padx=10, pady=5)

    folder_browse_button = ctk.CTkButton(brnroot, text="Browse", width = 80, command=browse_folder, font=(font_name, font_size))
    folder_browse_button.grid(row=1, column=2, padx=10, pady=5)

    # Treeview to display CSV data
    treeview = EditableTreeview(brnroot, columns=("Original Filename", "New Filename"), show="headings", height=10)
    treeview.heading("Original Filename", text="Original Filename")
    treeview.heading("New Filename", text="New Filename")
    treeview.grid(row=2, column=0, columnspan=3, pady=10)

    # Start Button
    start_button = ctk.CTkButton(brnroot, text="Start Rename", command=start_bulk_rename, font=(font_name, font_size))
    start_button.grid(row=3, column=1, pady=10)

    # Run the GUI
    brnroot.mainloop()

if __name__ == "__main__":
    bulk_rename_gui()