# sc_bulk_rename.py

import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from datetime import datetime

class EditableTreeview(ttk.Treeview):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._entry = None
        self._col = None

        # Bindings
        self.bind("<Button-3>", self.show_context_menu)
        self.bind("<Double-1>", self.on_double_click)
        self.bind("<FocusOut>", self.on_focus_out)

        # Context menu for removing rows
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Remove Row", command=self.remove_row)

    def show_context_menu(self, event):
        item = self.identify_row(event.y)
        if item:
            self.context_menu.post(event.x_root, event.y_root)

    def remove_row(self):
        item = self.focus()
        if item:
            self.delete(item)

    def on_double_click(self, event):
        item = self.focus()
        col = self.identify_column(event.x)
        if not item or col == "#0":
            return
        self._col = col
        col_index = int(col.replace("#", "")) - 1
        values = self.item(item, "values")
        if not values:
            return
        current_value = values[col_index]
        self.edit_cell(item, col_index, current_value, event)

    def on_focus_out(self, event):
        if self._entry:
            self.stop_editing()

    def edit_cell(self, item, col_index, old_value, click_event):
        x, y, width, height = self.bbox(item, f"#{col_index+1}")
        popup = ctk.CTkToplevel(self)
        popup.title("Edit Cell")
        popup.geometry(f"+{self.winfo_rootx()+x}+{self.winfo_rooty()+y+height}")

        var = tk.StringVar(value=old_value)
        entry = ctk.CTkEntry(popup, textvariable=var, width=150,
                             font=("Verdana", 9), border_width=1, corner_radius=3)
        entry.pack(padx=5, pady=5)
        entry.focus_set()

        def on_ok():
            new = var.get().strip()
            if new:
                vals = list(self.item(item, "values"))
                vals[col_index] = new
                self.item(item, values=vals)
            popup.destroy()

        btn = ctk.CTkButton(popup, text="OK", command=on_ok)
        btn.pack(pady=(0,5))

        popup.grab_set()
        popup.wait_window()

    def stop_editing(self, event=None):
        if self._entry:
            self._entry.destroy()
            self._entry = None
            self._col = None


def load_mapping(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xls", ".xlsx"):
        return pd.read_excel(path, header=None, usecols=[0,1], dtype=str)
    else:
        return pd.read_csv(path, header=None, usecols=[0,1], dtype=str)


def bulk_rename_files(mapping_path, root_folder):
    df = load_mapping(mapping_path)
    errors = []

    for _, row in df.iterrows():
        original = str(row.iloc[0]).strip()
        new_name = str(row.iloc[1]).strip()
        found = False

        for dirpath, _, files in os.walk(root_folder):
            for fname in files:
                if fname == original:
                    found = True
                    src = os.path.join(dirpath, fname)
                    dst = os.path.join(dirpath, new_name)
                    try:
                        os.rename(src, dst)
                    except Exception as e:
                        errors.append(f"{original} → {new_name}: {e}")
        if not found:
            errors.append(f"{original} → {new_name}: NOT FOUND")

    return errors


def display_mapping_in_treeview(treeview, mapping_path):
    for iid in treeview.get_children():
        treeview.delete(iid)

    df = load_mapping(mapping_path)
    for _, row in df.iterrows():
        treeview.insert("", "end", values=(row.iloc[0], row.iloc[1]))


def bulk_rename_gui():
    ctk.set_appearance_mode("dark")
    root = ctk.CTk()
    root.title("Bulk File Rename")

    font = ("Helvetica", 12)

    # Mapping file chooser
    lbl_map = ctk.CTkLabel(root, text="Mapping (CSV/Excel):", font=font)
    lbl_map.grid(row=0, column=0, padx=10, pady=5, sticky="e")
    mapping_entry = ctk.CTkEntry(root, width=300, placeholder_text="Select CSV or Excel file...", font=(font[0],9))
    mapping_entry.grid(row=0, column=1, padx=10, pady=5)
    btn_map = ctk.CTkButton(root, text="Browse", width=80, command=lambda: browse_mapping(mapping_entry, tree))
    btn_map.grid(row=0, column=2, padx=10, pady=5)

    # Root folder chooser
    lbl_folder = ctk.CTkLabel(root, text="Root Folder:", font=font)
    lbl_folder.grid(row=1, column=0, padx=10, pady=5, sticky="e")
    folder_entry = ctk.CTkEntry(root, width=300, placeholder_text="Select root folder...", font=(font[0],9))
    folder_entry.grid(row=1, column=1, padx=10, pady=5)
    btn_folder = ctk.CTkButton(root, text="Browse", width=80, command=lambda: browse_folder(folder_entry))
    btn_folder.grid(row=1, column=2, padx=10, pady=5)

    # Treeview for mapping
    tree = EditableTreeview(root, columns=("Original", "New"), show="headings", height=10)
    tree.heading("Original", text="Original Filename")
    tree.heading("New", text="New Filename")
    tree.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

    # Start Rename button
    btn_start = ctk.CTkButton(
        root, text="Start Rename", font=font,
        command=lambda: start_rename(mapping_entry.get(), folder_entry.get())
    )
    btn_start.grid(row=3, column=1, pady=10)

    # make columns expand
    root.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(2, weight=1)

    root.mainloop()


def browse_mapping(entry_widget, treeview):
    path = filedialog.askopenfilename(
        filetypes=[
            ("All supported", "*.csv;*.xlsx;*.xls"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx;*.xls")
        ]
    )
    if path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, path)
        display_mapping_in_treeview(treeview, path)


def browse_folder(entry_widget):
    path = filedialog.askdirectory()
    if path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, path)


def start_rename(mapping_path, folder_path):
    if not os.path.isfile(mapping_path):
        messagebox.showerror("Error", "Please select a valid CSV or Excel file.")
        return
    if not os.path.isdir(folder_path):
        messagebox.showerror("Error", "Please select a valid root folder.")
        return

    errors = bulk_rename_files(mapping_path, folder_path)
    if not errors:
        messagebox.showinfo("Success", "All files were renamed successfully.")
    else:
        # Compile error message
        msg = "Some renames failed or files were not found:\n\n" + "\n".join(errors)
        # Copy to clipboard
        temp = tk.Tk()
        temp.withdraw()
        temp.clipboard_clear()
        temp.clipboard_append(msg)
        temp.update()
        temp.destroy()
        # Show warning
        messagebox.showwarning("Partial Success", msg)

if __name__ == "__main__":
    bulk_rename_gui()
