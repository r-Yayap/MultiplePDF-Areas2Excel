import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
from CTkToolTip import CTkToolTip

from app.ui.constants import BUTTON_FONT

class EditableTreeview(ttk.Treeview):
    def __init__(self, root_window, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.root_window = root_window
        self._entry = None
        self._col = None

        self.bind("<Button-3>", self.show_context_menu)
        self.bind("<Double-Button-1>", self.on_double_click)

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
            self.context_menu.post(event.x_root, event.y_root)

    def remove_row(self):
        item = self.focus()
        if item:
            rectangle_index = self.index(item)
            self.root_window.pdf_viewer.canvas.delete(self.root_window.pdf_viewer.rectangle_list[rectangle_index])
            del self.root_window.pdf_viewer.areas[rectangle_index]
            del self.root_window.pdf_viewer.rectangle_list[rectangle_index]
            self.delete(item)
            self.root_window.update_areas_treeview()
            self.root_window.pdf_viewer.update_rectangles()
            print("Removed rectangle and updated canvas.")

    def edit_cell(self, item, col, _):
        def on_ok():
            new_value = entry_var.get()
            if new_value:
                current_values = list(self.item(item, "values"))
                current_values[col_index] = new_value
                self.item(item, values=tuple(current_values))
                self.update_areas_list()
            top.destroy()

        bbox = self.bbox(item, col)
        x, y, _, _ = bbox
        col_index = int(col.replace("#", "")) - 1

        top = ctk.CTkToplevel()
        top.title("Edit Cell")

        entry_var = ctk.StringVar()
        entry_var.set(self.item(item, "values")[col_index])

        entry = ctk.CTkEntry(
            top, justify="center", textvariable=entry_var,
            width=100, height=20, font=(BUTTON_FONT, 9),
            border_width=1, corner_radius=3
        )
        entry.pack(pady=5)

        ok_button = ctk.CTkButton(top, text="OK", command=on_ok)
        ok_button.pack()

        top.geometry(f"+{x}+{y}")
        top.grab_set()
        entry.focus_set()
        top.wait_window(top)

    def on_focus_out(self, _event):
        if self._entry is not None:
            self.stop_editing()

    def stop_editing(self, event=None):
        if self._entry is not None:
            new_value = self._entry.get()
            item = self.focus()
            if event and getattr(event, "keysym", "") == "Return" and item:
                current_values = self.item(item, "values")
                updated_values = [new_value if i == 0 else val for i, val in enumerate(current_values)]
                self.item(item, values=updated_values)
                self.update_areas_list()
            self._entry.destroy()
            self._entry = None
            self._col = None

    def update_areas_list(self):
        updated_areas = []
        for row_id in self.get_children():
            title, x0, y0, x1, y1 = self.item(row_id, "values")
            updated_areas.append({"title": title, "coordinates": [float(x0), float(y0), float(x1), float(y1)]})
        self.root_window.pdf_viewer.areas = updated_areas
        self.root_window.pdf_viewer.update_rectangles()


def create_tooltip(widget, message,
                   delay=0.3,
                   font=("Verdana", 9),
                   border_width=1,
                   border_color="gray50",
                   corner_radius=6,
                   justify="left"):
    return CTkToolTip(widget,
                      delay=delay,
                      justify=justify,
                      font=font,
                      border_width=border_width,
                      border_color=border_color,
                      corner_radius=corner_radius,
                      message=message)
