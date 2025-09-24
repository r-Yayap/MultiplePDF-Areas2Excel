# app/ui/ui_utils.py
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

        # --- context & edit ---
        self.bind("<Button-3>", self.show_context_menu)
        self.bind("<Double-Button-1>", self.on_double_click)

        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Remove Row", command=self.remove_row)

        # --- drag-to-reorder state ---
        self._drag_iid = None
        self._press_iid = None
        self._press_y = 0
        self._dragging = False
        self._drag_threshold = 4  # pixels before we consider it a drag

        # Keep default selection behavior; add our handlers instead of replacing
        self.bind("<ButtonPress-1>", self._on_press, add="+")
        self.bind("<B1-Motion>", self._on_motion, add="+")
        self.bind("<ButtonRelease-1>", self._on_release, add="+")

    # ---------------- selection + editing ----------------
    def on_double_click(self, event):
        # Let double-click always focus/select the clicked row first
        iid = self.identify_row(event.y)
        if iid:
            self.focus(iid)
            self.selection_set(iid)

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
            self.focus(item)
            self.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def remove_row(self):
        item = self.focus()
        if item:
            rectangle_index = self.index(item)
            # remove rectangle on canvas
            try:
                self.root_window.pdf_viewer.canvas.delete(
                    self.root_window.pdf_viewer.rectangle_list[rectangle_index]
                )
                del self.root_window.pdf_viewer.rectangle_list[rectangle_index]
            except Exception:
                pass

            # remove from model
            try:
                del self.root_window.pdf_viewer.areas[rectangle_index]
            except Exception:
                pass

            # remove from tree + refresh
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

        # column index (0-based into values tuple)
        col_index = int(col.replace("#", "")) - 1

        # popup editor
        top = ctk.CTkToplevel()
        top.title("Edit Cell")

        entry_var = ctk.StringVar()
        entry_var.set(self.item(item, "values")[col_index])

        entry = ctk.CTkEntry(
            top, justify="center", textvariable=entry_var,
            width=160, height=28, font=(BUTTON_FONT, 9),
            border_width=1, corner_radius=6
        )
        entry.pack(padx=10, pady=(12, 8))

        ok_button = ctk.CTkButton(top, text="OK", command=on_ok, width=90)
        ok_button.pack(pady=(0, 12))

        # center near mouse
        top.geometry(f"+{top.winfo_pointerx()}+{top.winfo_pointery()}")
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
        """Write current rows (in current on-screen order) back to the model and redraw."""
        updated_areas = []
        for row_id in self.get_children():
            title, x0, y0, x1, y1 = self.item(row_id, "values")
            updated_areas.append({"title": title, "coordinates": [float(x0), float(y0), float(x1), float(y1)]})
        self.root_window.pdf_viewer.areas = updated_areas
        self.root_window.pdf_viewer.update_rectangles()

    # ---------------- drag-to-reorder ----------------
    def _on_press(self, event):
        iid = self.identify_row(event.y)
        self._press_iid = iid
        self._drag_iid = iid
        self._press_y = event.y
        self._dragging = False  # not dragging yet

        if iid:
            # ensure normal selection behavior
            self.focus(iid)
            self.selection_set(iid)

    def _on_motion(self, event):
        # if editing popup is up, don't drag
        if self._entry is not None:
            return
        if not self._press_iid:
            return

        # start dragging only after a small move
        if not self._dragging:
            if abs(event.y - self._press_y) < self._drag_threshold:
                return
            self._dragging = True
            try:
                self.configure(cursor="fleur")
            except Exception:
                pass

        # while dragging, move row visually
        target = self.identify_row(event.y)
        if target and target != self._drag_iid:
            try:
                self.move(self._drag_iid, "", self.index(target))
            except Exception:
                pass

    def _on_release(self, _event):
        was_dragging = self._dragging
        self._dragging = False
        self._press_iid = None
        self._drag_iid = None
        try:
            self.configure(cursor="")
        except Exception:
            pass

        # only rewrite model if we actually dragged
        if was_dragging:
            self._apply_current_order_to_model()

    def _apply_current_order_to_model(self):
        """Sync current on-screen order -> model, redraw rectangles, and rebuild the table mapping."""
        updated_areas = []
        for row_id in self.get_children():
            vals = self.item(row_id, "values")
            if not vals:
                continue
            title, x0, y0, x1, y1 = vals
            updated_areas.append({
                "title": title,
                "coordinates": [float(x0), float(y0), float(x1), float(y1)]
            })

        # write back to domain
        self.root_window.pdf_viewer.areas = updated_areas
        # redraw rectangles with the new order
        self.root_window.pdf_viewer.update_rectangles()
        # rebuild the tree to refresh the GUI's index mapping used elsewhere
        self.root_window.update_areas_treeview()


class CTkOptionMenuNoArrow(ctk.CTkFrame):
    """Arrow-less dropdown that mirrors CTkOptionMenu's basic API."""
    def __init__(self, master, values, variable=None, command=None,
                 width=120, height=28, font=None,
                 fg_color=None, hover_color=None,
                 item_height=28, dropdown_max_items=8, **kwargs):
        super().__init__(master, width=width, height=height, **kwargs)
        self._values = list(values)
        self._command = command
        self._font = font or ("Segoe UI", 10)
        self._item_height = item_height
        self._dropdown_max_items = dropdown_max_items

        if variable is None:
            self._var = ctk.StringVar(value=self._values[0] if self._values else "")
        else:
            self._var = variable
            if not self._var.get() and self._values:
                self._var.set(self._values[0])

        self.pack_propagate(False)

        self.button = ctk.CTkButton(
            self,
            textvariable=self._var,
            font=self._font,
            width=width,
            height=height,
            fg_color=fg_color,
            hover_color=hover_color,
            corner_radius=6,
            command=self._toggle_menu
        )
        self.button.pack(fill="both", expand=True)

        # track var changes to keep button label in sync even if set externally
        self._var.trace_add("write", lambda *_: self._sync_selection_highlight())

        # dropdown popover (created on demand)
        self._menu = None

    # ---- public API-ish ----
    def set(self, value: str):
        if value in self._values:
            self._var.set(value)
            if callable(self._command):
                self._command(value)

    def configure(self, **kwargs):
        # pass sizing/styling through to the inner button where sensible
        if "font" in kwargs:
            self._font = kwargs.pop("font")
            self.button.configure(font=self._font)
        self.button.configure(**kwargs)
        super().configure(**kwargs)

    def cget(self, option):
        try:
            return self.button.cget(option)
        except Exception:
            return super().cget(option)

    # ---- internals ----
    def _toggle_menu(self):
        if self._menu and self._menu.winfo_exists():
            self._close_menu()
        else:
            self._open_menu()

    def _open_menu(self):
        if not self._values:
            return

        # build a tiny borderless top-level under the button
        self._menu = ctk.CTkToplevel(self)
        self._menu.overrideredirect(True)
        self._menu.attributes("-topmost", True)

        # position it just below
        bx = self.winfo_rootx()
        by = self.winfo_rooty() + self.winfo_height()
        self._menu.geometry(f"+{bx}+{by}")

        # container
        frame = ctk.CTkFrame(self._menu, corner_radius=6)
        frame.pack(fill="both", expand=True)

        # optional scrolling if many items
        max_items = min(self._dropdown_max_items, len(self._values))
        menu_height = max_items * self._item_height
        frame.configure(height=menu_height)

        # add items
        self._item_buttons = []
        for val in self._values:
            btn = ctk.CTkButton(
                frame,
                text=val,
                height=self._item_height,
                font=self._font,
                anchor="w",
                fg_color="transparent",
                hover_color="#2b2b2b",
                command=lambda v=val: self._pick(v)
            )
            btn.pack(fill="x", padx=6, pady=2)
            self._item_buttons.append((val, btn))

        self._sync_selection_highlight()

        # close when clicking elsewhere or hitting Escape
        self._outside_bind = self.winfo_toplevel().bind_all("<Button-1>", self._maybe_outside_click, add="+")
        self._esc_bind = self.winfo_toplevel().bind_all("<Escape>", lambda e: self._close_menu(), add="+")

    def _close_menu(self):
        if self._menu and self._menu.winfo_exists():
            self._menu.destroy()
        self._menu = None
        # remove global binds
        try:
            self.winfo_toplevel().unbind_all("<Button-1>")
            self.winfo_toplevel().unbind_all("<Escape>")
        except Exception:
            pass

    def _maybe_outside_click(self, event):
        if not self._menu or not self._menu.winfo_exists():
            return
        # if click is outside the menu and the button, close
        x, y = event.x_root, event.y_root
        in_menu = (self._menu.winfo_rootx() <= x <= self._menu.winfo_rootx()+self._menu.winfo_width() and
                   self._menu.winfo_rooty() <= y <= self._menu.winfo_rooty()+self._menu.winfo_height())
        in_btn = (self.winfo_rootx() <= x <= self.winfo_rootx()+self.winfo_width() and
                  self.winfo_rooty() <= y <= self.winfo_rooty()+self.winfo_height())
        if not (in_menu or in_btn):
            self._close_menu()

    def _pick(self, value):
        self._var.set(value)
        if callable(self._command):
            self._command(value)
        self._close_menu()

    def _sync_selection_highlight(self):
        if not getattr(self, "_item_buttons", None):
            return
        current = self._var.get()
        for val, btn in self._item_buttons:
            if val == current:
                btn.configure(fg_color="#3a3a3a")  # selected bg
            else:
                btn.configure(fg_color="transparent")


def create_tooltip(widget, message, delay=0.3,
                   font=("Verdana", 9),
                   border_width=1,
                   border_color="gray50",
                   corner_radius=6,
                   justify="left"):
    return CTkToolTip(
        widget,
        delay=delay,
        justify=justify,
        font=font,
        border_width=border_width,
        border_color=border_color,
        corner_radius=corner_radius,
        message=message,
    )
