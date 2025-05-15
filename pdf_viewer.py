# pdf_viewer.py

import fitz  # PyMuPDF
import customtkinter as ctk
import tkinter as tk
from tkinter import Menu
from constants import *
from tkinter.simpledialog import askstring  # For custom title input
from tkinterdnd2 import DND_ALL
import os

class PDFViewer:
    def __init__(self, parent, master):
        self.parent = parent  # `parent` is the XtractorGUI instance
        self.canvas = ctk.CTkCanvas(master, width=CANVAS_WIDTH, height=CANVAS_HEIGHT)

        # Add placeholder text
        self.placeholder_text_id = None
        self.show_placeholder()

        self.v_scrollbar = ctk.CTkScrollbar(master, orientation="vertical", command=self.canvas.yview,
                                            height=CANVAS_HEIGHT)


        self.h_scrollbar = ctk.CTkScrollbar(master, orientation="horizontal", command=self.canvas.xview,
                                            width=CANVAS_WIDTH)


        self.pdf_document = None
        self.page = None
        self.current_zoom = CURRENT_ZOOM
        self.areas = []
        self.rectangle_list = []
        self.current_rectangle = None
        self.original_coordinates = None
        self.canvas_image = None  # Holds the current PDF page image to prevent garbage collection
        self.resize_job = None  # Track the delayed update job

        # Initialize selected rectangle ID and title dictionary
        self.selected_rectangle = None
        self.selected_rectangle_id = None
        self.rectangle_titles = {}  # Dictionary to store {rectangle_id: title}



        # Create main context menu
        self.context_menu = Menu(self.canvas, tearoff=0)

        # Set Title submenu for title options
        self.set_title_menu = Menu(self.context_menu, tearoff=0)
        self.set_title_menu.add_command(label="Drawing No", command=lambda: self.set_rectangle_title("Drawing No"))
        self.set_title_menu.add_command(label="Drawing Title",
                                        command=lambda: self.set_rectangle_title("Drawing Title"))
        self.set_title_menu.add_command(label="Revision Description",
                                        command=lambda: self.set_rectangle_title("Revision Description"))
        self.set_title_menu.add_command(label="Custom...", command=self.set_custom_title)

        # Add Set Title submenu to context menu
        self.context_menu.add_cascade(label="Set Title", menu=self.set_title_menu)

        # Add Delete Rectangle option to the context menu
        self.context_menu.add_command(label="Delete Rectangle", command=self.delete_selected_rectangle)

        # Bind canvas resize and mouse events
        self.canvas.master.bind("<Configure>", lambda event: self.resize_canvas(event.width, event.height))
        self.canvas.bind("<ButtonPress-1>", self.start_rectangle)
        self.canvas.bind("<B1-Motion>", self.draw_rectangle)
        self.canvas.bind("<ButtonRelease-1>", self.end_rectangle)

        # Right-click context menu events
        self.canvas.bind("<Button-3>", self.show_context_menu)
        self.canvas.bind("<ButtonRelease-3>", self.show_context_menu)  # Alternative right-click event

        # Initialize selection state
        self.selected_rectangle_id = None
        self.selected_rectangle_index = None
        self.selected_rectangle_original_color = "red"  # Default color for rectangles

        # Scroll and zoom events
        self.canvas.bind("<MouseWheel>", self.handle_mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", self.handle_mousewheel)  # Shift for horizontal scroll
        self.canvas.bind("<Control-MouseWheel>", self.handle_mousewheel)  # Ctrl for zoom

        # Added for Revision History Table
        self.selection_mode = "area"  # Options: "area" or "revision"
        self.revision_area = None  # Holds single revision table rectangle
        self.revision_rectangle_id = None
        self.selection_mode = "area"  # Can be "area" or "revision"

        self.canvas.drop_target_register(DND_ALL)
        self.canvas.dnd_bind('<<Drop>>', self.handle_pdf_drop)

    def show_placeholder(self):
        """Displays the drag & drop hint centered in the canvas."""
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        # Fallback if size is not yet available (initial state)
        if canvas_width <= 1 or canvas_height <= 1:
            canvas_width = CANVAS_WIDTH - 200
            canvas_height = CANVAS_HEIGHT

        if self.placeholder_text_id:
            self.canvas.delete(self.placeholder_text_id)

        self.placeholder_text_id = self.canvas.create_text(
            canvas_width // 2,
            canvas_height // 2,
            text="DRAG & DROP\nSample PDF here",
            fill="gray70",
            font=("Arial", 20, "italic"),
            anchor="center"
        )

    def handle_pdf_drop(self, event):
        path = event.data.strip().replace("{", "").replace("}", "")
        if path.lower().endswith(".pdf") and os.path.isfile(path):
            self.parent.recent_pdf_path = path  # update recent path
            self.display_pdf(path)
            print(f"🪂 Dropped PDF loaded in canvas: {path}")
        else:
            print(f"⚠️ Dropped item is not a valid PDF: {path}")

    def set_custom_title(self):
        """Prompts user for a custom title and assigns it to the selected rectangle."""
        custom_title = askstring("Custom Title", "Enter a custom title:")
        if custom_title:
            self.set_rectangle_title(custom_title)  # Use the input title for the selected rectangle

    def handle_mousewheel(self, event):
        """Handles mouse wheel scrolling with Shift and Control modifiers."""
        if event.state & 0x1:  # Shift pressed for horizontal scrolling
            self.canvas.xview_scroll(-1 * int(event.delta / 120), "units")
        elif event.state & 0x4:  # Ctrl pressed for zoom
            if event.delta > 0:
                self.zoom_in(0.1)  # Zoom in by a small increment
            else:
                self.zoom_out(0.1)  # Zoom out by a small increment
            # Notify the GUI to update the zoom slider
            self.parent.update_zoom_slider(self.current_zoom)
        else:  # Regular vertical scrolling
            self.canvas.yview_scroll(-1 * int(event.delta / 120), "units")

    def zoom_in(self, increment=0.1):
        """Zoom in by increasing the current zoom level and refreshing the display."""
        self.current_zoom += increment
        self.update_display()

    def zoom_out(self, decrement=0.1):
        """Zoom out by decreasing the current zoom level and refreshing the display."""
        self.current_zoom = max(0.1, self.current_zoom - decrement)  # Prevent excessive zooming out
        self.update_display()

    def close_pdf(self):
        """Closes the displayed PDF and clears the canvas."""
        # Remove any displayed image from the canvas
        self.canvas.delete("pdf_image")

        # Close the PDF document if it is open
        if self.pdf_document:
            self.pdf_document.close()
            print("PDF document closed.")

        # Reset the pdf_document attribute to None to indicate no PDF is open
        self.pdf_document = None
        # Restore placeholder
        self.show_placeholder()

    def display_pdf(self, pdf_path):
        """Loads and displays the first page of a PDF document."""
        self.pdf_document = fitz.open(pdf_path)
        if self.pdf_document.page_count > 0:
            self.page = self.pdf_document[0]  # Display the first page by default
            # Remove placeholder if present
            if self.placeholder_text_id:
                self.canvas.delete(self.placeholder_text_id)
                self.placeholder_text_id = None

            self.pdf_width = int(self.page.rect.width)
            self.pdf_height = int(self.page.rect.height)
            # Update the display
            self.update_display()
            # Set the initial view to the top-left corner of the PDF
            self.canvas.xview_moveto(1)  # Horizontal scroll to start
            self.canvas.yview_moveto(1)  # Vertical scroll to start
        else:
            self.pdf_document = None
            print("Error: PDF has no pages.")

    def update_display(self):
        """Updates the canvas to display the current PDF page with zoom and scroll configurations."""

        # Only proceed if a valid page is loaded
        if not self.page:
            print("Error updating display: No valid page loaded.")
            return

        # Check if there is a valid PDF page to display
        if self.page is None:
            print("No valid page to display.")
            return

        try:
            # Clear any existing content on the canvas
            self.canvas.delete("all")

            # Generate a pixmap from the PDF page at the current zoom level
            pix = self.page.get_pixmap(matrix=fitz.Matrix(self.current_zoom, self.current_zoom))
            img = pix.tobytes("ppm")
            img_tk = tk.PhotoImage(data=img)

            # Display the updated image on the canvas
            self.canvas.create_image(0, 0, anchor=tk.NW, image=img_tk, tags="pdf_image")

            # Keep a reference to the image to prevent garbage collection
            self.canvas_image = img_tk

            # Calculate the zoomed dimensions
            zoomed_width = int(self.pdf_width * self.current_zoom)
            zoomed_height = int(self.pdf_height * self.current_zoom)

            # Configure the scroll region of the canvas to match the zoomed dimensions
            self.canvas.config(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set,
                               scrollregion=(0, 0, zoomed_width, zoomed_height))

        except ValueError as e:
            print(f"Error updating display: {e}")

        # Update any rectangle overlays or additional graphics
        self.update_rectangles()

    def start_rectangle(self, event):
        """Begins the rectangle selection process on mouse press."""
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.original_coordinates = [x, y]
        color = "red" if self.selection_mode == "area" else "green"
        self.current_rectangle = self.canvas.create_rectangle(x, y, x, y, outline=color, width=2)

    def draw_rectangle(self, event):
        """Adjusts the rectangle dimensions as the mouse is dragged."""
        if self.current_rectangle:
            x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
            self.canvas.coords(self.current_rectangle, self.original_coordinates[0], self.original_coordinates[1], x, y)
            self.auto_scroll_canvas(event.x, event.y)


    def end_rectangle(self, event):
        """Finalizes the rectangle selection and saves its coordinates."""
        if self.current_rectangle:
            x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
            self.canvas.coords(self.current_rectangle, self.original_coordinates[0], self.original_coordinates[1], x, y)
            bbox = self.canvas.bbox(self.current_rectangle)
            if not bbox:
                print("Error: Failed to retrieve bounding box coordinates")
                return

            x0, y0, x1, y1 = bbox
            adjusted_coordinates = [
                x0 / self.current_zoom,
                y0 / self.current_zoom,
                x1 / self.current_zoom,
                y1 / self.current_zoom
            ]

            if self.selection_mode == "area":
                self.areas.append({
                    "coordinates": adjusted_coordinates,
                    "title": f"Rectangle {len(self.areas) + 1}"
                })
                self.rectangle_list.append(self.current_rectangle)
                self.parent.update_areas_treeview()
            else:
                # Only allow one revision area
                if self.revision_rectangle_id:
                    self.canvas.delete(self.revision_rectangle_id)
                self.revision_area = {"coordinates": adjusted_coordinates, "title": "Revision Table"}
                self.revision_rectangle_id = self.current_rectangle
                print("Set Revision Table Rectangle:", self.revision_area)

        self.current_rectangle = None


    def auto_scroll_canvas(self, x, y):
        """Auto-scrolls the canvas if the mouse is near the edges during a drag operation."""
        global scroll_counter
        scroll_margin = 20  # Distance from the canvas edge to start scrolling

        # Only scroll every SCROLL_INCREMENT_THRESHOLD calls
        if scroll_counter < SCROLL_INCREMENT_THRESHOLD:
            scroll_counter += 1
            return  # Skip scrolling this call

        scroll_counter = 0  # Reset counter after threshold is reached

        # Check if the mouse is close to the edges and scroll in small increments
        if x >= self.canvas.winfo_width() - scroll_margin:
            self.canvas.xview_scroll(1, "units")
        elif x <= scroll_margin:
            self.canvas.xview_scroll(-1, "units")

        if y >= self.canvas.winfo_height() - scroll_margin:
            self.canvas.yview_scroll(1, "units")
        elif y <= scroll_margin:
            self.canvas.yview_scroll(-1, "units")

    def clear_areas(self):
        """Clears all rectangles, area selections, and Treeview entries from the canvas."""

        # Clear all rectangles from the canvas
        for rect_id in self.rectangle_list:
            self.canvas.delete(rect_id)
        self.rectangle_list.clear()

        # Clear the areas list
        self.areas.clear()

        # Clear the areas Treeview if it exists
        if hasattr(self, 'areas_tree') and self.areas_tree:
            for item in self.areas_tree.get_children():
                self.areas_tree.delete(item)

        # Update the canvas display to reflect changes
        self.update_display()

        self.parent.update_areas_treeview()  # Clear the table view as well

        # Optional: Print statement for debugging
        print("Cleared All Areas")

    def update_rectangles(self):
        """Redraws area and revision rectangles on the canvas."""
        for rect_id in self.rectangle_list:
            self.canvas.delete(rect_id)
        self.rectangle_list.clear()

        if self.revision_rectangle_id:
            self.canvas.delete(self.revision_rectangle_id)
            self.revision_rectangle_id = None

        for area in self.areas:
            x0, y0, x1, y1 = [coord * self.current_zoom for coord in area["coordinates"]]
            rect_id = self.canvas.create_rectangle(x0, y0, x1, y1, outline="red", width=2)
            self.rectangle_list.append(rect_id)

        # Draw revision area if present
        if self.revision_area:
            x0, y0, x1, y1 = [c * self.current_zoom for c in self.revision_area["coordinates"]]
            self.revision_rectangle_id = self.canvas.create_rectangle(x0, y0, x1, y1, outline="green", width=2)

        # Update Treeview only for normal areas
        self.parent.update_areas_treeview()

    def set_zoom(self, zoom_level):
        """Updates the zoom level and refreshes the display."""
        self.current_zoom = zoom_level
        self.update_display()  # Refresh the display with the new zoom level

    def resize_canvas(self, total_width, total_height, x_offset=0):
        """Resizes canvas and scrollbars based on available space, respecting left offset."""

        # Reserve space for margins and scrollbars
        canvas_margin = 20
        scrollbar_thickness = 14

        # Dynamically calculate canvas size
        canvas_width = max(200, total_width - x_offset - CANVAS_EXTRA_MARGIN - SCROLLBAR_THICKNESS)
        canvas_height = max(200, total_height - CANVAS_TOP_MARGIN - CANVAS_BOTTOM_MARGIN)

        # Reposition and resize canvas
        self.canvas.place_configure(x=x_offset, y=CANVAS_TOP_MARGIN)
        self.canvas.config(width=canvas_width, height=canvas_height)

        # Reposition and resize scrollbars
        self.v_scrollbar.place_configure(x=x_offset + canvas_width  +4, y=CANVAS_TOP_MARGIN)
        self.v_scrollbar.configure(height=canvas_height)

        self.h_scrollbar.place_configure(x=x_offset, y=CANVAS_TOP_MARGIN + canvas_height + 7)
        self.h_scrollbar.configure(width=canvas_width)

        # Scale canvas contents
        if self.canvas.find_all():
            scale_x = canvas_width / CANVAS_WIDTH
            scale_y = canvas_height / CANVAS_HEIGHT
            self.canvas.scale("all", 0, 0, scale_x, scale_y)
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        if self.pdf_document is None:
            self.show_placeholder()

    def _perform_resize(self):
        """Scales the canvas dynamically when the window resizes, only if elements exist."""
        if self.canvas.find_all():  # Only scale if there are elements on the canvas
            scale_x = self.canvas.winfo_width() / CANVAS_WIDTH
            scale_y = self.canvas.winfo_height() / CANVAS_HEIGHT
            self.canvas.scale("all", 0, 0, scale_x, scale_y)
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))  # Adjust scrolling
        else:
            print("Skipping scaling: No elements on the canvas yet.")

    def show_context_menu(self, event):
        """Displays context menu and highlights the rectangle if right-click occurs near the edge."""
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        edge_tolerance = 5  # Set the edge tolerance for detecting clicks near the boundary

        # Clear previous selection if any
        self.clear_selection()

        # Iterate over rectangles to find one that has been clicked near its edge
        for index, rect_id in enumerate(self.rectangle_list):
            bbox = self.canvas.bbox(rect_id)
            if bbox:
                x0, y0, x1, y1 = bbox

                # Check if the click is near the left or right edge within the tolerance
                near_left_edge = abs(x - x0) <= edge_tolerance and y0 <= y <= y1
                near_right_edge = abs(x - x1) <= edge_tolerance and y0 <= y <= y1

                # Check if the click is near the top or bottom edge within the tolerance
                near_top_edge = abs(y - y0) <= edge_tolerance and x0 <= x <= x1
                near_bottom_edge = abs(y - y1) <= edge_tolerance and x0 <= x <= x1

                # If click is near any edge, select this rectangle
                if near_left_edge or near_right_edge or near_top_edge or near_bottom_edge:
                    self.selected_rectangle_id = rect_id
                    self.selected_rectangle_index = index
                    self.selected_rectangle_original_color = self.canvas.itemcget(rect_id, "outline")

                    # Highlight the selected rectangle with a different color
                    self.canvas.itemconfig(rect_id, outline="blue")
                    print(f"Selected Rectangle at Index {index} with ID: {rect_id}")
                    break

        # Show context menu if a rectangle was selected by edge detection
        if self.selected_rectangle_id is not None:
            self.context_menu.post(event.x_root, event.y_root)
        else:
            # Hide menu if no rectangle edge was clicked
            print("No rectangle edge detected, context menu will not be shown.")
            self.context_menu.unpost()

    def clear_selection(self):
        """Clears the selection by resetting the color of the previously selected rectangle."""
        if self.selected_rectangle_id is not None:
            # Reset the previously selected rectangle's color
            self.canvas.itemconfig(self.selected_rectangle_id, outline=self.selected_rectangle_original_color)
            self.selected_rectangle_id = None
            self.selected_rectangle_index = None

    def set_rectangle_title(self, title):
        """Assigns a selected title to the currently selected rectangle and updates the Treeview."""
        if self.selected_rectangle_index is not None:
            # Update the title directly in `self.areas` based on the rectangle index
            self.areas[self.selected_rectangle_index]["title"] = title  # Update title in `self.areas`
            print(f"Title '{title}' set for rectangle at Index: {self.selected_rectangle_index}")

            # Update the Treeview to reflect the new title
            self.parent.update_areas_treeview()
        else:
            print("No rectangle selected. Title not set.")

    def delete_selected_rectangle(self):
        """Deletes the selected rectangle from the canvas and updates the list of areas."""
        if self.selected_rectangle_id:
            try:
                # Find the index of the selected rectangle in rectangle_list
                index = self.rectangle_list.index(self.selected_rectangle_id)

                # Delete the rectangle from canvas and remove from lists
                self.canvas.delete(self.selected_rectangle_id)
                del self.rectangle_list[index]
                del self.areas[index]

                # Update the Treeview and clear selection
                self.parent.update_areas_treeview()
                self.selected_rectangle_id = None
                print("Rectangle deleted.")

                # Reassign titles to reflect the new order
                for index, area in enumerate(self.areas):
                    area["title"] = f"Rectangle {index + 1}"  # Update titles in `areas`
                self.parent.update_areas_treeview()  # Refresh Treeview to reflect updated titles


            except ValueError:
                print("Selected rectangle ID not found in the rectangle list.")
        else:
            print("No rectangle selected for deletion.")
