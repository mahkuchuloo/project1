import tkinter as tk
from tkinter import ttk

class ColumnMappingFrame(ttk.Frame):
    def __init__(self, master, target_columns, source_columns):
        super().__init__(master)
        self.source_columns = source_columns
        self.target_columns = target_columns
        self.mappings = {}
        self.entry_widgets = {}  # Store widgets for access by PlatformConfigDialog
        
        # Create canvas with fixed dimensions
        self.canvas = tk.Canvas(self, height=400, width=800)  # Set fixed height and width
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        # Configure canvas
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack scrollbar and canvas
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Create window in canvas for the frame
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # Bind canvas configuration
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        
        # Bind mouse wheel
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        self.create_widgets()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        # Update the width of the frame to fill the canvas
        self.canvas.itemconfig(self.canvas_frame, width=event.width)

    def create_widgets(self):
        # Configure column weights
        self.scrollable_frame.grid_columnconfigure(0, weight=1, minsize=200)  # Final File Column
        self.scrollable_frame.grid_columnconfigure(1, weight=1, minsize=300)  # Target Column
        self.scrollable_frame.grid_columnconfigure(2, weight=1, minsize=200)  # Default Value
        
        # Headers
        ttk.Label(self.scrollable_frame, text="Final File Column").grid(row=0, column=0, padx=5, sticky="ew")
        ttk.Label(self.scrollable_frame, text="Target Column").grid(row=0, column=1, padx=5, sticky="ew")
        ttk.Label(self.scrollable_frame, text="Default Value").grid(row=0, column=2, padx=5, sticky="ew")

        # Mapping rows
        for i, source_col in enumerate(self.source_columns, start=1):
            ttk.Label(self.scrollable_frame, text=source_col).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            
            target_combo = ttk.Combobox(self.scrollable_frame, values=["N/A"] + self.target_columns, width=40)
            target_combo.grid(row=i, column=1, padx=5, pady=2, sticky="ew")
            target_combo.set("N/A")
            
            default_entry = ttk.Entry(self.scrollable_frame, width=30)
            default_entry.grid(row=i, column=2, padx=5, pady=2, sticky="ew")
            
            self.mappings[source_col] = {
                "target": target_combo,
                "default": default_entry
            }
            
            # Store widgets for external access
            self.entry_widgets[source_col] = {
                "target": target_combo,
                "default": default_entry
            }

    def get_mappings(self):
        return {
            source_col: {
                "target": widgets["target"].get(),
                "default": widgets["default"].get()
            }
            for source_col, widgets in self.mappings.items()
        }

    def set_mappings(self, mappings):
        for source_col, mapping in mappings.items():
            if source_col in self.mappings:
                self.mappings[source_col]["target"].set(mapping["target"])
                self.mappings[source_col]["default"].delete(0, tk.END)  # Clear existing content
                self.mappings[source_col]["default"].insert(0, mapping["default"])
