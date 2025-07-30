import tkinter as tk
from tkinter import ttk
import json

class OutputSelectionDialog(tk.Toplevel):
    def __init__(self, parent, output_selections):
        super().__init__(parent)
        self.title("Select Output Columns")
        self.output_selections = output_selections
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Center the dialog
        window_width = 500
        window_height = 600
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        self.create_widgets()
        
        # Wait for window to be closed
        self.wait_window()
    
    def create_widgets(self):
        # Main frame with padding
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Notebook for tabs
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Core Data tab
        core_frame = ttk.Frame(notebook)
        notebook.add(core_frame, text="Core Data")
        
        # Core data sections
        core_sections = {
            'Basic Information': ['Basic Info'],
            'RFM Analysis': ['RFM Scores'],
            'Gift Information': ['First Gift Info', 'Last Gift Info', 'Largest Gift Info'],
            'Giving Details': ['Monthly Gift Info', 'Giving Segments', 'Platform Info'],
            'Contact Details': ['Contact Info', 'Geographic Info']
        }
        
        row = 0
        for section, options in core_sections.items():
            # Section label
            ttk.Label(core_frame, text=section, font=('TkDefaultFont', 10, 'bold')).grid(
                row=row, column=0, sticky='w', padx=5, pady=(10,5))
            row += 1
            
            # Options
            for option in options:
                if option in self.output_selections:
                    ttk.Checkbutton(core_frame, text=option, variable=self.output_selections[option]).grid(
                        row=row, column=0, sticky='w', padx=20, pady=2)
                    row += 1
        
        # Dictionaries tab
        dict_frame = ttk.Frame(notebook)
        notebook.add(dict_frame, text="Dictionaries")
        
        # Load dictionary configurations
        with open('rfm_lookup_dictionaries.json', 'r') as f:
            dictionaries = json.load(f)
        
        # Dictionary options
        row = 0
        for dictionary in dictionaries:
            name = dictionary['name']
            if name in self.output_selections:
                ttk.Checkbutton(dict_frame, text=name, variable=self.output_selections[name]).grid(
                    row=row, column=0, sticky='w', padx=5, pady=2)
                row += 1
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(button_frame, text="Select All", command=self.select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Deselect All", command=self.deselect_all).pack(side=tk.LEFT)
        
        # OK button
        ttk.Button(main_frame, text="OK", command=self.destroy).pack()
    
    def select_all(self):
        for var in self.output_selections.values():
            var.set(True)
    
    def deselect_all(self):
        for var in self.output_selections.values():
            var.set(False)
