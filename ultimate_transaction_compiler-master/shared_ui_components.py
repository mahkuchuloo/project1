import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import queue
import os
import subprocess
from utils import configure_logging
from shared_config import FINAL_FILE_COLUMNS

class TooltipLabel(ttk.Label):
    def __init__(self, parent, tooltip_text, **kwargs):
        super().__init__(parent, **kwargs)
        self.tooltip = None
        self.tooltip_text = tooltip_text
        self.bind('<Enter>', self.show_tooltip)
        self.bind('<Leave>', self.hide_tooltip)

    def show_tooltip(self, event=None):
        x, y, _, _ = self.bbox("insert")
        x += self.winfo_rootx() + 25
        y += self.winfo_rooty() + 25

        self.tooltip = tk.Toplevel()
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = ttk.Label(self.tooltip, text=self.tooltip_text, 
                         justify=tk.LEFT, background="#ffffe0", 
                         relief=tk.SOLID, borderwidth=1)
        label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

class LookupDictionaryConfigDialog(tk.Toplevel):
    def __init__(self, parent, existing_dictionaries):
        super().__init__(parent)
        self.title("Lookup Dictionary Configuration")
        self.dictionaries = existing_dictionaries.copy()
        self.available_columns = FINAL_FILE_COLUMNS
        self.create_widgets()

    def create_widgets(self):
        # Dictionary list
        self.dict_listbox = tk.Listbox(self, width=30)
        self.dict_listbox.pack(side=tk.LEFT, fill=tk.Y)
        self.dict_listbox.bind('<<ListboxSelect>>', self.on_dictionary_select)

        # Dictionary details frame
        details_frame = ttk.Frame(self)
        details_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Name row with tooltip
        name_frame = ttk.Frame(details_frame)
        name_frame.grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(name_frame, text="Dictionary Name:").pack(side=tk.LEFT)
        self.name_entry = ttk.Entry(name_frame)
        self.name_entry.pack(side=tk.LEFT, padx=5)
        TooltipLabel(name_frame, text="ℹ", 
                    tooltip_text="The unique name for this dictionary").pack(side=tk.LEFT)

        # File row with tooltip
        file_frame = ttk.Frame(details_frame)
        file_frame.grid(row=1, column=0, columnspan=3, sticky="w")
        ttk.Label(file_frame, text="Dictionary File:").pack(side=tk.LEFT)
        self.file_entry = ttk.Entry(file_frame)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_file).pack(side=tk.LEFT)
        TooltipLabel(file_frame, text="ℹ", 
                    tooltip_text="Excel file containing the dictionary values").pack(side=tk.LEFT)

        # Multiple Values Dictionary checkbox with tooltip
        self.use_multiple_values = tk.BooleanVar()
        multiple_values_frame = ttk.Frame(details_frame)
        multiple_values_frame.grid(row=2, column=0, columnspan=2, sticky="w")
        self.multiple_values_checkbox = ttk.Checkbutton(multiple_values_frame, 
            text="Multiple Values Dictionary", 
            variable=self.use_multiple_values,
            command=self.toggle_multiple_values)
        self.multiple_values_checkbox.pack(side=tk.LEFT)
        TooltipLabel(multiple_values_frame, text="ℹ", 
                    tooltip_text="When enabled, all columns except the first will be treated as values.\n"
                                "First column is the key, other columns will be used as is.").pack(side=tk.LEFT)

        # Lookup Column frame (always visible)
        lookup_frame = ttk.Frame(details_frame)
        lookup_frame.grid(row=3, column=0, columnspan=2, sticky="w")
        ttk.Label(lookup_frame, text="Lookup Column:").pack(side=tk.LEFT)
        self.lookup_column_combo = ttk.Combobox(lookup_frame, values=self.available_columns)
        self.lookup_column_combo.pack(side=tk.LEFT, padx=5)
        TooltipLabel(lookup_frame, text="ℹ", 
                    tooltip_text="The column to match against the dictionary key").pack(side=tk.LEFT)

        # Single Value Dictionary Options frame
        self.single_value_frame = ttk.Frame(details_frame)
        self.single_value_frame.grid(row=4, column=0, columnspan=2, sticky="w")

        output_frame = ttk.Frame(self.single_value_frame)
        output_frame.pack(fill=tk.X)
        ttk.Label(output_frame, text="Output Column:").pack(side=tk.LEFT)
        self.output_column_entry = ttk.Entry(output_frame)
        self.output_column_entry.pack(side=tk.LEFT, padx=5)
        TooltipLabel(output_frame, text="ℹ", 
                    tooltip_text="The column name where matched values will be stored").pack(side=tk.LEFT)

        # Options frame for checkboxes that should be hidden in multiple values mode
        self.options_frame = ttk.Frame(details_frame)
        self.options_frame.grid(row=5, column=0, columnspan=2, sticky="w")

        # Post-merger name logic checkbox with tooltip
        self.use_post_merger = tk.BooleanVar()
        post_merger_frame = ttk.Frame(self.options_frame)
        post_merger_frame.pack(fill=tk.X)
        self.post_merger_checkbox = ttk.Checkbutton(post_merger_frame, 
            text="Use Post-Merger Name Logic", 
            variable=self.use_post_merger,
            command=self.toggle_values_frame)
        self.post_merger_checkbox.pack(side=tk.LEFT)
        TooltipLabel(post_merger_frame, text="ℹ", 
                    tooltip_text="Enable support for pre/post merger name matching").pack(side=tk.LEFT)

        # Zip code validation checkbox with tooltip
        self.use_zip_validation = tk.BooleanVar()
        zip_frame = ttk.Frame(self.options_frame)
        zip_frame.pack(fill=tk.X)
        self.zip_validation_checkbox = ttk.Checkbutton(zip_frame, 
            text="Use Zip Code Validation", 
            variable=self.use_zip_validation)
        self.zip_validation_checkbox.pack(side=tk.LEFT)
        TooltipLabel(zip_frame, text="ℹ", 
                    tooltip_text="Validate zip codes when matching").pack(side=tk.LEFT)

        # Include in Last Gift Values checkbox with tooltip
        self.include_in_last_gift = tk.BooleanVar()
        last_gift_frame = ttk.Frame(self.options_frame)
        last_gift_frame.pack(fill=tk.X)
        self.include_in_last_gift_checkbox = ttk.Checkbutton(last_gift_frame, 
            text="Include in Last Gift Values", 
            variable=self.include_in_last_gift)
        self.include_in_last_gift_checkbox.pack(side=tk.LEFT)
        TooltipLabel(last_gift_frame, text="ℹ", 
                    tooltip_text="Include these values in last gift calculations").pack(side=tk.LEFT)

        # Set default value for not found entries with tooltip
        self.use_default_value = tk.BooleanVar()
        default_frame = ttk.Frame(self.options_frame)
        default_frame.pack(fill=tk.X)
        self.default_value_checkbox = ttk.Checkbutton(default_frame, 
            text="Set Value for Not Found Entries",
            variable=self.use_default_value,
            command=self.toggle_default_value_entry)
        self.default_value_checkbox.pack(side=tk.LEFT)
        TooltipLabel(default_frame, text="ℹ", 
                    tooltip_text="Use a default value when no match is found").pack(side=tk.LEFT)

        # Set empty value for empty entries with tooltip
        self.use_empty_value = tk.BooleanVar()
        empty_frame = ttk.Frame(self.options_frame)
        empty_frame.pack(fill=tk.X)
        self.empty_value_checkbox = ttk.Checkbutton(empty_frame, 
            text="Set Value for Empty Entries",
            variable=self.use_empty_value,
            command=self.toggle_empty_value_entry)
        self.empty_value_checkbox.pack(side=tk.LEFT)
        TooltipLabel(empty_frame, text="ℹ", 
                    tooltip_text="Use a specific value when a lookup entry exists but has an empty value").pack(side=tk.LEFT)

        # Default value entry
        self.default_value_frame = ttk.Frame(details_frame)
        self.default_value_frame.grid(row=6, column=0, columnspan=2, sticky="w")
        ttk.Label(self.default_value_frame, text="Default Value:").pack(side=tk.LEFT)
        self.default_value_entry = ttk.Entry(self.default_value_frame)
        self.default_value_entry.pack(side=tk.LEFT, padx=5)
        self.default_value_frame.grid_remove()  # Initially hidden

        # Empty value entry
        self.empty_value_frame = ttk.Frame(details_frame)
        self.empty_value_frame.grid(row=7, column=0, columnspan=2, sticky="w")
        ttk.Label(self.empty_value_frame, text="Empty Value:").pack(side=tk.LEFT)
        self.empty_value_entry = ttk.Entry(self.empty_value_frame)
        self.empty_value_entry.pack(side=tk.LEFT, padx=5)
        self.empty_value_frame.grid_remove()  # Initially hidden

        # Dictionary values frame
        self.values_frame = ttk.Frame(details_frame)
        self.values_frame.grid(row=8, column=0, columnspan=3, pady=10)
        self.create_values_widgets()

        # Buttons
        button_frame = ttk.Frame(details_frame)
        button_frame.grid(row=9, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="New", command=self.new_dictionary).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save", command=self.save_dictionary).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Delete", command=self.delete_dictionary).pack(side=tk.LEFT, padx=5)

        ttk.Button(self, text="Done", command=self.done).pack(side=tk.BOTTOM, pady=10)
        self.update_dictionary_list()

        # Initially hide the values frame
        self.values_frame.grid_remove()

    def toggle_multiple_values(self):
        """Toggle visibility of options based on multiple values setting."""
        if self.use_multiple_values.get():
            self.single_value_frame.grid_remove()
            self.options_frame.grid_remove()
            self.values_frame.grid_remove()
            # Reset and disable other checkboxes
            self.use_post_merger.set(False)
            self.use_zip_validation.set(False)
            self.include_in_last_gift.set(False)
            self.use_default_value.set(False)
        else:
            self.single_value_frame.grid()
            self.options_frame.grid()

    def toggle_default_value_entry(self):
        if self.use_default_value.get():
            self.default_value_frame.grid()
        else:
            self.default_value_frame.grid_remove()

    def toggle_empty_value_entry(self):
        if self.use_empty_value.get():
            self.empty_value_frame.grid()
        else:
            self.empty_value_frame.grid_remove()

    def create_values_widgets(self):
        ttk.Label(self.values_frame, text="Key").grid(row=0, column=0)
        ttk.Label(self.values_frame, text="Merger Key").grid(row=0, column=1)
        ttk.Label(self.values_frame, text="Value").grid(row=0, column=2)
        ttk.Label(self.values_frame, text="Clean Name").grid(row=0, column=3)
        ttk.Label(self.values_frame, text="Clean Merger Name").grid(row=0, column=4)

        self.values_tree = ttk.Treeview(self.values_frame, columns=("key", "merger_key", "value", "clean_name", "clean_merger_name"), show="headings")
        self.values_tree.heading("key", text="Key")
        self.values_tree.heading("merger_key", text="Merger Key")
        self.values_tree.heading("value", text="Value")
        self.values_tree.heading("clean_name", text="Clean Name")
        self.values_tree.heading("clean_merger_name", text="Clean Merger Name")
        self.values_tree.grid(row=1, column=0, columnspan=6, sticky="nsew")

        scrollbar = ttk.Scrollbar(self.values_frame, orient="vertical", command=self.values_tree.yview)
        scrollbar.grid(row=1, column=6, sticky="ns")
        self.values_tree.configure(yscrollcommand=scrollbar.set)

        ttk.Button(self.values_frame, text="Add/Edit Value", command=self.add_edit_value).grid(row=2, column=0, columnspan=3, pady=5)
        ttk.Button(self.values_frame, text="Remove Value", command=self.remove_value).grid(row=2, column=3, columnspan=3, pady=5)

    def toggle_values_frame(self):
        if self.use_post_merger.get():
            self.values_frame.grid()
        else:
            self.values_frame.grid_remove()

    def add_edit_value(self):
        selected = self.values_tree.selection()
        if selected:
            # Edit existing value
            item = self.values_tree.item(selected[0])
            key, merger_key, value, clean_name, clean_merger_name = item['values']
        else:
            # Add new value
            key, merger_key, value, clean_name, clean_merger_name = '', '', '', '', ''

        dialog = tk.Toplevel(self)
        dialog.title("Add/Edit Value")

        ttk.Label(dialog, text="Key:").grid(row=0, column=0, sticky="e")
        key_entry = ttk.Entry(dialog)
        key_entry.grid(row=0, column=1)
        key_entry.insert(0, key)

        ttk.Label(dialog, text="Merger Key:").grid(row=1, column=0, sticky="e")
        merger_key_entry = ttk.Entry(dialog)
        merger_key_entry.grid(row=1, column=1)
        merger_key_entry.insert(0, merger_key)

        ttk.Label(dialog, text="Value:").grid(row=2, column=0, sticky="e")
        value_entry = ttk.Entry(dialog)
        value_entry.grid(row=2, column=1)
        value_entry.insert(0, value)

        ttk.Label(dialog, text="Clean Name:").grid(row=3, column=0, sticky="e")
        clean_name_entry = ttk.Entry(dialog)
        clean_name_entry.grid(row=3, column=1)
        clean_name_entry.insert(0, clean_name if clean_name else key)

        ttk.Label(dialog, text="Clean Merger Name:").grid(row=4, column=0, sticky="e")
        clean_merger_name_entry = ttk.Entry(dialog)
        clean_merger_name_entry.grid(row=4, column=1)
        clean_merger_name_entry.insert(0, clean_merger_name if clean_merger_name else merger_key)

        def save_value():
            new_key = key_entry.get()
            new_merger_key = merger_key_entry.get()
            new_value = value_entry.get()
            new_clean_name = clean_name_entry.get() or new_key
            new_clean_merger_name = clean_merger_name_entry.get() or new_merger_key

            if selected:
                self.values_tree.item(selected[0], values=(new_key, new_merger_key, new_value, new_clean_name, new_clean_merger_name))
            else:
                self.values_tree.insert('', 'end', values=(new_key, new_merger_key, new_value, new_clean_name, new_clean_merger_name))
            dialog.destroy()

        ttk.Button(dialog, text="Save", command=save_value).grid(row=5, column=0, columnspan=2, pady=10)

    def remove_value(self):
        selected = self.values_tree.selection()
        if selected:
            self.values_tree.delete(selected[0])

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            if not self.use_post_merger.get():
                self.load_dictionary_values(file_path)

    def load_dictionary_values(self, file_path):
        try:
            if self.use_multiple_values.get():
                df = pd.read_excel(file_path)
                columns = df.columns.tolist()
                key_column = columns[0]
                value_columns = columns[1:]
                
                for _, row in df.iterrows():
                    values = {}
                    for col in value_columns:
                        values[col] = row[col]
                    self.values_tree.insert('', 'end', values=(row[key_column], '', str(values), row[key_column], ''))
            else:
                df = pd.read_excel(file_path, header=None, names=['key', 'value'])
                for _, row in df.iterrows():
                    self.values_tree.insert('', 'end', values=(row['key'], '', row['value'], row['key'], ''))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load dictionary values: {str(e)}")

    def update_dictionary_list(self):
        self.dict_listbox.delete(0, tk.END)
        for dictionary in self.dictionaries:
            self.dict_listbox.insert(tk.END, dictionary['name'])

    def on_dictionary_select(self, event):
        selected = self.dict_listbox.curselection()
        if selected:
            dictionary = self.dictionaries[selected[0]]
            self.load_dictionary_details(dictionary)

    def load_dictionary_details(self, dictionary):
        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, dictionary['name'])
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, dictionary['path'])
        
        self.use_multiple_values.set(dictionary.get('use_multiple_values', False))
        self.toggle_multiple_values()
        
        self.lookup_column_combo.set(dictionary['lookup_column'])
        
        if not dictionary.get('use_multiple_values', False):
            self.output_column_entry.delete(0, tk.END)
            self.output_column_entry.insert(0, dictionary['output_column'])
        
        self.use_post_merger.set(dictionary.get('use_post_merger', False))
        self.use_zip_validation.set(dictionary.get('use_zip_validation', False))
        self.include_in_last_gift.set(dictionary.get('include_in_last_gift', False))
        self.use_default_value.set(dictionary.get('use_default_value', False))
        self.use_empty_value.set(dictionary.get('use_empty_value', False))
        
        self.default_value_entry.delete(0, tk.END)
        if dictionary.get('use_default_value', False):
            self.default_value_entry.insert(0, dictionary.get('default_value', ''))
            self.default_value_frame.grid()
        else:
            self.default_value_frame.grid_remove()
            
        self.empty_value_entry.delete(0, tk.END)
        if dictionary.get('use_empty_value', False):
            self.empty_value_entry.insert(0, dictionary.get('empty_value', 'EMPTY'))
            self.empty_value_frame.grid()
        else:
            self.empty_value_frame.grid_remove()
        
        self.toggle_values_frame()
        
        # Load dictionary values
        self.values_tree.delete(*self.values_tree.get_children())
        if 'values' in dictionary:
            for value in dictionary['values']:
                self.values_tree.insert('', 'end', values=(
                    value['key'],
                    value.get('merger_key', ''),
                    value['value'],
                    value.get('clean_name', value['key']),
                    value.get('clean_merger_name', value.get('merger_key', ''))
                ))

    def new_dictionary(self):
        self.name_entry.delete(0, tk.END)
        self.file_entry.delete(0, tk.END)
        self.lookup_column_combo.set('')
        self.output_column_entry.delete(0, tk.END)
        self.use_multiple_values.set(False)
        self.use_post_merger.set(False)
        self.use_zip_validation.set(False)
        self.include_in_last_gift.set(False)
        self.use_default_value.set(False)
        self.use_empty_value.set(False)
        self.default_value_entry.delete(0, tk.END)
        self.empty_value_entry.delete(0, tk.END)
        self.default_value_frame.grid_remove()
        self.empty_value_frame.grid_remove()
        self.toggle_values_frame()
        self.toggle_multiple_values()
        self.values_tree.delete(*self.values_tree.get_children())

    def save_dictionary(self):
        name = self.name_entry.get()
        if not name:
            messagebox.showerror("Error", "Dictionary name is required.")
            return

        values = []
        if self.use_post_merger.get():
            for item in self.values_tree.get_children():
                key, merger_key, value, clean_name, clean_merger_name = self.values_tree.item(item)['values']
                values.append({
                    'key': key,
                    'merger_key': merger_key,
                    'value': value,
                    'clean_name': clean_name,
                    'clean_merger_name': clean_merger_name
                })

        dictionary = {
            'name': name,
            'path': self.file_entry.get(),
            'lookup_column': self.lookup_column_combo.get(),
            'use_multiple_values': self.use_multiple_values.get(),
            'use_post_merger': self.use_post_merger.get(),
            'use_zip_validation': self.use_zip_validation.get(),
            'include_in_last_gift': self.include_in_last_gift.get(),
            'use_default_value': self.use_default_value.get(),
        }

        if not self.use_multiple_values.get():
            dictionary['output_column'] = self.output_column_entry.get()

        if self.use_default_value.get():
            dictionary['default_value'] = self.default_value_entry.get()
            
        if self.use_empty_value.get():
            dictionary['use_empty_value'] = True
            dictionary['empty_value'] = self.empty_value_entry.get()

        if values:
            dictionary['values'] = values

        # Check if the dictionary already exists
        existing_dict = next((d for d in self.dictionaries if d['name'] == name), None)
        if existing_dict:
            existing_dict.update(dictionary)
        else:
            self.dictionaries.append(dictionary)

        self.update_dictionary_list()

    def delete_dictionary(self):
        selected = self.dict_listbox.curselection()
        if not selected:
            messagebox.showerror("Error", "No dictionary selected.")
            return

        del self.dictionaries[selected[0]]
        self.update_dictionary_list()

    def done(self):
        self.result = self.dictionaries
        self.destroy()

class BaseToolFrame(ttk.Frame):
    """Base class for tool frames with common UI elements and functionality."""
    
    def __init__(self, master, title, window_size="800x550"):
        super().__init__(master)
        self.master = master
        self.master.title(title)
        self.master.geometry(window_size)
        
        # Pack the frame first
        self.pack(fill=tk.BOTH, expand=True)
        
        # Initialize queues for logging and progress
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        
        # Configure logging
        configure_logging(self.log_queue)
        
        self.create_common_widgets()
        self.create_tool_specific_widgets()
        
        # Start checking queues
        self._check_queues()

    def create_common_widgets(self):
        """Create widgets common to all tools."""
        # Create menu bar
        self.menu_bar = tk.Menu(self.master)
        self.master.config(menu=self.menu_bar)

        # Options menu
        options_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Options", menu=options_menu)
        options_menu.add_command(label="Configure Lookup Dictionaries", command=self.configure_lookups)

        # Common buttons and progress elements - now using self instead of self.master
        self.import_button = ttk.Button(self, text="Select Input File", command=self.select_input_file)
        self.import_button.pack(pady=10)

        self.process_button = tk.Button(self, text="Process Data", command=self.start_processing, state=tk.DISABLED)
        self.process_button.pack(pady=10)

        self.progress_bar = ttk.Progressbar(self, length=300, mode='determinate')
        self.progress_bar.pack(pady=10)

        self.result_label = ttk.Label(self, text="")
        self.result_label.pack(pady=10)

        self.file_link = tk.Label(self, text="", fg="blue", cursor="hand2")
        self.file_link.pack(pady=5)
        self.file_link.bind("<Button-1>", self.open_file)

        self.log_text = tk.Text(self, height=10, width=50)
        self.log_text.pack(pady=10)

    def create_tool_specific_widgets(self):
        """Override this method to add tool-specific widgets."""
        pass

    def _check_queues(self):
        """Internal method to check queues and update UI."""
        try:
            # Process all pending log messages
            while True:
                try:
                    msg = self.log_queue.get_nowait()
                    self.log_text.insert(tk.END, msg + '\n')
                    self.log_text.see(tk.END)
                    self.log_text.update_idletasks()  # Update just the text widget
                except queue.Empty:
                    break

            # Process all pending progress updates
            while True:
                try:
                    progress = self.progress_queue.get_nowait()
                    self.progress_bar['value'] = progress
                    self.progress_bar.update_idletasks()  # Update just the progress bar
                except queue.Empty:
                    break

        except Exception as e:
            print(f"Error in queue processing: {str(e)}")

        finally:
            # Schedule next check
            self.master.after(100, self._check_queues)

    def select_input_file(self):
        """Handle input file selection."""
        self.input_file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")]
        )
        if self.input_file_path:
            self.log(f"Input file selected: {self.input_file_path}")
            self.process_button['state'] = tk.NORMAL

    def configure_lookups(self):
        """Configure lookup dictionaries."""
        config_dialog = LookupDictionaryConfigDialog(self.master, self.dict_manager.lookups)
        self.master.wait_window(config_dialog)
        if hasattr(config_dialog, 'result'):
            self.dict_manager.lookups = config_dialog.result
            self.dict_manager.save_dictionaries()
            self.log(f"Configured {len(self.dict_manager.lookups)} lookup dictionaries.")

    def start_processing(self):
        """Start data processing."""
        raise NotImplementedError("Subclasses must implement start_processing")

    def open_file(self, event):
        """Open the output file location."""
        if hasattr(self, 'output_path') and self.output_path:
            filepath = os.path.normpath(self.output_path)
            subprocess.run(['explorer', '/select,', filepath])

    def log(self, message):
        """Add a message to the log queue."""
        self.log_queue.put(message)

    def reset_ui(self):
        """Reset UI elements to their default state."""
        self.process_button.config(bg='SystemButtonFace', fg='black', text="Process Data")
        self.progress_bar['value'] = 0
        self.result_label.config(text="")
        self.file_link.config(text="")
        self.log_text.delete('1.0', tk.END)
