import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from data_platform import Platform
from column_mapping_frame import ColumnMappingFrame

class PlatformConfigDialog(tk.Toplevel):
    source_columns = [
        'Transaction ID', 'Secondary ID', 'Recipient', 'Contribution Form URL',
        'Display Name', 'Donor First Name', 'Donor Last Name',
        'Donor Address Line 1',
        'Donor City', 'Donor State', 'Donor ZIP', 'Donor Country',
        'Donor Occupation', 'Donor Employer', 'Donor Email', 'Donor Phone',
        'Recurring ID', 'Initial Recurring Contribution Date', 'Is Recurring',
        'Match?'
    ]

    def __init__(self, parent, platforms, save_callback):
        super().__init__(parent)
        self.title("Platform Configuration")
        self.platforms = platforms
        self.save_callback = save_callback
        self.column_mapping_frame = None
        self.create_widgets()

    def create_widgets(self):
        # Platform list
        self.platform_listbox = tk.Listbox(self, width=30)
        self.platform_listbox.pack(side=tk.LEFT, fill=tk.Y)
        self.update_platform_list()
        self.platform_listbox.bind('<<ListboxSelect>>', self.on_platform_select)

        # Platform details frame
        details_frame = ttk.Frame(self)
        details_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Sample file frame at the top
        sample_frame = ttk.Frame(details_frame)
        sample_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        self.upload_button = ttk.Button(sample_frame, text="Upload Sample File", command=self.upload_sample_file)
        self.upload_button.pack(side=tk.LEFT, padx=5)

        ttk.Label(details_frame, text="Platform Name:").grid(row=1, column=0, sticky="w")
        self.name_entry = ttk.Entry(details_frame)
        self.name_entry.grid(row=1, column=1)

        ttk.Label(details_frame, text="File Pattern:").grid(row=2, column=0, sticky="w")
        self.file_pattern_entry = ttk.Entry(details_frame)
        self.file_pattern_entry.grid(row=2, column=1)

        ttk.Label(details_frame, text="Date Field:").grid(row=3, column=0, sticky="w")
        self.date_field_combo = ttk.Combobox(details_frame, values=[])
        self.date_field_combo.grid(row=3, column=1)

        ttk.Label(details_frame, text="Date Fallback Field (Optional):").grid(row=4, column=0, sticky="w")
        self.date_fallback_combo = ttk.Combobox(details_frame, values=[])
        self.date_fallback_combo.grid(row=4, column=1)

        ttk.Label(details_frame, text="Amount Field:").grid(row=5, column=0, sticky="w")
        self.amount_field_combo = ttk.Combobox(details_frame, values=[])
        self.amount_field_combo.grid(row=5, column=1)

        ttk.Label(details_frame, text="ID Field (Transaction Id):").grid(row=6, column=0, sticky="w")
        self.id_field_combo = ttk.Combobox(details_frame, values=[])
        self.id_field_combo.grid(row=6, column=1)

        ttk.Label(details_frame, text="Relationship ID Key (Primary Lookup e.g: Email):").grid(row=7, column=0, sticky="w")
        self.relationship_id_key_combo = ttk.Combobox(details_frame, values=[])
        self.relationship_id_key_combo.grid(row=7, column=1)

        ttk.Label(details_frame, text="Secondary ID Field (Secondary Lookup e.g: Act Blue Id):").grid(row=8, column=0, sticky="w")
        self.secondary_id_field_combo = ttk.Combobox(details_frame, values=[])
        self.secondary_id_field_combo.grid(row=8, column=1)

        ttk.Label(details_frame, text="Has Display Name:").grid(row=9, column=0, sticky="w")
        self.has_display_name_var = tk.BooleanVar()
        self.has_display_name_check = ttk.Checkbutton(details_frame, variable=self.has_display_name_var, command=self.toggle_display_name_fields)
        self.has_display_name_check.grid(row=9, column=1)

        # Add Recurring True Value field
        ttk.Label(details_frame, text="Recurring True Value:").grid(row=10, column=0, sticky="w")
        self.recurring_value_combo = ttk.Combobox(details_frame, state='disabled')
        self.recurring_value_combo.grid(row=10, column=1, sticky="ew")
        self.recurring_value_combo.set("Select Is Recurring mapping first")

        ttk.Label(details_frame, text="Is Base Platform:").grid(row=11, column=0, sticky="w")
        self.is_base_platform_var = tk.BooleanVar()
        self.is_base_platform_check = ttk.Checkbutton(details_frame, variable=self.is_base_platform_var)
        self.is_base_platform_check.grid(row=11, column=1)

        # Buttons
        button_frame = ttk.Frame(details_frame)
        button_frame.grid(row=12, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="New", command=self.new_platform).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save", command=self.save_platform).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Delete", command=self.delete_platform).pack(side=tk.LEFT, padx=5)

        # Column mapping frame
        self.mapping_frame = ttk.LabelFrame(details_frame, text="Column Mapping")
        self.mapping_frame.grid(row=13, column=0, columnspan=2, sticky="nsew", pady=10)

    def update_recurring_values(self):
        """Update recurring value options based on Is Recurring mapping"""
        if not self.column_mapping_frame:
            return

        mappings = self.column_mapping_frame.get_mappings()
        if 'Is Recurring' not in mappings or mappings['Is Recurring']['target'] == 'N/A':
            self.recurring_value_combo.set("Select Is Recurring mapping first")
            self.recurring_value_combo.configure(state='disabled')
            return

        selected = self.platform_listbox.curselection()
        if not selected:
            return

        platform_name = self.platform_listbox.get(selected[0])
        platform = self.platforms[platform_name]
        
        # Use cached values from platform config
        values = platform.get_recurring_values()
        if not values:
            self.recurring_value_combo.set("No recurring values available")
            self.recurring_value_combo.configure(state='disabled')
            return

        self.recurring_value_combo.configure(values=values, state='readonly')
        if platform.recurring_true_value in values:
            self.recurring_value_combo.set(platform.recurring_true_value)
        else:
            self.recurring_value_combo.set(values[0])

    def update_platform_list(self):
        self.platform_listbox.delete(0, tk.END)
        for platform in self.platforms.values():
            self.platform_listbox.insert(tk.END, platform.name)

    def on_platform_select(self, event):
        selected = self.platform_listbox.curselection()
        if selected:
            platform_name = self.platform_listbox.get(selected[0])
            platform = self.platforms[platform_name]
            self.load_platform_details(platform)

    def toggle_display_name_fields(self):
        if not hasattr(self, 'column_mapping_frame') or not self.column_mapping_frame:
            return
            
        has_display_name = self.has_display_name_var.get()
        
        # Only handle Display Name field state
        if 'Display Name' in self.column_mapping_frame.entry_widgets:
            entry_widgets = self.column_mapping_frame.entry_widgets['Display Name']
            entry_widgets['target'].configure(state='normal' if has_display_name else 'disabled')
            entry_widgets['default'].configure(state='normal' if has_display_name else 'disabled')

    def load_platform_details(self, platform):
        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, platform.name)
        self.file_pattern_entry.delete(0, tk.END)
        self.file_pattern_entry.insert(0, platform.file_pattern)

        # Update upload button text based on whether sample columns exist
        self.upload_button.configure(text="Reupload Sample File" if platform.sample_columns else "Upload Sample File")

        # Use sample columns if available, otherwise use mapping targets
        column_names = platform.sample_columns if platform.sample_columns else [mapping['target'] for mapping in platform.column_mapping.values()]
        self.date_field_combo['values'] = column_names
        self.date_field_combo.set(platform.date_field)
        self.date_fallback_combo['values'] = [''] + column_names
        self.date_fallback_combo.set(platform.date_fallback_field if platform.date_fallback_field else '')
        self.amount_field_combo['values'] = column_names
        self.amount_field_combo.set(platform.amount_field)
        self.id_field_combo['values'] = column_names
        self.id_field_combo.set(platform.id_field)
        self.secondary_id_field_combo['values'] = column_names
        self.secondary_id_field_combo.set(platform.secondary_id_field)
        self.is_base_platform_var.set(platform.is_base_platform())
        self.relationship_id_key_combo['values'] = column_names
        self.relationship_id_key_combo.set(platform.relationship_id_key)
        self.has_display_name_var.set(getattr(platform, 'has_display_name', False))

        if self.column_mapping_frame:
            self.column_mapping_frame.destroy()
        
        self.column_mapping_frame = ColumnMappingFrame(self.mapping_frame, column_names, self.source_columns)
        self.column_mapping_frame.pack(fill=tk.BOTH, expand=True)
        self.column_mapping_frame.set_mappings(platform.column_mapping)

        # Update recurring value options
        self.update_recurring_values()

    def new_platform(self):
        # Clear all fields
        self.name_entry.delete(0, tk.END)
        self.file_pattern_entry.delete(0, tk.END)
        self.date_field_combo.set('')
        self.date_fallback_combo.set('')
        self.amount_field_combo.set('')
        self.id_field_combo.set('')
        self.secondary_id_field_combo.set('')
        self.is_base_platform_var.set(False)
        self.relationship_id_key_combo.set('')
        self.recurring_value_combo.set("Select Is Recurring mapping first")
        self.recurring_value_combo.configure(state='disabled')
        self.upload_button.configure(text="Upload Sample File")
        if self.column_mapping_frame:
            self.column_mapping_frame.destroy()
            self.column_mapping_frame = None

    def save_platform(self):
        name = self.name_entry.get()
        if not name:
            messagebox.showerror("Error", "Platform name is required.")
            return

        # Check if the platform already exists
        if name in self.platforms:
            platform = self.platforms[name]
            platform.file_pattern = self.file_pattern_entry.get()
            platform.date_field = self.date_field_combo.get()
            platform.date_fallback_field = self.date_fallback_combo.get() or None
            platform.amount_field = self.amount_field_combo.get()
            platform.id_field = self.id_field_combo.get()
            platform.secondary_id_field = self.secondary_id_field_combo.get()
            platform._is_base_platform = self.is_base_platform_var.get()
            platform.relationship_id_key = self.relationship_id_key_combo.get()
            platform.has_display_name = self.has_display_name_var.get()
            platform.recurring_true_value = self.recurring_value_combo.get() if self.recurring_value_combo['state'] != 'disabled' else None
        else:
            platform = Platform(
                name,
                self.file_pattern_entry.get(),
                self.date_field_combo.get(),
                self.amount_field_combo.get(),
                self.id_field_combo.get(),
                self.secondary_id_field_combo.get(),
                self.is_base_platform_var.get(),
                self.relationship_id_key_combo.get(),
                self.has_display_name_var.get(),
                self.date_fallback_combo.get() or None
            )
            platform.recurring_true_value = self.recurring_value_combo.get() if self.recurring_value_combo['state'] != 'disabled' else None

        if self.column_mapping_frame:
            platform.column_mapping = self.column_mapping_frame.get_mappings()

        self.platforms[name] = platform
        self.update_platform_list()
        self.save_callback()

    def delete_platform(self):
        selected = self.platform_listbox.curselection()
        if not selected:
            messagebox.showerror("Error", "No platform selected.")
            return

        name = self.platform_listbox.get(selected[0])
        del self.platforms[name]
        self.update_platform_list()
        self.save_callback()

    def upload_sample_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            if df is None or not isinstance(df, pd.DataFrame):
                raise ValueError("Failed to read file as DataFrame")
            if df.empty:
                raise ValueError("The uploaded file is empty")
                
            column_names = df.columns.tolist()
            if not column_names:
                raise ValueError("No columns found in file")
            
            # Store column names and file path in platform if one is selected
            selected = self.platform_listbox.curselection()
            if selected:
                platform_name = self.platform_listbox.get(selected[0])
                platform = self.platforms[platform_name]
                platform.sample_columns = column_names
                platform.sample_file_path = file_path
                # Only update recurring values if we have a valid mapping
                if platform.column_mapping.get('Is Recurring', {}).get('target', 'N/A') != 'N/A':
                    platform.update_recurring_values(df)
            
            # Update combo boxes with the new column names
            for combo in [self.date_field_combo, self.amount_field_combo, 
                         self.id_field_combo, self.secondary_id_field_combo, 
                         self.relationship_id_key_combo]:
                combo['values'] = column_names
            self.date_fallback_combo['values'] = [''] + column_names

            # Create new mapping frame while preserving existing mappings
            existing_mappings = None
            if self.column_mapping_frame:
                existing_mappings = self.column_mapping_frame.get_mappings()
                self.column_mapping_frame.destroy()
            
            self.column_mapping_frame = ColumnMappingFrame(self.mapping_frame, column_names, self.source_columns)
            self.column_mapping_frame.pack(fill=tk.BOTH, expand=True)
            
            # Restore existing mappings if available
            if existing_mappings:
                self.column_mapping_frame.set_mappings(existing_mappings)

            # Bind to mapping changes to update recurring values
            for widgets in self.column_mapping_frame.mappings.values():
                widgets["target"].bind('<<ComboboxSelected>>', lambda e: self.update_recurring_values())

            # Update button text to show it's a reupload
            self.upload_button.configure(text="Reupload Sample File")
            
            # Update recurring values after uploading sample file
            self.update_recurring_values()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {str(e)}")
