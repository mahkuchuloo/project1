import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
from datetime import datetime, date
import time
import subprocess
from utils import update_progress, configure_logging, check_queues
import queue

class LookupDictionaryConfigDialog(tk.Toplevel):
    def __init__(self, parent, columns, existing_dictionaries):
        super().__init__(parent)
        self.title("Lookup Dictionary Configuration")
        self.columns = columns
        self.dictionaries = existing_dictionaries.copy()
        self.create_widgets()

    def create_widgets(self):
        # Dictionary list
        self.dict_listbox = tk.Listbox(self, width=30)
        self.dict_listbox.pack(side=tk.LEFT, fill=tk.Y)
        self.dict_listbox.bind('<<ListboxSelect>>', self.on_dictionary_select)

        # Dictionary details frame
        details_frame = ttk.Frame(self)
        details_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(details_frame, text="Dictionary Name:").grid(row=0, column=0, sticky="w")
        self.name_entry = ttk.Entry(details_frame)
        self.name_entry.grid(row=0, column=1)

        ttk.Label(details_frame, text="Dictionary File:").grid(row=1, column=0, sticky="w")
        self.file_entry = ttk.Entry(details_frame)
        self.file_entry.grid(row=1, column=1)
        ttk.Button(details_frame, text="Browse", command=self.browse_file).grid(row=1, column=2)

        ttk.Label(details_frame, text="Lookup Column:").grid(row=2, column=0, sticky="w")
        self.lookup_column_combo = ttk.Combobox(details_frame, values=self.columns)
        self.lookup_column_combo.grid(row=2, column=1)

        ttk.Label(details_frame, text="Output Column:").grid(row=3, column=0, sticky="w")
        self.output_column_entry = ttk.Entry(details_frame)
        self.output_column_entry.grid(row=3, column=1)

        # Buttons
        button_frame = ttk.Frame(details_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="New", command=self.new_dictionary).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save", command=self.save_dictionary).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Delete", command=self.delete_dictionary).pack(side=tk.LEFT, padx=5)

        ttk.Button(self, text="Done", command=self.done).pack(side=tk.BOTTOM, pady=10)
        self.update_dictionary_list()

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

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
        self.lookup_column_combo.set(dictionary['lookup_column'])
        self.output_column_entry.delete(0, tk.END)
        self.output_column_entry.insert(0, dictionary['output_column'])

    def new_dictionary(self):
        self.name_entry.delete(0, tk.END)
        self.file_entry.delete(0, tk.END)
        self.lookup_column_combo.set('')
        self.output_column_entry.delete(0, tk.END)

    def save_dictionary(self):
        name = self.name_entry.get()
        if not name:
            messagebox.showerror("Error", "Dictionary name is required.")
            return

        dictionary = {
            'name': name,
            'path': self.file_entry.get(),
            'lookup_column': self.lookup_column_combo.get(),
            'output_column': self.output_column_entry.get(),
        }

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

class RFMAnalyzerHelper:
    def __init__(self, master):
        self.master = master
        self.master.title("RFM Analyzer Helper")
        self.master.geometry("800x550")  # Increased height to accommodate new UI elements
        self.final_data = None
        self.output_paths = {}
        self.lookups = []
        self.final_file_path = None
        self.load_dictionaries()
        self.create_widgets()
        
        # Initialize queues for logging and progress
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        
        # Configure logging
        configure_logging(self.log_queue)
        
        # Start checking queues
        self.master.after(100, self.check_queues)

    def create_widgets(self):
        self.import_button = ttk.Button(self.master, text="Select Final File", command=self.select_final_file)
        self.import_button.pack(pady=10)

        self.lookup_button = ttk.Button(self.master, text="Configure Lookup Dictionaries", command=self.configure_lookups, state=tk.DISABLED)
        self.lookup_button.pack(pady=10)

        self.export_type_var = tk.StringVar()
        self.export_type_var.set("Export B")
        export_types = ["Export B", "Export F", "Output 1", "All"]
        self.export_type_menu = ttk.OptionMenu(self.master, self.export_type_var, "Export B", *export_types)
        self.export_type_menu.pack(pady=10)

        self.process_button = ttk.Button(self.master, text="Process Data", command=self.process_data, state=tk.DISABLED)
        self.process_button.pack(pady=10)

        self.progress_bar = ttk.Progressbar(self.master, length=300, mode='determinate')
        self.progress_bar.pack(pady=10)

        self.result_label = ttk.Label(self.master, text="")
        self.result_label.pack(pady=10)

        # File link label
        self.file_link = tk.Label(self.master, text="", fg="blue", cursor="hand2")
        self.file_link.pack(pady=5)
        self.file_link.bind("<Button-1>", self.open_file)

        self.log_text = tk.Text(self.master, height=10, width=50)
        self.log_text.pack(pady=10)

    def check_queues(self):
        check_queues(self.log_queue, self.progress_queue, self.log_text, self.progress_bar, self.master)

    def select_final_file(self):
        self.final_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
        if self.final_file_path:
            self.log(f"Final file selected: {self.final_file_path}")
            self.process_button['state'] = tk.NORMAL
            self.lookup_button['state'] = tk.NORMAL

    def import_final_file(self):
        if not self.final_file_path:
            self.log("No file selected. Please select a file first.")
            return

        self.log("Importing file. Please wait...")
        update_progress(self.progress_queue, 0)

        # Read the Excel file
        self.final_data = pd.read_excel(self.final_file_path)
        update_progress(self.progress_queue, 25)
        self.log("File imported. Applying lookup dictionaries...")

        # Apply lookup dictionaries to final_data
        self.apply_lookup_dictionaries_to_final_data()

        update_progress(self.progress_queue, 50)
        self.log("Final file imported and processed successfully.")

    def apply_lookup_dictionaries_to_final_data(self):
        total_lookups = len(self.lookups)
        for i, lookup in enumerate(self.lookups, 1):
            self.log(f"Applying lookup dictionary: {lookup['name']}")
            try:
                lookup_df = pd.read_excel(lookup['path'], header=None, names=['key', 'value'])
                lookup_dict = dict(zip(lookup_df['key'], lookup_df['value']))
                
                # Merge the lookup values with the final_data
                self.final_data[lookup['output_column']] = self.final_data[lookup['lookup_column']].map(lookup_dict)
                
                # Update progress
                progress = 25 + (25 * i // total_lookups)
                update_progress(self.progress_queue, progress)
            except Exception as e:
                self.log(f"Error applying lookup dictionary {lookup['name']}: {str(e)}")

    def configure_lookups(self):
        if self.final_data is None:
            messagebox.showerror("Error", "Please import the final file first.")
            return
        
        config_dialog = LookupDictionaryConfigDialog(self.master, list(self.final_data.columns), self.lookups)
        self.master.wait_window(config_dialog)
        if hasattr(config_dialog, 'result'):
            self.lookups = config_dialog.result
            self.log(f"Configured {len(self.lookups)} lookup dictionaries.")
            self.save_dictionaries()
            # Re-apply lookup dictionaries to final_data after configuration
            self.apply_lookup_dictionaries_to_final_data()

    def ask_output_paths(self):
        export_type = self.export_type_var.get()
        if export_type == "All" or export_type == "Export B":
            self.output_paths['Export B'] = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"Export B {datetime.now().strftime('%d%m%Y - %H%M%S')}.xlsx"
            )
        if export_type == "All" or export_type == "Export F":
            self.output_paths['Export F'] = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"Export F {datetime.now().strftime('%d%m%Y - %H%M%S')}.xlsx"
            )
        if export_type == "All" or export_type == "Output 1":
            self.output_paths['Output 1'] = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"Output 1 {datetime.now().strftime('%d%m%Y - %H%M%S')}.xlsx"
            )

    def process_data(self):
        if not self.final_file_path:
            self.log("Please select the final file first.")
            return

        start_time = time.time()

        self.ask_output_paths()
        self.import_final_file()  # Import the file when processing starts

        export_type = self.export_type_var.get()
        update_progress(self.progress_queue, 0)
        
        if export_type == "All":
            total_steps = 3
            self.create_export_b()
            update_progress(self.progress_queue, (1 / total_steps) * 100)
            self.create_export_f()
            update_progress(self.progress_queue, (2 / total_steps) * 100)
            self.create_output_1()
            update_progress(self.progress_queue, 100)
        elif export_type == "Export B":
            self.create_export_b()
            update_progress(self.progress_queue, 100)
        elif export_type == "Export F":
            self.create_export_f()
            update_progress(self.progress_queue, 100)
        elif export_type == "Output 1":
            self.create_output_1()
            update_progress(self.progress_queue, 100)
        else:
            self.log("Invalid export type selected.")

        end_time = time.time()
        processing_time = end_time - start_time

        self.log("All processing completed.")
        self.update_ui_after_processing(processing_time)

    def create_export_b(self):
        self.log("Creating Export B...")
        
        export_b = pd.DataFrame()

        # Define the columns for Export B
        export_b_columns = [
            "Constituent ID", "Donation ID", "Giving Platform", "Received Date",
            "Donation Amount", "Campaign Name", "Appeal Name", "Parent Donation ID"
        ]

        # Map the columns from final_data to export_b
        column_mapping = {
            "Constituent ID": "Relationship ID",
            "Donation ID": "Transaction ID",
            "Giving Platform": "Giving Platform",
            "Received Date": "Date Clean",
            "Donation Amount": "Amount",
            "Campaign Name": "Recipient",
            "Appeal Name": "Transaction ID",  # Using Transaction ID as Appeal Name
            "Parent Donation ID": "Recurring ID"
        }

        total_columns = len(column_mapping)
        for i, (export_col, final_col) in enumerate(column_mapping.items(), 1):
            export_b[export_col] = self.final_data[final_col]
            progress = 50 + (25 * i // total_columns)
            update_progress(self.progress_queue, progress)

        self.log("Export B data prepared")

        if self.output_paths.get('Export B'):
            export_b.to_excel(self.output_paths['Export B'], index=False)
            self.log(f"Export B saved to: {self.output_paths['Export B']}")
        else:
            self.log("Export B save cancelled.")

        self.log("Export B completed")

    def create_export_f(self):
        self.log("Creating Export F...")
        
        # Group by Recurring ID and calculate required values
        grouped = self.final_data.groupby('Recurring ID')
        
        export_f = pd.DataFrame({
            'Constituent Number': grouped['Relationship ID'].first(),
            'Date Created':   pd.to_datetime(grouped['Initial Recurring Contribution Date'].first()).dt.date,
            'Parent Donation ID': grouped.apply(lambda x: x.name),  # Use the group name (Recurring ID) as Parent Donation ID
            'Sum Donation Amount by Parent Donation ID': grouped['Amount'].sum(),
            'First Gift Date': pd.to_datetime(grouped['Date Clean'].min()).dt.date,
            'Last Gift Date': pd.to_datetime(grouped['Date Clean'].max()).dt.date,
            'First Recurring Amount': grouped.apply(lambda x: x.loc[x['Date Clean'] == x['Date Clean'].min(), 'Amount'].iloc[0])
        })

        # Reset index to make 'Recurring ID' a regular column
        export_f = export_f.reset_index()

        self.log("Export F data prepared")
        update_progress(self.progress_queue, 75)

        if self.output_paths.get('Export F'):
            export_f.to_excel(self.output_paths['Export F'], index=False)
            self.log(f"Export F saved to: {self.output_paths['Export F']}")
        else:
            self.log("Export F save cancelled.")

        self.log("Export F completed")
        update_progress(self.progress_queue, 100)

    def create_output_1(self):
        self.log("Creating Output 1...")

        # Get unique Relationship IDs
        unique_ids = self.final_data['Relationship ID'].unique()

        # Create Output 1 dataframe with unique Constituent IDs (which are Relationship IDs in this case)
        output_1 = pd.DataFrame({
            'Constituent ID': unique_ids
        })

        # Remove duplicates based on 'Relationship ID' to ensure unique values
        deduped_data = self.final_data.drop_duplicates(subset='Relationship ID', keep='first')

        # Define the order of columns
        column_order = [
            "Constituent ID", "Contact Channel Status", "Current Employer Name",
            "Most Recent DS Score in Database", "Most Recent DS Wealth Based Capacity in Database",
            "Current Portfolio Assignment in Database", "Age", "Age Range", "Race", "Gender"
        ]

        # Add lookup dictionary columns between Gender and MSA
        for lookup in self.lookups:
            column_name = lookup['output_column']
            if column_name not in column_order:
                column_order.append(column_name)

        # Add remaining columns
        column_order.extend(["Longitude", "Latitude"])

        # Copy lookup dictionary columns from final_data to output_1
        total_columns = len(column_order)
        for i, column_name in enumerate(column_order, 1):
            if column_name == 'Constituent ID':
                continue  # We already have this column
            
            if column_name in deduped_data.columns:
                # Map the corresponding values from deduped_data to output_1 using 'Relationship ID'
                output_1 = output_1.merge(
                    deduped_data[['Relationship ID', column_name]],
                    how='left',  # Left join to retain all 'Constituent ID' values
                    left_on='Constituent ID',  # Map based on 'Constituent ID'
                    right_on='Relationship ID'  # From deduped_data 'Relationship ID'
                ).drop(columns=['Relationship ID'])  # Drop 'Relationship ID' after the merge
            else:
                self.log(f"Column {column_name} not found in final_data.")
                output_1[column_name] = None
            
            progress = 50 + (25 * i // total_columns)
            update_progress(self.progress_queue, progress)

        # Reorder columns
        output_1 = output_1[column_order]

        self.log("Output 1 data prepared")
        update_progress(self.progress_queue, 75)

        # Save to output path if provided
        if self.output_paths.get('Output 1'):
            output_1.to_excel(self.output_paths['Output 1'], index=False)
            self.log(f"Output 1 saved to: {self.output_paths['Output 1']}")
        else:
            self.log("Output 1 save cancelled.")

        self.log("Output 1 completed")
        update_progress(self.progress_queue, 100)

    def load_dictionaries(self):
        try:
            with open('rfm_lookup_dictionaries.json', 'r') as f:
                self.lookups = json.load(f)
            print(f"Loaded {len(self.lookups)} lookup dictionaries.")
        except FileNotFoundError:
            print("No existing lookup dictionaries found. Starting with an empty list.")
            self.lookups = []
        except json.JSONDecodeError:
            print("Error decoding rfm_lookup_dictionaries.json. Starting with an empty list.")
            self.lookups = []

    def save_dictionaries(self):
        try:
            with open('rfm_lookup_dictionaries.json', 'w') as f:
                json.dump(self.lookups, f)
            self.log("Lookup dictionaries saved successfully.")
        except Exception as e:
            self.log(f"Error saving lookup dictionaries: {str(e)}")

    def log(self, message):
        self.log_queue.put(message)

    def update_ui_after_processing(self, processing_time):
        update_progress(self.progress_queue, 100)
        self.process_button.config(state=tk.NORMAL)
        self.process_button['text'] = "Processing Complete"
        
        result_text = f"Files generated in {processing_time:.2f} seconds.\n"
        for export_type, file_path in self.output_paths.items():
            if file_path:
                result_text += f"{export_type} saved at: {file_path}\n"
        
        self.result_label.config(text=result_text)
        
        if self.output_paths:
            self.file_link.config(text="Open Output Directory")
        else:
            self.file_link.config(text="")

    def open_file(self, event):
        if self.output_paths:
            # Get the directory of the first output file
            output_dir = os.path.dirname(next(iter(self.output_paths.values())))
            # Normalize the file path for Windows
            filepath = os.path.normpath(output_dir)
            # Open File Explorer and select the directory
            subprocess.run(['explorer', filepath])

if __name__ == "__main__":
    root = tk.Tk()
    app = RFMAnalyzerHelper(root)
    root.mainloop()
