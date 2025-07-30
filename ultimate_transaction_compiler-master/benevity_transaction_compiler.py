import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import logging
import traceback
import queue
import threading
import time
import os
import subprocess
import csv
from datetime import datetime
from utils import configure_logging, check_queues, update_progress

class BenevityTransactionCompiler:
    def __init__(self, master):
        logging.info("Initializing BenevityTransactionCompiler")
        self.master = master
        self.master.title("Benevity Transaction Compiler")
        self.master.geometry("800x700")

        self.input_files = []
        self.output_file = ""

        # Define expected columns in order
        self.expected_columns = [
            'Company',
            'Project',
            'Donation Date',
            'Donor First Name',
            'Donor Last Name',
            'Email',
            'Address',
            'City',
            'State/Prov',
            'Postal Code',
            'Activity',
            'Comment',
            'Transaction ID',
            'Donation Frequency',
            'Currency',
            'Project Remote ID',
            'Source',
            'Reason',
            'Total Donation to be Acknowledged',
            'Match Amount',
            'Cause Support Fee',
            'Merchant Fee',
            'Fee Comment'
        ]

        # Set up logging to GUI
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        configure_logging(self.log_queue)

        self.create_widgets()

        # Start checking the queues
        self.master.after(100, check_queues, self.log_queue, self.progress_queue, self.log_text, self.progress, self.master)

    def create_widgets(self):
      
        # Main frame to hold all widgets
        main_frame = tk.Frame(self.master)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Add files button
        self.add_button = tk.Button(main_frame, text="Add Benevity Files", command=self.add_files)
        self.add_button.pack(pady=5)

        # Create a frame to contain the canvas and scrollbar
        container_frame = tk.Frame(main_frame)
        container_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Add a canvas
        canvas = tk.Canvas(container_frame, height=150)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add a scrollbar
        scrollbar = ttk.Scrollbar(container_frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Configure the canvas
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create a frame inside the canvas to hold the files
        self.files_frame = tk.Frame(canvas)
        self.files_frame_window = canvas.create_window((0, 0), window=self.files_frame, anchor='nw', tags="files_frame")

        # Update scroll region when the size of the files frame changes
        self.files_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(self.files_frame_window, width=e.width))

        # Process button
        self.process_button = tk.Button(main_frame, text="Process Files", command=self.start_processing)
        self.process_button.pack(pady=10)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, length=700, mode='determinate')
        self.progress.pack(pady=10)

        # Log text area
        self.log_text = tk.Text(main_frame, height=15, width=90)
        self.log_text.pack(pady=10)

        # Result label
        self.result_label = tk.Label(main_frame, text="", justify=tk.LEFT, wraplength=750)
        self.result_label.pack(pady=10)

        # File link label
        self.file_link = tk.Label(main_frame, text="", fg="blue", cursor="hand2")
        self.file_link.pack(pady=5)
        self.file_link.bind("<Button-1>", self.open_file)

    def add_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[
            ("Supported files", "*.xlsx;*.xls;*.csv"),
            ("Excel files", "*.xlsx;*.xls"),
            ("CSV files", "*.csv")
        ])
        for file_path in file_paths:
            self.add_file_to_list(file_path)
            logging.info(f"File added: {file_path}")

    def add_file_to_list(self, file_path):
        # Create file frame
        file_frame = tk.Frame(self.files_frame)
        file_frame.pack(fill=tk.X, pady=2)
        
        # Create and add tag label
        tag_label = tk.Label(file_frame, text="Benevity", bg="#1f77b4", fg="white", padx=5, pady=2)
        tag_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # Create and add file label
        file_label = tk.Label(file_frame, text=os.path.basename(file_path), anchor="w")
        file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Create and add remove button
        remove_button = tk.Button(file_frame, text="X", command=lambda: self.remove_file(file_frame, file_path))
        remove_button.pack(side=tk.RIGHT)
        
        # Add file to list
        self.input_files.append(file_path)

    def remove_file(self, file_frame, file_path):
        file_frame.destroy()
        self.input_files.remove(file_path)

    def start_processing(self):
        if not self.input_files:
            logging.error("No files selected")
            messagebox.showerror("Error", "Please select at least one Benevity file.")
            return

        # Reset UI elements
        self.reset_ui()

        # Ask for output file location
        self.output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"BenevityFinalFile {datetime.now().strftime('%d%m%Y - %H%M%S')}.xlsx"
        )

        if not self.output_file:
            logging.warning("Output file not selected")
            return

        # Start processing in a separate thread
        self.process_button.config(state=tk.DISABLED)
        threading.Thread(target=self.process_files, daemon=True).start()

    def reset_ui(self):
        self.process_button.config(bg='SystemButtonFace', fg='black', text="Process Files")
        self.progress['value'] = 0
        self.result_label.config(text="")
        self.file_link.config(text="")
        self.log_text.delete('1.0', tk.END)

    def convert_csv_file(self, file_path):
        """Convert problematic Benevity CSV file to a proper CSV format."""
        temp_file = file_path + '.temp'
        try:
            max_columns = 23  # We know Benevity files have a maximum of 23 columns in the data section
            
            with open(file_path, 'r', encoding='utf-8-sig') as infile, \
                 open(temp_file, 'w', newline='', encoding='utf-8') as outfile:
                csv_writer = csv.writer(outfile)
                lines = infile.readlines()
                
                for line in lines:
                    line = line.strip()
                    if not line or line.startswith('"#---'):
                        continue
                    
                    # Parse the line using csv.reader to handle quotes properly
                    row = next(csv.reader([line]))
                    
                    # Skip empty rows
                    if not any(cell.strip() for cell in row):
                        continue
                        
                    # Pad row with empty strings to match max columns
                    padded_row = row + [''] * (max_columns - len(row))
                    csv_writer.writerow(padded_row)
            
            os.replace(temp_file, file_path)
            return True
        except Exception as e:
            logging.error(f"Error converting CSV file {file_path}: {str(e)}")
            if os.path.exists(temp_file):
                os.remove(temp_file)
            return False

    def process_files(self):
        logging.info("Processing files")
        start_time = time.time()
        try:
            # Process each file
            all_data = []
            total_files = len(self.input_files)
            
            for idx, file_path in enumerate(self.input_files):
                logging.info(f"Processing file {idx + 1}/{total_files}: {file_path}")
                
                # Determine file type and read accordingly
                file_ext = os.path.splitext(file_path)[1].lower()
                if file_ext == '.csv':
                    # Convert CSV file first to ensure proper format
                    if not self.convert_csv_file(file_path):
                        raise ValueError(f"Failed to convert CSV file: {file_path}")
                
                # Read file based on type
                if file_ext == '.csv':
                    df = pd.read_csv(file_path, header=None, encoding='utf-8-sig')
                else:
                    df = pd.read_excel(file_path, header=None)
                
                # Extract metadata
                metadata = {}
                for i in range(9):
                    row = df.iloc[i]
                    if pd.notna(row[0]) and pd.notna(row[1]):
                        key = str(row[0]).strip()
                        value = str(row[1]).strip()
                        if key in ['Disbursement ID', 'Period Ending', 'Charity ID']:
                            metadata[key] = value

                # Find totals
                totals = {}
                for idx, row in df.iterrows():
                    if pd.notna(row[0]):
                        key = str(row[0]).strip()
                        if key in ['Total Donations (Gross)', 'Net Total Payment']:
                            totals[key] = row[1]

                # Find header row
                header_row = None
                for idx, row in df.iterrows():
                    if pd.notna(row[1]) and str(row[1]).strip() == 'Project':
                        header_row = idx
                        break

                if header_row is None:
                    raise ValueError(f"Could not find header row in {file_path}")

                # Read data section with proper headers (using same file since it's already converted)
                if file_ext == '.csv':
                    data_df = pd.read_csv(file_path, skiprows=header_row+1, names=self.expected_columns, encoding='utf-8-sig')
                else:
                    data_df = pd.read_excel(file_path, skiprows=header_row+1, names=self.expected_columns)
                
                # Find where totals section starts
                totals_start = None
                for idx, row in data_df.iterrows():
                    if pd.notna(row['Company']) and str(row['Company']).strip().lower() == 'totals':
                        totals_start = idx
                        break
                
                # Remove totals section if found
                if totals_start is not None:
                    data_df = data_df.iloc[:totals_start]

                # Keep Donation Date in ISO format
                if 'Donation Date' in data_df.columns:
                    data_df['Donation Date'] = pd.to_datetime(data_df['Donation Date']).dt.strftime('%Y-%m-%dT%H:%M:%SZ')

                # Add metadata columns
                for key in ['Disbursement ID', 'Period Ending', 'Charity ID']:
                    data_df[key] = metadata.get(key, '')
                for key in ['Total Donations (Gross)', 'Net Total Payment']:
                    data_df[key] = totals.get(key, '')

                # Ensure private information format is consistent
                for col in ['Email', 'Address', 'State/Prov']:
                    mask = data_df[col].str.contains('Not shared', case=False, na=False)
                    data_df.loc[mask, col] = 'Not shared by donor'

                all_data.append(data_df)
                
                # Update progress
                progress = int((idx + 1) / total_files * 100)
                update_progress(self.progress_queue, progress)

            # Combine all data
            final_df = pd.concat(all_data, ignore_index=True)

            # Convert monetary columns to numeric, handling any non-numeric values
            final_df['Total Donation to be Acknowledged'] = pd.to_numeric(final_df['Total Donation to be Acknowledged'], errors='coerce')
            final_df['Match Amount'] = pd.to_numeric(final_df['Match Amount'], errors='coerce')

            # Create Amount column (sum of Total Donation to be Acknowledged and Match Amount)
            final_df['Amount'] = final_df['Total Donation to be Acknowledged'].fillna(0) + final_df['Match Amount'].fillna(0)

            # Create Relationship Id column
            def generate_relationship_id(row):
                if pd.notna(row['Email']) and row['Email'] != '' and row['Email'] != 'Not shared by donor':
                    return row['Email']
                else:
                    # Check if donor info is not shared (indicating anonymous)
                    if (row['Donor First Name'] == 'Not shared by donor' or 
                        row['Donor Last Name'] == 'Not shared by donor'):
                        return f"Anonymous_{row['Company']}_Benevity"
                    else:
                        return f"{row['Company']}_{row['Donor First Name']}_{row['Donor Last Name']}"

            final_df['Relationship ID'] = final_df.apply(generate_relationship_id, axis=1)

            # Reorder columns to put metadata first
            metadata_cols = [
                'Disbursement ID',
                'Period Ending',
                'Total Donations (Gross)',
                'Net Total Payment',
                'Charity ID'
            ]
            # Add remaining columns in expected order
            final_cols = metadata_cols + [col for col in self.expected_columns if col not in metadata_cols] + ['Amount', 'Relationship ID']
            final_df = final_df[final_cols]

            # Save to Excel
            final_df.to_excel(self.output_file, index=False)
            
            end_time = time.time()
            processing_time = end_time - start_time

            self.master.after(0, self.update_ui_after_processing, processing_time)

        except Exception as e:
            logging.error(f"Error occurred: {str(e)}")
            logging.error(traceback.format_exc())
            self.master.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {str(e)}\nCheck the log file for more details."))

    def update_ui_after_processing(self, processing_time):
        self.progress['value'] = 100
        self.process_button.config(state=tk.NORMAL, bg='green', fg='white', text="Processing Complete")
        
        result_text = f"Final file generated in {processing_time:.2f} seconds.\n"
        result_text += f"File saved at: {self.output_file}"
        
        self.result_label.config(text=result_text)
        self.file_link.config(text="Open File")

    def open_file(self, event):
        filepath = os.path.normpath(self.output_file)
        subprocess.run(['explorer', '/select,', filepath])

if __name__ == "__main__":
    root = tk.Tk()
    app = BenevityTransactionCompiler(root)
    root.mainloop()
