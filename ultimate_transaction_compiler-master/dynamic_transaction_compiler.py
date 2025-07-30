import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from collections import defaultdict
from datetime import datetime
import logging
import traceback
import queue
import threading
import time
import os
import subprocess
import json
from data_platform import Platform
from platform_config_dialog import PlatformConfigDialog
from utils import configure_logging, check_queues, update_progress, generate_fallback_id, add_unique_id

class DynamicTransactionCompiler:
    def __init__(self, master):
        logging.info("Initializing DynamicTransactionCompiler")
        self.master = master
        self.master.title("Dynamic Transaction Compiler")
        self.master.geometry("800x700")

        self.input_files = defaultdict(list)
        self.output_file = ""
        self.platforms = {}  # This will be populated with Platform instances

        # Set up logging to GUI
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        configure_logging(self.log_queue)

        self.create_widgets()

        # Start checking the queues
        self.master.after(100, check_queues, self.log_queue, self.progress_queue, self.log_text, self.progress, self.master)

        self.load_platforms()

    def create_widgets(self):
        # Create menu
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Configure Platforms", command=self.open_platform_config)

        # Main frame to hold all widgets
        main_frame = tk.Frame(self.master)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # File selection buttons frame
        self.file_buttons_frame = tk.Frame(main_frame)
        self.file_buttons_frame.pack(fill=tk.X, pady=5)

        # File list frames
        self.file_frames = {}

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

    def update_file_buttons(self):
        for widget in self.file_buttons_frame.winfo_children():
            widget.destroy()

        for platform in self.platforms:
            tk.Button(self.file_buttons_frame, text=f"Add {platform} Files", 
                      command=lambda p=platform: self.add_files(p)).pack(side=tk.LEFT, padx=5)

        for platform in self.platforms:
            if platform not in self.file_frames:
                self.file_frames[platform] = tk.Frame(self.master)
                self.file_frames[platform].pack(pady=5, fill=tk.X)

    def open_platform_config(self):
        PlatformConfigDialog(self.master, self.platforms, self.save_platforms)

    def load_platforms(self):
        try:
            with open('platform_config.json', 'r') as f:
                platforms_data = json.load(f)
            
            for platform_data in platforms_data:
                platform = Platform.from_dict(platform_data)
                self.platforms[platform.name] = platform
            
            logging.info("Platforms loaded successfully")
        except FileNotFoundError:
            logging.info("No platform configuration file found. Using default platforms.")
            everyaction = Platform('EveryAction', '*.xlsx', 'Date Received', 'Amount', 'VANID', 'ActBlue ID', True, 'Personal Email')
            everyaction.column_mapping = {
                'Transaction ID': {'target': 'Contribution ID', 'default': ''},
                'Secondary ID': {'target': 'ActBlue ID', 'default': ''},
                'Date Clean': {'target': 'Date Received', 'default': ''},
                'Recipient': {'target': 'Designation', 'default': ''},
                'Donor First Name': {'target': 'First Name', 'default': ''},
                'Donor Last Name': {'target': 'Last Name', 'default': ''},
                'Donor Address Line 1': {'target': 'Home Street Address', 'default': ''},
                'Donor City': {'target': 'Home City', 'default': ''},
                'Donor State': {'target': 'Home State/Province', 'default': ''},
                'Donor ZIP': {'target': 'Home Zip/Postal', 'default': ''},
                'Donor Country': {'target': 'Home Country', 'default': ''},
                'Donor Email': {'target': 'Personal Email', 'default': ''},
                'Donor Phone': {'target': 'Home Phone', 'default': ''},
                'Initial Recurring Contribution Date': {'target': 'Start Date', 'default': ''},
                'Is Recurring': {'target': 'Is Recurring Commitment', 'default': 'FALSE'},
                'Recurring ID': {'target': 'Recurring Commitment ID', 'default': ''},

            }

            actblue = Platform('ActBlue', '*.xlsx', 'Paid At', 'Amount', 'Order Number', 'Lineitem ID', False, 'Donor Email')
            actblue.column_mapping = {
                'Transaction ID': {'target': 'Lineitem ID', 'default': ''},
                'Secondary ID': {'target': 'Order Number', 'default': ''},
                'Date Clean': {'target': 'Paid At', 'default': ''},
                'Recipient': {'target': 'Recipient', 'default': ''},
                'Contribution Form URL': {'target': 'Contribution Form URL', 'default': ''},
                'Donor First Name': {'target': 'Donor First Name', 'default': ''},
                'Donor Last Name': {'target': 'Donor Last Name', 'default': ''},
                'Donor Address Line 1': {'target': 'Donor Address Line 1', 'default': ''},
                'Donor City': {'target': 'Donor City', 'default': ''},
                'Donor State': {'target': 'Donor State', 'default': ''},
                'Donor ZIP': {'target': 'Donor ZIP', 'default': ''},
                'Donor Country': {'target': 'Donor Country', 'default': ''},
                'Donor Occupation': {'target': 'Donor Occupation', 'default': ''},
                'Donor Employer': {'target': 'Donor Employer', 'default': ''},
                'Donor Email': {'target': 'Donor Email', 'default': ''},
                'Donor Phone': {'target': 'Donor Phone', 'default': ''},
                'Initial Recurring Contribution Date': {'target': 'Initial Recurring Contribution Date', 'default': ''},
                'Is Recurring': {'target': 'Is Recurring', 'default': 'FALSE'},
                'Recurring ID': {'target': 'N/A', 'default': ''},  # Don't map Recurring ID - we'll set it manually
            }

            self.platforms = {
                'EveryAction': everyaction,
                'ActBlue': actblue
            }
        except Exception as e:
            logging.error(f"Error loading platforms: {str(e)}")
            messagebox.showerror("Error", f"Failed to load platforms: {str(e)}")

        self.update_file_buttons()

    def save_platforms(self):
        try:
            platforms_data = [platform.to_dict() for platform in self.platforms.values()]
            with open('platform_config.json', 'w') as f:
                json.dump(platforms_data, f, indent=2)
            logging.info("Platforms saved successfully")
            self.update_file_buttons()
        except Exception as e:
            logging.error(f"Error saving platforms: {str(e)}")
            messagebox.showerror("Error", f"Failed to save platforms: {str(e)}")

    def add_files(self, platform):
        logging.info(f"Adding {platform} files")
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
        for file_path in file_paths:
            self.add_file_to_list(file_path, platform)
            logging.info(f"{platform} file added: {file_path}")

    def add_file_to_list(self, file_path, platform):
        # Define a list of 10 distinct colors with good contrast
        platform_colors = [
            "#1f77b4",  # Blue
            "#ff7f0e",  # Orange
            "#2ca02c",  # Green
            "#d62728",  # Red
            "#9467bd",  # Purple
            "#8c564b",  # Brown
            "#e377c2",  # Pink
            "#7f7f7f",  # Gray
            "#bcbd22",  # Yellow-green
            "#17becf"   # Cyan
        ]
        
        # Get consistent color index based on platform name
        color_index = hash(platform) % len(platform_colors)
        bg_color = platform_colors[color_index]
        
        # Append file to the list
        self.input_files[platform].append(file_path)
        
        # Create file frame
        file_frame = tk.Frame(self.file_frames[platform])
        file_frame.pack(fill=tk.X, pady=2)
        
        # Create and add tag label
        tag_label = tk.Label(file_frame, text=platform, bg=bg_color, fg="white", padx=5, pady=2)
        tag_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # Create and add file label
        file_label = tk.Label(file_frame, text=os.path.basename(file_path), anchor="w")
        file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Create and add remove button
        remove_button = tk.Button(file_frame, text="X", command=lambda: self.remove_file(file_frame, file_path, platform))
        remove_button.pack(side=tk.RIGHT)

    def remove_file(self, file_frame, file_path, platform):
        file_frame.destroy()
        self.input_files[platform].remove(file_path)

    def start_processing(self):
        if not all(self.input_files.values()):
            logging.error("Files not selected for all platforms")
            messagebox.showerror("Error", "Please select at least one file for each platform.")
            return

        # Reset UI elements
        self.reset_ui()

        # Ask for output file location
        self.output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"DynamicFinalFile {datetime.now().strftime('%d%m%Y - %H%M%S')}.xlsx"
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

    def process_files(self):
        logging.info("Processing files")
        start_time = time.time()
        try:
            # Read and combine input files
            platform_dfs = {}
            for platform, files in self.input_files.items():
                logging.info(f"Reading {platform} files")
                dfs = []
                for file in files:
                    df = self.read_excel_file(file)
                    dfs.append(df)
                platform_dfs[platform] = pd.concat(dfs, ignore_index=True)
                update_progress(self.progress_queue, 10 * len(platform_dfs))

            # Generate Transaction Values (Relationship IDs)
            logging.info("Generating transaction values")
            platform_dfs = self.generate_transaction_values(platform_dfs)
            update_progress(self.progress_queue, 60)

            # Process data
            logging.info("Creating final file")
            final_df = self.create_final_file(platform_dfs)
            update_progress(self.progress_queue, 80)

            # Save the result
            logging.info(f"Saving file to: {self.output_file}")
            self.save_excel_file(final_df, self.output_file)
            update_progress(self.progress_queue, 100)

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
        # Normalize the file path for Windows
        filepath = os.path.normpath(self.output_file)
        # Open File Explorer and select the file
        subprocess.run(['explorer', '/select,', filepath])

    def read_excel_file(self, file_path):
        try:
            df = pd.read_excel(file_path)
            logging.info(f"Successfully read file: {file_path}")
            logging.debug(f"DataFrame info:\n{df.info()}")
            return df
        except Exception as e:
            logging.error(f"Error reading file: {file_path}")
            logging.error(str(e))
            logging.error(traceback.format_exc())
            raise

    def save_excel_file(self, df, file_path):
        try:
            df.to_excel(file_path, index=False)
            logging.info(f"Successfully saved file to: {file_path}")
        except Exception as e:
            logging.error(f"Error saving file to: {file_path}")
            logging.error(str(e))
            logging.error(traceback.format_exc())
            raise

    def generate_transaction_values(self, platform_dfs):
        logging.info("Generating transaction values")
        try:
            # Create dictionaries for faster lookup
            dict_indices = defaultdict(list)
            dict_id = {}
            dict_primary_id = defaultdict(dict)
            dict_primary_unique = defaultdict(dict)
            dict_secondary_id = defaultdict(dict)
            dict_secondary_id_unique = defaultdict(dict)
            pre_rel_ids = {}

            # Set up progress tracking
            total_rows = sum(len(df) for df in platform_dfs.values())
            rows_processed = 0

            # PHASE 1: Build indices and unique transaction keys
            logging.info("Phase 1: Building indices and unique transaction keys")
            
            # Process base platform data
            base_platform = next(platform for platform in self.platforms.values() if platform.is_base_platform())
            base_df = platform_dfs[base_platform.get_platform_name()]
            
            for idx, row in base_df.iterrows():
                primary_key = base_platform.get_relationship_id_key(row)
                secondary_key = str(row[base_platform.get_secondary_id_field()]) if pd.notnull(row[base_platform.get_secondary_id_field()]) else ''
                amount = row[base_platform.get_amount_field()]
                date = pd.to_datetime(row[base_platform.get_date_field()]).date()
                id_value = row[base_platform.get_id_field()]

                dict_id[idx] = id_value
                unique_key = base_platform.get_unique_transaction_key(row)
                
                base_df.at[idx, 'Unique Transaction'] = unique_key

                if primary_key:
                    dict_indices[primary_key].append(idx)
                if secondary_key:
                    dict_indices[secondary_key].append(idx)

                rows_processed += 1
                if rows_processed % 100 == 0:
                    progress = 20 + int((rows_processed / total_rows) * 10)
                    update_progress(self.progress_queue, progress)

            # Process other platform data for indices
            for platform_name, df in platform_dfs.items():
                if platform_name == base_platform.get_platform_name():
                    continue  # Skip base platform as it's already processed
                
                platform = self.platforms[platform_name]
                for idx, row in df.iterrows():
                    primary_key = platform.get_relationship_id_key(row)
                    secondary_key = str(row[platform.get_id_field()])
                    amount = row[platform.get_amount_field()]
                    date = pd.to_datetime(row[platform.get_date_field()]).date()

                    if primary_key:
                        if primary_key in dict_indices:
                            if idx not in dict_primary_id[platform_name]:
                                all_ids_primary = dict_indices[primary_key]
                                unique_keys_primary = set()
                                for id_value in all_ids_primary:
                                    id_value = dict_id[id_value]
                                    unique_key = f"{id_value}{int(date.strftime('%Y%m%d'))}{amount}"
                                    unique_keys_primary.add(unique_key)
                                dict_primary_id[platform_name][idx] = all_ids_primary
                                dict_primary_unique[platform_name][idx] = unique_keys_primary
                                df.at[idx, f'Unique Transaction {platform_name}'] = '|'.join(unique_keys_primary)

                    if secondary_key:
                        if secondary_key in dict_indices:
                            if idx not in dict_secondary_id[platform_name]:
                                all_ids_secondary = dict_indices[secondary_key]
                                unique_keys_secondary = set()
                                for id_value in all_ids_secondary:
                                    id_value = dict_id[id_value]
                                    unique_key = f"{id_value}{int(date.strftime('%Y%m%d'))}{amount}"
                                    unique_keys_secondary.add(unique_key)
                                dict_secondary_id[platform_name][idx] = all_ids_secondary
                                dict_secondary_id_unique[platform_name][idx] = unique_keys_secondary
                                df.at[idx, f'Unique Transaction {platform_name} ID'] = '|'.join(unique_keys_secondary)

                    rows_processed += 1
                    if rows_processed % 100 == 0:
                        progress = 30 + int((rows_processed / total_rows) * 10)
                        update_progress(self.progress_queue, progress)

            # PHASE 2: Build complete relationship ID mappings
            logging.info("Phase 2: Building complete relationship ID mappings")
            
            # First process base platform data to build initial relationship IDs
            for idx, row in base_df.iterrows():
                primary_key = base_platform.get_relationship_id_key(row)
                secondary_key = str(row[base_platform.get_secondary_id_field()]) if pd.notnull(row[base_platform.get_secondary_id_field()]) else ''
                id_value = row[base_platform.get_id_field()]

                # Handle preRelIDs
                if primary_key and secondary_key:
                    if secondary_key in pre_rel_ids and primary_key in pre_rel_ids:
                        combined_id = add_unique_id(pre_rel_ids[secondary_key], pre_rel_ids[primary_key])
                        pre_rel_ids[secondary_key] = combined_id
                        pre_rel_ids[primary_key] = combined_id
                    elif secondary_key in pre_rel_ids:
                        pre_rel_ids[primary_key] = pre_rel_ids[secondary_key]
                    elif primary_key in pre_rel_ids:
                        pre_rel_ids[secondary_key] = pre_rel_ids[primary_key]
                    else:
                        pre_rel_ids[secondary_key] = str(id_value)
                        pre_rel_ids[primary_key] = str(id_value)
                elif primary_key:
                    if primary_key not in pre_rel_ids:
                        pre_rel_ids[primary_key] = str(id_value)
                    else:
                        pre_rel_ids[primary_key] = add_unique_id(pre_rel_ids[primary_key], str(id_value))
                elif secondary_key:
                    if secondary_key not in pre_rel_ids:
                        pre_rel_ids[secondary_key] = str(id_value)
                    else:
                        pre_rel_ids[secondary_key] = add_unique_id(pre_rel_ids[secondary_key], str(id_value))

                if idx % 100 == 0:
                    progress = 40 + int((idx / len(base_df)) * 10)
                    update_progress(self.progress_queue, progress)

            # Then process other platforms to complete relationship ID mappings
            for platform_name, df in platform_dfs.items():
                if platform_name == base_platform.get_platform_name():
                    continue  # Skip base platform as it's already processed
                
                platform = self.platforms[platform_name]
                for idx, row in df.iterrows():
                    primary_key = platform.get_relationship_id_key(row)
                    secondary_key = str(row[platform.get_id_field()])

                    # Update relationship ID mappings
                    if primary_key and primary_key in pre_rel_ids and secondary_key and secondary_key in pre_rel_ids:
                        # Both keys exist, combine them
                        combined_id = add_unique_id(pre_rel_ids[primary_key], pre_rel_ids[secondary_key])
                        pre_rel_ids[primary_key] = combined_id
                        pre_rel_ids[secondary_key] = combined_id
                    elif primary_key and primary_key in pre_rel_ids:
                        if secondary_key:
                            pre_rel_ids[secondary_key] = pre_rel_ids[primary_key]
                    elif secondary_key and secondary_key in pre_rel_ids:
                        if primary_key:
                            pre_rel_ids[primary_key] = pre_rel_ids[secondary_key]

                if len(df) > 0 and idx % 100 == 0:
                    progress = 50 + int((idx / len(df)) * 10)
                    update_progress(self.progress_queue, progress)

            # PHASE 3: Apply final relationship IDs to all rows
            logging.info("Phase 3: Applying final relationship IDs to all rows")
            
            # Apply to base platform
            for idx, row in base_df.iterrows():
                primary_key = base_platform.get_relationship_id_key(row)
                secondary_key = str(row[base_platform.get_secondary_id_field()]) if pd.notnull(row[base_platform.get_secondary_id_field()]) else ''
                id_value = row[base_platform.get_id_field()]
                
                # Apply the most up-to-date relationship ID
                if primary_key and primary_key in pre_rel_ids:
                    base_df.at[idx, 'Relationship ID'] = pre_rel_ids[primary_key]
                elif secondary_key and secondary_key in pre_rel_ids:
                    base_df.at[idx, 'Relationship ID'] = pre_rel_ids[secondary_key]
                else:
                    base_df.at[idx, 'Relationship ID'] = id_value

                if idx % 100 == 0:
                    progress = 60 + int((idx / len(base_df)) * 10)
                    update_progress(self.progress_queue, progress)

            # Apply to other platforms
            for platform_name, df in platform_dfs.items():
                if platform_name == base_platform.get_platform_name():
                    continue  # Skip base platform as it's already processed
                
                platform = self.platforms[platform_name]
                for idx, row in df.iterrows():
                    primary_key = platform.get_relationship_id_key(row)
                    secondary_key = str(row[platform.get_id_field()])
                    
                    # Apply the most up-to-date relationship ID
                    if primary_key and primary_key in pre_rel_ids:
                        df.at[idx, 'Relationship ID'] = pre_rel_ids[primary_key]
                    elif secondary_key and secondary_key in pre_rel_ids:
                        df.at[idx, 'Relationship ID'] = pre_rel_ids[secondary_key]
                    else:
                        # Handle cases where no match is found
                        unique_id = platform.get_relationship_id_key(row)
                        if pd.isna(unique_id) or unique_id == '':
                            unique_id = generate_fallback_id(row)
                        df.at[idx, 'Relationship ID'] = unique_id

                if len(df) > 0 and idx % 100 == 0:
                    progress = 70 + int((idx / len(df)) * 5)
                    update_progress(self.progress_queue, progress)

            # Find unique transactions across platforms and set duplicate flags
            all_unique_transactions = set()
            for platform_name in dict_primary_unique:
                for idx in dict_primary_unique[platform_name]:
                    all_unique_transactions.update(dict_primary_unique[platform_name][idx])
            for platform_name in dict_secondary_id_unique:
                for idx in dict_secondary_id_unique[platform_name]:
                    all_unique_transactions.update(dict_secondary_id_unique[platform_name][idx])

            for platform_name, df in platform_dfs.items():
                platform = self.platforms[platform_name]
                if platform.is_base_platform():
                    for idx, row in df.iterrows():
                        unique_key = row['Unique Transaction']
                        if unique_key in all_unique_transactions:
                            df.at[idx, platform.get_duplicate_column_name()] = 'Duplicate'
                        else:
                            df.at[idx, platform.get_duplicate_column_name()] = 'Not Duplicate'

            # Handle additional processing for each platform
            for platform_name, df in platform_dfs.items():
                platform = self.platforms[platform_name]
                if not platform.is_base_platform():
                    for idx, row in df.iterrows():
                        # Handle ActBlue recurring donations
                        is_recurring = row.get('Is Recurring', False)
                        
                        # Handle both string and boolean values
                        if isinstance(is_recurring, str):
                            is_recurring = is_recurring.upper() == 'TRUE'
                        
                        if is_recurring:
                            df.at[idx, 'Is Recurring'] = True
                            order_number = row[platform.get_id_field()]  # This is Order Number for ActBlue
                            if pd.notnull(order_number) and order_number != '':
                                df.at[idx, 'Recurring ID'] = order_number
                        else:
                            df.at[idx, 'Is Recurring'] = False
                else:
                    for idx, row in df.iterrows():
                        # Handle EveryAction recurring flag
                        if row.get('Is Recurring Commitment', 0) == 1:
                            df.at[idx, 'Is Recurring Commitment'] = True
                        elif row.get('Is Recurring Commitment', 0) == 0:
                            df.at[idx, 'Is Recurring Commitment'] = False

            logging.info("Transaction values generation completed")
            return platform_dfs

        except Exception as e:
            logging.error("Error in generate_transaction_values function")
            logging.error(str(e))
            logging.error(traceback.format_exc())
            raise

    def create_final_file(self, platform_dfs):
        logging.info("Creating final file")
        try:
            final_dfs = []
            for platform_name, df in platform_dfs.items():
                platform_obj = self.platforms[platform_name]
                processed_df = platform_obj.process_data(df)
                final_dfs.append(processed_df)

            # Combine the processed dataframes
            logging.info("Combining processed dataframes")
            final_columns = [
                'Relationship ID', 'Transaction ID', 'Giving Platform', 'Secondary ID',
                'Date Clean', 'Amount', 'Is Recurring', 'Recipient', 'Contribution Form URL',
                'Display Name', 'Donor First Name', 'Donor Last Name', 'Donor Address Line 1',
                'Donor City', 'Donor State', 'Donor ZIP', 'Donor Country', 'Donor Occupation',
                'Donor Employer', 'Donor Email', 'Donor Phone', 'Recurring ID',
                'Initial Recurring Contribution Date', 'Match?'
            ]
            
            # Ensure all columns exist in all dataframes
            for df in final_dfs:
                for col in final_columns:
                    if col not in df.columns:
                        df[col] = ''

            # Concatenate dataframes
            final_df = pd.concat([df[final_columns] for df in final_dfs], ignore_index=True)
            update_progress(self.progress_queue, 75)

            # Store original values before converting to numeric
            original_values = final_df['Relationship ID'].copy()

            # Convert 'Relationship ID' to numeric where possible, and 'NaN' for non-convertible values
            final_df['Relationship ID'] = pd.to_numeric(final_df['Relationship ID'], errors='coerce')

            # Fill back non-numeric (NaN) values with their original string equivalents
            final_df['Relationship ID'].fillna(original_values, inplace=True)

            # Store original Recurring ID values
            recurring_ids = final_df['Recurring ID'].copy()

            # Map Is Recurring to 'TRUE'/'FALSE'
            final_df['Is Recurring'] = final_df.get('Is Recurring', False).map({True: 'TRUE', False: 'FALSE'})
            
            # Convert 'Date Clean' to date format
            final_df['Date Clean'] = pd.to_datetime(final_df['Date Clean']).dt.date

            # Restore Recurring IDs
            final_df['Recurring ID'] = recurring_ids

            # Set invalid email values to null
            final_df.loc[~final_df['Donor Email'].str.contains('@', na=False), 'Donor Email'] = None

            # Sort the final dataframe
            logging.info("Sorting final dataframe")
            final_df = final_df.sort_values(by=['Date Clean', 'Amount'])
            update_progress(self.progress_queue, 80)

            logging.info("Final file created successfully")
            return final_df

        except Exception as e:
            logging.error("Error in create_final_file function")
            logging.error(str(e))
            logging.error(traceback.format_exc())
            raise

if __name__ == "__main__":
    root = tk.Tk()
    app = DynamicTransactionCompiler(root)
    root.mainloop()
