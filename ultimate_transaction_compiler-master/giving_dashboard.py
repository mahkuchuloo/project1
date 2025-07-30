import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os
import subprocess
import threading
import time
from utils import update_progress
from dictionary_lookup_manager import DictionaryLookupManager
from shared_ui_components import BaseToolFrame

class GivingDashboard(BaseToolFrame):
    def __init__(self, master):
        super().__init__(master, "Giving Dashboard")
        self.dict_manager = DictionaryLookupManager('lookup_dictionaries.json')
        self.final_data = None
        self.input_file_path = ""
        self.output_path = ""
        self.output_file = ""

    def start_processing(self):
        if not self.input_file_path:
            self.log("Please import the final file first.")
            return

        self.output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"Input A {datetime.now().strftime('%d%m%Y - %H%M%S')}.xlsx"
        )

        if not self.output_path:
            self.log("Please set the output path.")
            return

        # Disable buttons during processing
        self.import_button['state'] = tk.DISABLED
        self.process_button['state'] = tk.DISABLED

        # Start processing in a separate thread
        threading.Thread(target=self.process_data, daemon=True).start()

    def process_data(self):
        start_time = time.time()
        try:
            # Read the input file
            update_progress(self.progress_queue, 0)
            self.log("Reading input file...")
            self.final_data = pd.read_excel(self.input_file_path)
            update_progress(self.progress_queue, 10)

            # Define the steps for processing
            steps = [
                ("Generating Gift Range Values", self.generate_gift_range_values),
                ("Generating Gift Number Segment", self.generate_gift_number_segment),
                ("Applying Lookup Dictionaries", self.apply_lookup_dictionaries),
                ("Generating Last Gift Values", self.generate_last_gift_values),
                ("Generating Income Segment", self.generate_income_segment),
                ("Setting Month Column", self.set_month_column),
                ("Setting Year Column", self.set_year_column),
                ("Generating Last Gift Date Segment Values", self.generate_last_gift_date_segment_values),
                ("Generating Gift Date Segment Values", self.generate_gift_date_segment_values)
            ]

            # Process each step
            for i, (step_name, step_function) in enumerate(steps, 1):
                self.log(f"Step {i}/{len(steps)}: {step_name}")
                step_function()
                update_progress(self.progress_queue, 10 + (i / len(steps)) * 80)

            # Save the processed data
            self.log("Saving processed data...")
            self.final_data.to_excel(self.output_path, index=False)
            self.output_file = self.output_path
            update_progress(self.progress_queue, 100)

            self.log(f"Data processing completed. Output saved to: {self.output_file}")
        except Exception as e:
            self.log(f"Error during data processing: {str(e)}")
        finally:
            # Re-enable buttons after processing
            processing_time = time.time() - start_time
            self.master.after(0, lambda: self.update_ui_after_processing(processing_time))

    def update_ui_after_processing(self, processing_time):
        self.progress_bar['value'] = 100
        self.process_button.config(state=tk.NORMAL, text="Processing Complete")
        
        result_text = f"Final file generated in {processing_time:.2f} seconds.\n"
        result_text += f"File saved at: {self.output_file}"
        
        self.result_label.config(text=result_text)
        self.file_link.config(text="Open File")

        self.import_button['state'] = tk.NORMAL

    def generate_gift_range_values(self):
        def get_gift_range_value(amount):
            ranges = [
                (1000000, "$1M+"), (500000, "$500K+"), (200000, "$200K+"), (100000, "$100K+"),
                (75000, "$75K+"), (50000, "$50K+"), (25000, "$25K+"), (10000, "$10K+"),
                (5000, "$5K+"), (2500, "$2.5K+"), (1500, "$1.5K+"), (1000, "$1K+"),
                (500, "$500+"), (200, "$200+"), (100, "$100+"), (75, "$75+"),
                (50, "$50+"), (25, "$25+"), (10, "$10+"), (5, "$5+")
            ]
            for threshold, label in ranges:
                if amount >= threshold:
                    return label
            return "Less than $5" if amount > 0 else "No Amount"

        self.final_data['Gift Range Chart'] = self.final_data['Amount'].apply(get_gift_range_value)

    def generate_gift_number_segment(self):
        # Sort the data by Relationship ID and Date Clean
        self.final_data = self.final_data.sort_values(['Relationship ID', 'Date Clean'])
        
        # Group by Relationship ID and create a cumulative count
        self.final_data['Gift Number'] = self.final_data.groupby('Relationship ID').cumcount() + 1
        
        # Create a function to determine the gift segment
        def get_gift_segment(gift_number):
            if gift_number == 1:
                return "First Gift"
            elif gift_number == 2:
                return "Second Gift"
            elif gift_number == 3:
                return "Third Gift"
            elif gift_number == 4:
                return "Fourth Gift"
            else:
                return f"{gift_number}th Gift"
        
        # Apply the function to create the Gift Segment column
        self.final_data['Gift Segment'] = self.final_data['Gift Number'].apply(get_gift_segment)

    def apply_lookup_dictionaries(self):
        self.final_data = self.dict_manager.apply_lookup_dictionaries(self.final_data)

    def generate_last_gift_values(self):
        self.final_data = self.dict_manager.get_last_gift_columns(self.final_data)

    def generate_income_segment(self):
        def get_income_segment(row):
            if pd.isnull(row['Last Gift Date Clean']):
                return "New"
            elif (row['Date Clean'] - row['Last Gift Date Clean']).days > 395:
                return "Restore"
            elif row['Is Recurring'] == 'monthly' or row['Last Gift Is Recurring'] == 'monthly':
                return "Retained"
            else:
                return "Returning"

        # Ensure 'Date Clean' and 'Last Gift Date Clean' are in datetime format
        self.final_data['Date Clean'] = pd.to_datetime(self.final_data['Date Clean'])
        self.final_data['Last Gift Date Clean'] = pd.to_datetime(self.final_data['Last Gift Date Clean'])

        # Apply the function to create the Income Segment column
        self.final_data['Income Segment'] = self.final_data.apply(get_income_segment, axis=1)

    def set_month_column(self):
        self.final_data['Month'] = pd.to_datetime(self.final_data['Date Clean']).dt.month

    def set_year_column(self):
        self.final_data['Year'] = pd.to_datetime(self.final_data['Date Clean']).dt.year

    def generate_last_gift_date_segment_values(self):
        self.final_data['Last Gift Date Segment'] = self.final_data['Last Gift Date Clean'].apply(self.get_date_segment)

    def generate_gift_date_segment_values(self):
        self.final_data['Gift Date Segment'] = self.final_data['Date Clean'].apply(self.get_date_segment)

    def get_date_segment(self, date_value):
        if pd.isnull(date_value):
            return ""
        
        today = pd.Timestamp.now().floor('D')
        date_value = pd.to_datetime(date_value).floor('D')
        
        if date_value == today:
            return "A. Earlier Today"
        elif date_value == today - pd.Timedelta(days=1):
            return "B. Yesterday"
        elif date_value >= today - pd.Timedelta(days=today.dayofweek):
            return "C. This Week"
        elif date_value >= today - pd.Timedelta(days=today.dayofweek+7):
            return "D. Last Week"
        elif date_value.month == today.month and date_value.year == today.year:
            return "E. This Month"
        elif (date_value.year == today.year and date_value.month == today.month - 1) or \
                (date_value.year == today.year - 1 and today.month == 1 and date_value.month == 12):
            return "F. Last Month"
        elif date_value.quarter == today.quarter and date_value.year == today.year:
            return "G. This Quarter"
        elif (date_value.year == today.year and date_value.quarter == today.quarter - 1) or \
                (date_value.year == today.year - 1 and today.quarter == 1 and date_value.quarter == 4):
            return "H. Last Quarter"
        elif date_value.year == today.year:
            return "I. This Year"
        elif date_value.year == today.year - 1:
            return "J. Last Year"
        else:
            return "K. Before Last Year"

if __name__ == "__main__":
    root = tk.Tk()
    app = GivingDashboard(root)
    root.mainloop()
