import sys
import os

# Add the parent directory to the sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from datetime import datetime
import time
import threading
from dictionary_lookup_manager import DictionaryLookupManager
from shared_ui_components import BaseToolFrame
from rfm_score import RFMScorer

class RFMAnalyzer(BaseToolFrame):
    def __init__(self, master):
        super().__init__(master, "RFM Analyzer")
        self.dict_manager = DictionaryLookupManager('rfm_lookup_dictionaries.json')
        self.final_data = None
        self.output_path = None
        self.input_file_path = None
        # Cache for dictionary DataFrames
        self.dictionary_cache = {}
        
        # Add scoring method selection
        self.scoring_methods = {
            "Percentile (Original VBA)": RFMScorer.percentile_scoring,
            "Quartile": RFMScorer.quartile_scoring,
            "Equal Width": RFMScorer.equal_width_scoring,
            "Z-Score": RFMScorer.zscore_scoring,
            "Logarithmic": RFMScorer.logarithmic_scoring
        }
        
        # Add scoring method dropdown
        scoring_frame = tk.Frame(self)
        scoring_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(scoring_frame, text="RFM Scoring Method:").pack(side=tk.LEFT)
        self.scoring_method = ttk.Combobox(scoring_frame, 
                                         values=list(self.scoring_methods.keys()),
                                         state="readonly")
        self.scoring_method.set("Percentile (Original VBA)")
        self.scoring_method.pack(side=tk.LEFT, padx=5)
        
        # Add threshold inputs for threshold scoring
        threshold_frame = tk.Frame(self)
        threshold_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(threshold_frame, text="Custom Thresholds (comma-separated):").pack(side=tk.LEFT)
        self.threshold_entry = tk.Entry(threshold_frame)
        self.threshold_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.threshold_entry.insert(0, "100,500,1000,5000,10000")
        
        # Add threshold scoring to methods after creating entry
        self.scoring_methods["Threshold"] = self.threshold_scoring

    def get_dictionary_df(self, path):
        """Get dictionary DataFrame from cache or load it."""
        if path not in self.dictionary_cache:
            self.dictionary_cache[path] = pd.read_excel(path)
        return self.dictionary_cache[path]

    def threshold_scoring(self, series, ascending=True):
        """Wrapper for threshold scoring that gets thresholds from UI"""
        try:
            thresholds = [float(x.strip()) for x in self.threshold_entry.get().split(',')]
            return RFMScorer.threshold_scoring(series, thresholds, ascending)
        except ValueError as e:
            messagebox.showerror("Error", "Invalid threshold values. Please enter comma-separated numbers.")
            raise e

    def start_processing(self):
        if not self.input_file_path:
            self.log("Please select the input file first.")
            return

        self.output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            initialfile=f"Output 3 {datetime.now().strftime('%d%m%Y - %H%M%S')}.xlsx"
        )

        if not self.output_path:
            self.log("Output file selection cancelled.")
            return

        self.reset_ui()
        self.process_button.config(state=tk.DISABLED)
        threading.Thread(target=self.process_data, daemon=True).start()

    def process_data(self):
        self.log("Processing data. Please wait...")
        self.progress_queue.put(0)

        start_time = time.time()

        try:
            # Read the input file
            self.final_data = self.read_input_file()
            self.progress_queue.put(10)
            self.log("File imported. Applying lookup dictionaries...")

            # Clear dictionary cache
            self.dictionary_cache.clear()

            # Apply lookup dictionaries using the dictionary manager
            self.final_data = self.dict_manager.apply_lookup_dictionaries(self.final_data)
            self.progress_queue.put(20)
            self.log("Lookup dictionaries applied. Performing RFM analysis...")

            # Perform RFM analysis
            result = self.rfm_analyzer(self.final_data)
            self.progress_queue.put(90)

            # Save the result
            result.to_excel(self.output_path, index=False)
            self.log(f"RFM Analysis output has been written to {self.output_path}")
            self.progress_queue.put(100)

            end_time = time.time()
            processing_time = end_time - start_time

            self.master.after(0, self.update_ui_after_processing, processing_time)

        except Exception as e:
            self.log(f"An error occurred during RFM analysis: {str(e)}")
            self.master.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {str(e)}"))
            self.master.after(0, lambda: self.process_button.config(state=tk.NORMAL))

    def update_ui_after_processing(self, processing_time):
        self.progress_bar['value'] = 100
        self.process_button.config(state=tk.NORMAL, bg='green', fg='white', text="Processing Complete")
        
        result_text = f"RFM Analysis completed in {processing_time:.2f} seconds.\n"
        result_text += f"Output saved at: {self.output_path}"
        
        self.result_label.config(text=result_text)
        self.file_link.config(text="Open File Location")

    def read_input_file(self):
        try:
            if self.input_file_path.endswith('.csv'):
                return pd.read_csv(self.input_file_path)
            else:
                return pd.read_excel(self.input_file_path)
        except Exception as e:
            self.log(f"Error reading input file: {str(e)}")
            return None

    def rfm_analyzer(self, df):
        # Convert date columns to datetime
        df['Date Clean'] = pd.to_datetime(df['Date Clean'])
        
        # First identify all Relationship IDs that have recurring donations
        recurring_ids = set(df[df['Recurring ID'].notna()]['Relationship ID'].unique())
        
        # Group by Relationship ID
        grouped = df.groupby('Relationship ID')
        
        # Calculate RFM components for each customer
        rfm_data = []
        total_groups = len(grouped)
        
        for i, (name, group) in enumerate(grouped):
            # Calculate RFM metrics exactly as VBA does
            recency = group['Date Clean'].max()
            frequency = len(group['Transaction ID'].unique())
            monetary = group['Amount'].sum()
            
            # Basic customer info
            customer_info = {
                'Relationship VAN ID': name,
                'Email': group['Donor Email'].iloc[0],
                'Phone': group['Donor Phone'].iloc[0],
                'Address': group['Donor Address Line 1'].iloc[0],
                'City': group['Donor City'].iloc[0],
                'State': group['Donor State'].iloc[0],
                'Zip': group['Donor ZIP'].iloc[0],
                'Total Number of Gifts': frequency,
                'Lifetime Giving': monetary,
                'Last Gift Date': recency,
                'Last Gift Amount': group.loc[group['Date Clean'].idxmax(), 'Amount'],
                'Last Gift Giving Platform': group.loc[group['Date Clean'].idxmax(), 'Giving Platform'],
                'Last Gift Designation': group.loc[group['Date Clean'].idxmax(), 'Recipient'],
                'Last Gift Transaction ID': group.loc[group['Date Clean'].idxmax(), 'Transaction ID'],
                'Recency Criteria': recency,
                'Frequency Criteria': frequency,
                'Monetary Criteria': monetary,
                'Digital Monthly Indicator': 'Digital Monthly' if name in recurring_ids else 'Not Digital Monthly'
            }
            
            # Add lookup dictionary columns
            last_gift_idx = group['Date Clean'].idxmax()
            for lookup in self.dict_manager.lookups:
                if lookup.get('use_multiple_values', False):
                    # Get dictionary DataFrame from cache
                    dict_df = self.get_dictionary_df(lookup['path'])
                    value_columns = dict_df.columns[1:]  # All columns except the first
                    
                    if lookup.get('include_in_last_gift', False):
                        # Add all columns from last gift at once
                        for col in value_columns:
                            customer_info[f"Last Gift {col}"] = group.loc[last_gift_idx, col]
                    else:
                        # Add all columns from first row at once
                        for col in value_columns:
                            customer_info[col] = group.iloc[0][col]
                else:
                    # Handle standard single-value dictionaries
                    if lookup.get('include_in_last_gift', False):
                        customer_info[f"Last Gift {lookup['output_column']}"] = group.loc[last_gift_idx, lookup['output_column']]
                    else:
                        customer_info[lookup['output_column']] = group[lookup['output_column']].iloc[0]
            
            rfm_data.append(customer_info)
            
            # Update progress
            if i % 100 == 0 or i == total_groups - 1:
                progress = 30 + int((i / total_groups) * 60)
                self.progress_queue.put(progress)
        
        # Convert to DataFrame
        result = pd.DataFrame(rfm_data)
        
        # Get selected scoring method
        scoring_method = self.scoring_methods[self.scoring_method.get()]
        
        # Calculate RFM Scores using selected method
        result['Recency Score'] = scoring_method(result['Recency Criteria'], ascending=False)
        result['Frequency Score'] = scoring_method(result['Frequency Criteria'], ascending=True)
        result['Monetary Score'] = scoring_method(result['Monetary Criteria'], ascending=True)
        result['RFM Score'] = result['Recency Score'] + result['Frequency Score'] + result['Monetary Score']
        
        # Calculate ranges and indicators
        result['Last Gift Amount Range'] = result['Last Gift Amount'].apply(self.calculate_gift_amount_range)
        
        return result

    def calculate_gift_amount_range(self, amount):
        ranges = [
            (0, 5, 'A. <$5'),
            (5, 10, 'B. $5 - $9.99'),
            (10, 25, 'C. $10 - $24'),
            (25, 50, 'D. $25 - $49'),
            (50, 100, 'E. $50 - $99'),
            (100, 250, 'F. $100 - $249'),
            (250, 500, 'G. $250 - $499'),
            (500, 1000, 'H. $500 - $999'),
            (1000, 2500, 'I. $1,000 - $2,499'),
            (2500, 5000, 'J. $2,500 - $4,999'),
            (5000, 10000, 'K. $5,000 - $9,999'),
            (10000, 25000, 'L. $10,000 - $24,999'),
            (25000, 50000, 'M. $25,000 - $49,999'),
            (50000, 100000, 'N. $50,000 - $99,999'),
            (100000, 250000, 'O. $100,000 - $249,999'),
            (250000, 500000, 'P. $250,000 - $499,999'),
            (500000, 1000000, 'Q. $500,000 - $999,999'),
            (1000000, 2500000, 'R. $1,000,000 - $2,499,999'),
            (2500000, 5000000, 'S. $2,500,000 - $4,999,999'),
            (5000000, 10000000, 'T. $5,000,000 - $9,999,999'),
            (10000000, 25000000, 'U. $10,000,000 - $24,999,999'),
            (25000000, 50000000, 'V. $25,000,000 - $49,999,999'),
            (50000000, 100000000, 'W. $50,000,000 - $99,999,999'),
            (100000000, 250000000, 'X. $100,000,000 - $249,999,999'),
            (250000000, 500000000, 'Y. $250,000,000 - $499,999,999'),
            (500000000, 1000000000, 'Z. $500,000,000 - $999,999,999'),
        ]
        
        if pd.isna(amount):
            return 'No Amount'
            
        for low, high, label in ranges:
            if low <= amount < high:
                return label
        return 'AA. $1,000,000,000+'

if __name__ == "__main__":
    root = tk.Tk()
    app = RFMAnalyzer(root)
    root.mainloop()
