import sys
import os
import traceback

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
from rfm_analyzer.rfm_score import RFMScorer
from .column_selection_dialog import ColumnSelectionDialog
from .column_config_manager import ColumnConfigManager

class AbstractRFMAnalyzer(BaseToolFrame):
    def __init__(self, master):
        super().__init__(master, "Abstract RFM Analyzer")
        self.dict_manager = DictionaryLookupManager('rfm_lookup_dictionaries.json')
        self.column_manager = ColumnConfigManager()
        self.final_data = None
        self.output_path = None
        self.input_file_path = None
        self.dictionary_cache = {}
        
        # Define column order to match final_rfm_analyzer
        self.column_order = [
            'Relationship VAN ID',
            'Email',
            'Phone',
            'Address',
            'City',
            'State',
            'Zip',
            'Total Number of Gifts',
            'Lifetime Giving',
            'Last Gift Date',
            'Last Gift Amount',
            'Last Gift Platform',
            'Last Gift Campaign Name',
            'Last Gift Appeal Name',
            'Last Gift Date Range',
            'Last Gift Amount Range',
            'First Gift Date',
            'First Gift Amount',
            'First Gift Platform',
            'First Gift Campaign Name',
            'First Gift Appeal Name',
            'First Gift Date Range A',
            'First Gift Date Range B',
            'First Gift Amount Range',
            'Largest Gift Date',
            'Largest Gift Amount',
            'Largest Gift Platform',
            'Largest Gift Campaign Name',
            'Largest Gift Appeal Name',
            'Largest Gift Date Range',
            'Largest Gift Amount Range',
            'Last Monthly Gift Date',
            'Last Monthly Gift Amount',
            'Last Monthly Gift Date Range',
            'Last Monthly Gift Amount Range',
            'Digital Monthly Indicator',
            'Primary Giving Platform',
            'Primary Giving Platform %',
            'Giving Segment A',
            'Giving Segment B',
            'Recency Criteria',
            'Frequency Criteria',
            'Monetary Criteria',
            'RFM Score',
            'RFM Percentile',
            'Recency Score',
            'Frequency Score',
            'Monetary Score',
            'Recency Percentile',
            'Frequency Percentile',
            'Monetary Percentile'
        ]
        
        # Add Columns menu
        columns_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Columns", menu=columns_menu)
        columns_menu.add_command(label="Column Selection", command=self.open_column_selection)
        
        # Add scoring method selection
        scoring_frame = tk.Frame(self)
        scoring_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Scoring methods
        self.scoring_methods = {
            "Percentile (Original VBA)": RFMScorer.percentile_scoring,
            "Quartile": RFMScorer.quartile_scoring,
            "Equal Width": RFMScorer.equal_width_scoring,
            "Z-Score": RFMScorer.zscore_scoring,
            "Logarithmic": RFMScorer.logarithmic_scoring
        }
        
        tk.Label(scoring_frame, text="RFM Scoring Method:").pack(side=tk.LEFT)
        self.scoring_method = ttk.Combobox(scoring_frame, 
                                         values=list(self.scoring_methods.keys()),
                                         state="readonly")
        self.scoring_method.set("Percentile (Original VBA)")
        self.scoring_method.pack(side=tk.LEFT, padx=5)
        
        # Add threshold inputs
        threshold_frame = tk.Frame(self)
        threshold_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(threshold_frame, text="Custom Thresholds (comma-separated):").pack(side=tk.LEFT)
        self.threshold_entry = tk.Entry(threshold_frame)
        self.threshold_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.threshold_entry.insert(0, "100,500,1000,5000,10000")
        
        # Add threshold scoring to methods after creating entry
        self.scoring_methods["Threshold"] = self.threshold_scoring

    def open_column_selection(self):
        """Open the column selection dialog"""
        dialog = ColumnSelectionDialog(self)
        self.wait_window(dialog)

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
            self.log(f"Error details: {traceback.format_exc()}")
            self.master.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {str(e)}\n\nDetails: {traceback.format_exc()}"))
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

    def calculate_percentile_category(self, value, series):
        """Calculate percentile category (A. Top 1%, B. Top 2%, etc.)"""
        try:
            # Calculate percentile rank for all values
            ranks = series.rank(pct=True) * 100
            # Find the percentile for the current value
            percentile = ranks[series == value].iloc[0]
            
            if percentile >= 99:
                return "A. Top 1%"
            elif percentile >= 98:
                return "B. Top 2%"
            elif percentile >= 95:
                return "C. Top 5%"
            elif percentile >= 90:
                return "D. Top 10%"
            elif percentile >= 50:
                return "E. 50%"
            else:
                return "F. Bottom 50%"
        except Exception as e:
            self.log(f"Error calculating percentile category for value {value}: {str(e)}")
            return "F. Bottom 50%"  # Default to bottom 50% on error

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

    def calculate_giving_segment_a(self, total_gifts):
        """Calculate Giving Segment A based on total number of gifts"""
        if total_gifts == 0:
            return "A. No Gifts"
        elif total_gifts == 1:
            return "A. First Gift Segment"
        elif total_gifts == 2:
            return "B. Second Gift Segment"
        elif total_gifts == 3:
            return "C. Third Gift Segment"
        elif total_gifts == 4:
            return "D. Fourth Gift Segment"
        elif total_gifts == 5:
            return "E. Fifth Gift Segment"
        elif total_gifts == 6:
            return "F. Sixth Gift Segment"
        elif total_gifts == 7:
            return "G. Seventh Gift Segment"
        elif total_gifts > 7:
            return "H. Seventh Gift Segment +"
        return "I. None"

    def calculate_giving_segment_b(self, total_gifts):
        """Calculate Giving Segment B based on total number of gifts"""
        if total_gifts < 13:
            return "A. Less than 13"
        elif 13 <= total_gifts <= 17:
            return "A. 13-17 gifts"
        elif 18 <= total_gifts <= 22:
            return "B. 18-22 gifts"
        elif 23 <= total_gifts <= 27:
            return "C. 23-27 gifts"
        elif 28 <= total_gifts <= 32:
            return "D. 28-32 gifts"
        elif 33 <= total_gifts <= 36:
            return "E. 33-36 gifts"
        elif total_gifts >= 37:
            return "F. 37+ gifts"
        return "G. None"

    def calculate_date_range(self, date):
        """Calculate date range category"""
        if pd.isna(date):
            return "No Date"
            
        now = pd.Timestamp.now()
        
        if date.year < now.year - 1:
            return "Before Last Year"
        elif date.year == now.year - 1:
            return "Last Year"
        elif date.year == now.year + 1:
            return "Next Year"
        elif date.year > now.year + 1:
            return "After Next Year"
        
        # Current year comparisons
        if date.month == now.month - 1:
            return "Last Month"
        elif (date.month - 1) // 3 == (now.month - 1) // 3 - 1:
            return "Last Quarter"
        elif date.isocalendar()[1] == now.isocalendar()[1] - 1:
            return "Last Week"
        elif date.isocalendar()[1] == now.isocalendar()[1] + 1:
            return "Next Week"
        
        # Same month comparisons
        if date.month == now.month:
            if date.day == now.day - 1:
                return "Yesterday"
            elif date.day == now.day + 1:
                if date.hour < 12:
                    return "Tomorrow morning"
                elif date.hour < 18:
                    return "Tomorrow afternoon"
                else:
                    return "Tomorrow evening"
            elif date.day == now.day:
                if date < now:
                    return "Earlier Today"
                elif date.hour >= 18:
                    return "Tonight"
                else:
                    return "Today"
        
        return "This Year"

    def calculate_first_gift_date_range_a(self, date):
        """Calculate First Gift Date Range A"""
        if pd.isna(date):
            return "No Date"
            
        if date.year < 2017:
            return "A. Before 2017"
        elif 2017 <= date.year <= 2019:
            return "B. 2017-2019"
        elif date.year == 2019:
            return "C. 2019"
        elif date.year == 2020:
            if date.month <= 5 and date.day <= 25:
                return "D. 2020 Pre Anguish and Action"
            elif date.month == 5:
                return "E. May 2020 Anguish and Action"
            elif date.month == 6:
                return "F. June 2020 Anguish and Action"
            elif date.month == 7:
                return "G. July 2020 Anguish and Action"
            elif date.month == 8:
                return "H. August 2020 Anguish And Action"
            else:
                return "J. September - December 2020"
        elif date.year == 2021:
            return "K. 2021"
        elif date.year == 2022:
            return "L. 2022"
        return "M. After 2022"

    def rfm_analyzer(self, df):
        try:
            self.log("Starting RFM analysis...")
            
            # Get selected columns
            selected_columns = self.column_manager.get_columns()
            
            # Convert date columns to datetime
            self.log("Converting date columns...")
            df['Date Clean'] = pd.to_datetime(df['Date Clean'])
            
            # First identify all Relationship IDs that have recurring donations
            self.log("Identifying recurring donations...")
            recurring_ids = set(df[df['Recurring ID'].notna()]['Relationship ID'].unique())
            
            # Group by Relationship ID
            self.log("Grouping by Relationship ID...")
            grouped = df.groupby('Relationship ID')
            
            # Calculate RFM components for each customer
            rfm_data = []
            total_groups = len(grouped)
            
            self.log(f"Processing {total_groups} groups...")
            
            for i, (name, group) in enumerate(grouped):
                try:
                    self.log(f"Processing group {i+1}/{total_groups}: {name}")
                    
                    # Calculate RFM metrics exactly as VBA does
                    recency = group['Date Clean'].max()
                    frequency = len(group['Transaction ID'].unique())
                    monetary = group['Amount'].sum()
                    
                    self.log(f"Group {name} metrics - Recency: {recency}, Frequency: {frequency}, Monetary: {monetary}")
                    
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
                        'Recency Criteria': recency,
                        'Frequency Criteria': frequency,
                        'Monetary Criteria': monetary,
                        'Digital Monthly Indicator': 'Digital Monthly' if name in recurring_ids else 'Not Digital Monthly'
                    }
                    
                    # Add lookup dictionary columns
                    self.log(f"Processing lookup dictionaries for group {name}")
                    last_gift_idx = group['Date Clean'].idxmax()
                    for lookup in self.dict_manager.lookups:
                        try:
                            if lookup.get('use_multiple_values', False):
                                # Get dictionary DataFrame from cache
                                dict_df = self.get_dictionary_df(lookup['path'])
                                value_columns = dict_df.columns[1:]  # All columns except the first
                                
                                if lookup.get('include_in_last_gift', False):
                                    # Add all columns from last gift at once
                                    for col in value_columns:
                                        if f"Last Gift {col}" in selected_columns:
                                            customer_info[f"Last Gift {col}"] = group.loc[last_gift_idx, col]
                                else:
                                    # Add all columns from first row at once
                                    for col in value_columns:
                                        if col in selected_columns:
                                            customer_info[col] = group.iloc[0][col]
                            else:
                                # Handle standard single-value dictionaries
                                if lookup.get('include_in_last_gift', False):
                                    if f"Last Gift {lookup['output_column']}" in selected_columns:
                                        customer_info[f"Last Gift {lookup['output_column']}"] = group.loc[last_gift_idx, lookup['output_column']]
                                else:
                                    if lookup['output_column'] in selected_columns:
                                        customer_info[lookup['output_column']] = group[lookup['output_column']].iloc[0]
                        except Exception as e:
                            self.log(f"Error processing lookup dictionary for group {name}: {str(e)}")
                            raise
                    
                    # Calculate first gift information if selected
                    first_gift_idx = group['Date Clean'].idxmin()
                    if 'First Gift Date' in selected_columns:
                        customer_info['First Gift Date'] = group.loc[first_gift_idx, 'Date Clean']
                    if 'First Gift Amount' in selected_columns:
                        customer_info['First Gift Amount'] = group.loc[first_gift_idx, 'Amount']
                    if 'First Gift Platform' in selected_columns and 'Giving Platform' in group.columns:
                        customer_info['First Gift Platform'] = group.loc[first_gift_idx, 'Giving Platform']
                    
                    # Calculate largest gift information if selected
                    largest_gift_idx = group['Amount'].idxmax()
                    if 'Largest Gift Date' in selected_columns:
                        customer_info['Largest Gift Date'] = group.loc[largest_gift_idx, 'Date Clean']
                    if 'Largest Gift Amount' in selected_columns:
                        customer_info['Largest Gift Amount'] = group.loc[largest_gift_idx, 'Amount']
                    if 'Largest Gift Platform' in selected_columns and 'Giving Platform' in group.columns:
                        customer_info['Largest Gift Platform'] = group.loc[largest_gift_idx, 'Giving Platform']
                    
                    # Calculate monthly gift information if selected
                    monthly_gifts = group[group['Recurring ID'].notna()]
                    if not monthly_gifts.empty:
                        last_monthly_idx = monthly_gifts['Date Clean'].idxmax()
                        if 'Last Monthly Gift Date' in selected_columns:
                            customer_info['Last Monthly Gift Date'] = monthly_gifts.loc[last_monthly_idx, 'Date Clean']
                        if 'Last Monthly Gift Amount' in selected_columns:
                            customer_info['Last Monthly Gift Amount'] = monthly_gifts.loc[last_monthly_idx, 'Amount']
                    
                    # Calculate ranges and indicators if selected
                    if 'Last Gift Amount Range' in selected_columns:
                        customer_info['Last Gift Amount Range'] = self.calculate_gift_amount_range(customer_info['Last Gift Amount'])
                    if 'First Gift Amount Range' in selected_columns and 'First Gift Amount' in customer_info:
                        customer_info['First Gift Amount Range'] = self.calculate_gift_amount_range(customer_info['First Gift Amount'])
                    if 'Largest Gift Amount Range' in selected_columns and 'Largest Gift Amount' in customer_info:
                        customer_info['Largest Gift Amount Range'] = self.calculate_gift_amount_range(customer_info['Largest Gift Amount'])
                    if 'Last Monthly Gift Amount Range' in selected_columns and 'Last Monthly Gift Amount' in customer_info:
                        customer_info['Last Monthly Gift Amount Range'] = self.calculate_gift_amount_range(customer_info['Last Monthly Gift Amount'])
                    
                    # Calculate date ranges if selected
                    if 'First Gift Date Range A' in selected_columns and 'First Gift Date' in customer_info:
                        customer_info['First Gift Date Range A'] = self.calculate_first_gift_date_range_a(customer_info['First Gift Date'])
                    if 'First Gift Date Range B' in selected_columns and 'First Gift Date' in customer_info:
                        customer_info['First Gift Date Range B'] = self.calculate_date_range(customer_info['First Gift Date'])
                    if 'Last Gift Date Range' in selected_columns:
                        customer_info['Last Gift Date Range'] = self.calculate_date_range(customer_info['Last Gift Date'])
                    if 'Largest Gift Date Range' in selected_columns and 'Largest Gift Date' in customer_info:
                        customer_info['Largest Gift Date Range'] = self.calculate_date_range(customer_info['Largest Gift Date'])
                    if 'Last Monthly Gift Date Range' in selected_columns and 'Last Monthly Gift Date' in customer_info:
                        customer_info['Last Monthly Gift Date Range'] = self.calculate_date_range(customer_info['Last Monthly Gift Date'])
                    
                    # Calculate giving segments if selected
                    if 'Giving Segment A' in selected_columns:
                        customer_info['Giving Segment A'] = self.calculate_giving_segment_a(frequency)
                    if 'Giving Segment B' in selected_columns:
                        customer_info['Giving Segment B'] = self.calculate_giving_segment_b(frequency)
                    
                    rfm_data.append(customer_info)
                    
                    # Update progress
                    if i % 100 == 0 or i == total_groups - 1:
                        progress = 30 + int((i / total_groups) * 60)
                        self.progress_queue.put(progress)
                        
                except Exception as e:
                    self.log(f"Error processing group {name}: {str(e)}")
                    self.log(f"Group data: {group.head()}")
                    raise
            
            # Convert to DataFrame
            self.log("Converting to DataFrame...")
            result = pd.DataFrame(rfm_data)
            
            # Get selected scoring method
            self.log("Calculating RFM scores...")
            scoring_method = self.scoring_methods[self.scoring_method.get()]
            
            # Calculate RFM Scores using selected method
            if 'Recency Score' in selected_columns:
                result['Recency Score'] = scoring_method(result['Recency Criteria'], ascending=False)
            
            if 'Frequency Score' in selected_columns:
                result['Frequency Score'] = scoring_method(result['Frequency Criteria'], ascending=True)
            
            if 'Monetary Score' in selected_columns:
                result['Monetary Score'] = scoring_method(result['Monetary Criteria'], ascending=True)
            
            if 'RFM Score' in selected_columns:
                result['RFM Score'] = (
                    result.get('Recency Score', 0) + 
                    result.get('Frequency Score', 0) + 
                    result.get('Monetary Score', 0)
                )
            
            # Calculate percentile categories if selected
            self.log("Calculating percentile categories...")
            if 'RFM Percentile' in selected_columns and 'RFM Score' in result.columns:
                result['RFM Percentile'] = result.apply(
                    lambda x: self.calculate_percentile_category(x['RFM Score'], result['RFM Score']), axis=1)
            
            if 'Recency Percentile' in selected_columns and 'Recency Score' in result.columns:
                result['Recency Percentile'] = result.apply(
                    lambda x: self.calculate_percentile_category(x['Recency Score'], result['Recency Score']), axis=1)
            
            if 'Frequency Percentile' in selected_columns and 'Frequency Score' in result.columns:
                result['Frequency Percentile'] = result.apply(
                    lambda x: self.calculate_percentile_category(x['Frequency Score'], result['Frequency Score']), axis=1)
            
            if 'Monetary Percentile' in selected_columns and 'Monetary Score' in result.columns:
                result['Monetary Percentile'] = result.apply(
                    lambda x: self.calculate_percentile_category(x['Monetary Score'], result['Monetary Score']), axis=1)
            
            # Filter and order columns
            if selected_columns:
                # Always include Relationship VAN ID first
                ordered_columns = ['Relationship VAN ID']
                
                # Add other selected columns in the defined order
                for col in self.column_order:
                    if col in selected_columns and col != 'Relationship VAN ID':
                        ordered_columns.append(col)
                
                # Filter to only include columns that were actually calculated
                available_columns = set(result.columns)
                valid_columns = [col for col in ordered_columns if col in available_columns]
                
                result = result[valid_columns]
            
            self.log("RFM analysis completed successfully")
            return result
            
        except Exception as e:
            self.log(f"Error in rfm_analyzer: {str(e)}")
            self.log(f"Error details: {traceback.format_exc()}")
            raise

if __name__ == "__main__":
    root = tk.Tk()
    app = AbstractRFMAnalyzer(root)
    root.mainloop()
