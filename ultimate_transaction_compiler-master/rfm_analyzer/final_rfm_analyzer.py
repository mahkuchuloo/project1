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
from .rfm_score import RFMScorer

class FinalRFMAnalyzer(BaseToolFrame):
    def __init__(self, master):
        super().__init__(master, "Final RFM Analyzer")
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
            "Logarithmic": RFMScorer.logarithmic_scoring,
            "Threshold": self.threshold_scoring
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
        
        # Bind the selection event
        self.scoring_method.bind('<<ComboboxSelected>>', self.on_scoring_method_change)
        
        # Add threshold inputs for threshold scoring
        self.threshold_frame = tk.Frame(self)
        
        # Create a sub-frame for the label and info icon
        label_frame = tk.Frame(self.threshold_frame)
        label_frame.pack(side=tk.LEFT)
        
        tk.Label(label_frame, text="Custom Thresholds (comma-separated):").pack(side=tk.LEFT)
        
        # Add info icon with tooltip
        info_label = tk.Label(label_frame, text="ℹ️", cursor="hand2")
        info_label.pack(side=tk.LEFT, padx=2)
        
        # Create tooltip
        tooltip_text = ("How to use thresholds:\n"
                       "1. Enter values separated by commas (e.g., 100,500,1000,5000,10000)\n"
                       "2. Values determine score boundaries:\n"
                       "   - Values ≤ first threshold get score 1\n"
                       "   - Values ≤ second threshold get score 2\n"
                       "   - And so on until the last threshold\n"
                       "   - Values > last threshold get score 10\n"
                       "3. For Recency (days since last purchase), scoring is reversed")
        
        def show_tooltip(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            # Add a frame with a border
            frame = tk.Frame(tooltip, borderwidth=1, relief="solid", bg="lightyellow")
            frame.pack(fill="both", expand=True)
            
            # Add multiline label with tooltip text
            label = tk.Label(frame, text=tooltip_text, justify=tk.LEFT, 
                           bg="lightyellow", padx=5, pady=5, wraplength=400)
            label.pack()
            
            def hide_tooltip(event=None):
                tooltip.destroy()
            
            tooltip.bind('<Leave>', hide_tooltip)
            label.bind('<Leave>', hide_tooltip)
        
        info_label.bind('<Enter>', show_tooltip)
        
        self.threshold_entry = tk.Entry(self.threshold_frame)
        self.threshold_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.threshold_entry.insert(0, "100,500,1000,5000,10000")
        
        # Initially hide threshold frame
        if self.scoring_method.get() != "Threshold":
            self.threshold_frame.pack_forget()

    def on_scoring_method_change(self, event):
        """Handle scoring method selection change"""
        if self.scoring_method.get() == "Threshold":
            self.threshold_frame.pack(fill=tk.X, padx=5, pady=5)
        else:
            self.threshold_frame.pack_forget()

    def threshold_scoring(self, series, ascending=True):
        """Wrapper for threshold scoring that gets thresholds from UI"""
        try:
            thresholds = [float(x.strip()) for x in self.threshold_entry.get().split(',')]
            return RFMScorer.threshold_scoring(series, thresholds, ascending)
        except ValueError as e:
            messagebox.showerror("Error", "Invalid threshold values. Please enter comma-separated numbers.")
            raise e

    def get_dictionary_df(self, path):
        """Get dictionary DataFrame from cache or load it."""
        if path not in self.dictionary_cache:
            self.dictionary_cache[path] = pd.read_excel(path)
        return self.dictionary_cache[path]

    def calculate_gift_amount_range(self, amount):
        """Calculate gift amount range category"""
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

    def calculate_percentile_category_vectorized(self, series):
        """Vectorized version of percentile category calculation"""
        try:
            ranks = series.rank(pct=True) * 100
            conditions = [
                ranks >= 99,
                ranks >= 98,
                ranks >= 95,
                ranks >= 90,
                ranks >= 50
            ]
            choices = [
                "A. Top 1%",
                "B. Top 2%",
                "C. Top 5%",
                "D. Top 10%",
                "E. 50%"
            ]
            return np.select(conditions, choices, default="F. Bottom 50%")
        except Exception as e:
            self.log(f"Error in vectorized percentile calculation: {str(e)}")
            return pd.Series("F. Bottom 50%", index=series.index)

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

    def rfm_analyzer(self, df):
        try:
            self.log("Starting RFM analysis...")
            
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
                # Calculate RFM metrics with null handling
                valid_dates = group['Date Clean'].dropna()
                valid_transactions = group['Transaction ID'].dropna()
                valid_amounts = group['Amount'].dropna()
                
                recency = valid_dates.max() if len(valid_dates) > 0 else pd.NaT
                frequency = len(valid_transactions.unique())
                monetary = valid_amounts.sum() if len(valid_amounts) > 0 else 0
                
                # Get indices for different gift types
                valid_dates = group['Date Clean'].dropna()
                valid_amounts = group['Amount'].dropna()
                default_idx = group.index[0]
                
                first_gift_idx = valid_dates.idxmin() if len(valid_dates) > 0 else default_idx
                last_gift_idx = valid_dates.idxmax() if len(valid_dates) > 0 else default_idx
                largest_gift_idx = valid_amounts.idxmax() if len(valid_amounts) > 0 else default_idx
                
                # Get last monthly gift info (if exists)
                monthly_gifts = group[group['Recurring ID'].notna()]
                last_monthly_gift = None if monthly_gifts.empty else monthly_gifts.loc[monthly_gifts['Date Clean'].idxmax()]
                
                # Calculate primary giving platform
                platform_counts = group['Giving Platform'].value_counts()
                primary_platform = platform_counts.index[0] if not platform_counts.empty else None
                primary_platform_pct = (platform_counts.iloc[0] / len(group) * 100) if not platform_counts.empty else 0
                
                # Basic customer info with all columns from rfm_analyzer.py
                customer_info = {
                    'Relationship VAN ID': name,
                    'Email': group['Donor Email'].iloc[0] if 'Donor Email' in group else None,
                    'Phone': group['Donor Phone'].iloc[0] if 'Donor Phone' in group else None,
                    'Address': group['Donor Address Line 1'].iloc[0] if 'Donor Address Line 1' in group else None,
                    'City': group['Donor City'].iloc[0] if 'Donor City' in group else None,
                    'State': group['Donor State'].iloc[0] if 'Donor State' in group else None,
                    'Zip': group['Donor ZIP'].iloc[0] if 'Donor ZIP' in group else None,
                    'Total Number of Gifts': frequency,
                    'Lifetime Giving': monetary,
                     'Lifetime Giving Sum': group['Amount'].sum(),
                    'Recency Criteria': recency,
                    'Frequency Criteria': frequency,
                    'Monetary Criteria': monetary,
                    
                    # Additional columns with proper implementation
                    'First Gift Date': group.loc[first_gift_idx, 'Date Clean'],
                    'First Gift Amount': group.loc[first_gift_idx, 'Amount'],
                    'First Gift Platform': group.loc[first_gift_idx, 'Giving Platform'] if 'Giving Platform' in group else None,
                    'First Gift Campaign Name': group.loc[first_gift_idx, 'Campaign Name'] if 'Campaign Name' in group else None,
                    'First Gift Appeal Name': group.loc[first_gift_idx, 'Appeal Name'] if 'Appeal Name' in group else None,

                     # Last Gift Information
                        'Last Gift Date': recency,
                        'Last Gift Amount': group.loc[last_gift_idx, 'Amount'],
                        'Last Gift Platform': group.loc[last_gift_idx, 'Giving Platform'] if 'Giving Platform' in group else None,
                        'Last Gift Campaign Name': group.loc[last_gift_idx, 'Campaign Name'] if 'Campaign Name' in group else None,
                        'Last Gift Appeal Name': group.loc[last_gift_idx, 'Appeal Name'] if 'Appeal Name' in group else None,
                    
                    'Largest Gift Date': group.loc[largest_gift_idx, 'Date Clean'],
                    'Largest Gift Amount': group.loc[largest_gift_idx, 'Amount'],
                    'Largest Gift Platform': group.loc[largest_gift_idx, 'Giving Platform'] if 'Giving Platform' in group else None,
                    'Largest Gift Campaign Name': group.loc[largest_gift_idx, 'Campaign Name'] if 'Campaign Name' in group else None,
                    'Largest Gift Appeal Name': group.loc[largest_gift_idx, 'Appeal Name'] if 'Appeal Name' in group else None,
                    
                    'Last Monthly Gift Date': last_monthly_gift['Date Clean'] if last_monthly_gift is not None else None,
                    'Last Monthly Gift Amount': last_monthly_gift['Amount'] if last_monthly_gift is not None else None,
                    'Digital Monthly Indicator': 'Digital Monthly' if name in recurring_ids else 'Not Digital Monthly',
                    
                    'Primary Giving Platform': primary_platform,
                    'Primary Giving Platform %': primary_platform_pct,
                    
                    'Contact Channel Status': group['Contact Channel Status'].iloc[0] if 'Contact Channel Status' in group else None,
                    'Current Employer Name': group['Current Employer Name'].iloc[0] if 'Current Employer Name' in group else None,
                    'Most Recent DS Score in Database': group['Most Recent DS Score in Database'].iloc[0] if 'Most Recent DS Score in Database' in group else None,
                    'Most Recent DS Wealth Based Capacity in Database': group['Most Recent DS Wealth Based Capacity in Database'].iloc[0] if 'Most Recent DS Wealth Based Capacity in Database' in group else None,
                    'Current Portfolio Assignment in Database': group['Current Portfolio Assignment in Database'].iloc[0] if 'Current Portfolio Assignment in Database' in group else None,
                    'Longitude': group['Longitude'].iloc[0] if 'Longitude' in group else None,
                    'Latitude': group['Latitude'].iloc[0] if 'Latitude' in group else None
                }
                
                # Calculate ranges and segments
                customer_info.update({
                    'First Gift Date Range A': self.calculate_first_gift_date_range_a(customer_info['First Gift Date']),
                    'First Gift Date Range B': self.calculate_date_range(customer_info['First Gift Date']),
                    'First Gift Amount Range': self.calculate_gift_amount_range(customer_info['First Gift Amount']),
                    'Last Gift Date Range': self.calculate_date_range(customer_info['Last Gift Date']),
                    'Last Gift Amount Range': self.calculate_gift_amount_range(customer_info['Last Gift Amount']),
                    'Largest Gift Date Range': self.calculate_date_range(customer_info['Largest Gift Date']),
                    'Largest Gift Amount Range': self.calculate_gift_amount_range(customer_info['Largest Gift Amount']),
                    'Last Monthly Gift Date Range': self.calculate_date_range(customer_info['Last Monthly Gift Date']),
                    'Last Monthly Giving Amount Range': self.calculate_gift_amount_range(customer_info['Last Monthly Gift Amount']),
                    'Lifetime Giving Range A': self.calculate_gift_amount_range(customer_info['Lifetime Giving']),
                    'Lifetime Giving Range B': self.calculate_gift_amount_range(customer_info['Lifetime Giving']),
                    'Giving Segment A': self.calculate_giving_segment_a(customer_info['Total Number of Gifts']),
                    'Giving Segment B': self.calculate_giving_segment_b(customer_info['Total Number of Gifts'])
                })
                
                # Add lookup dictionary columns
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
            
            # Initialize RFM score columns
            for col in ['Recency Score', 'Frequency Score', 'Monetary Score', 'RFM Score', 
                       'RFM Percentile', 'Recency Percentile', 'Frequency Percentile', 'Monetary Percentile']:
                result[col] = pd.NA
            
            # Only calculate scores for records with valid dates
            valid_mask = result['Recency Criteria'].notna()
            if valid_mask.any():
                result.loc[valid_mask, 'Recency Score'] = scoring_method(result.loc[valid_mask, 'Recency Criteria'], ascending=False)
                result.loc[valid_mask, 'Frequency Score'] = scoring_method(result.loc[valid_mask, 'Frequency Criteria'], ascending=True)
                result.loc[valid_mask, 'Monetary Score'] = scoring_method(result.loc[valid_mask, 'Monetary Criteria'], ascending=True)
                result.loc[valid_mask, 'RFM Score'] = result.loc[valid_mask, ['Recency Score', 'Frequency Score', 'Monetary Score']].sum(axis=1)
            
            return result
            
        except Exception as e:
            self.log(f"Error in rfm_analyzer: {str(e)}")
            self.log(f"Error details: {traceback.format_exc()}")
            raise

if __name__ == "__main__":
    root = tk.Tk()
    app = FinalRFMAnalyzer(root)
    root.mainloop()
