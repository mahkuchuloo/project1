import sys
import os
import traceback
import json

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
from abstract_rfm.output_selection_dialog import OutputSelectionDialog

class FinalRFMAnalyzer(BaseToolFrame):
    def __init__(self, master=None):
        # Initialize without UI if master is None
        if master is not None:
            super().__init__(master, "Final RFM Analyzer")
            self.is_ui_mode = True
        else:
            self.is_ui_mode = False
            self.progress_queue = None
            self.log_queue = None
        
        self.dict_manager = DictionaryLookupManager('rfm_lookup_dictionaries.json')
        self.final_data = None
        self.output_path = None
        self.input_file_path = None
        self.dictionary_cache = {}
        
        # Initialize output selections
        self.output_selections = {}
        
        # Core output selections
        core_selections = [
            'Basic Info',
            'RFM Scores',
            'First Gift Info',
            'Last Gift Info',
            'Largest Gift Info',
            'Monthly Gift Info',
            'Giving Segments',
            'Platform Info',
            'Contact Info',
            'Geographic Info',
            'Campaign Info',
            'Appeal Info',
            'Employer Info',
            'Portfolio Info',
            'DS Scores',
            'Contact Channel',
            'Date Ranges',
            'Amount Ranges'
        ]
        
        # Add core selections
        for selection in core_selections:
            if master is not None:
                self.output_selections[selection] = tk.BooleanVar(value=True)
            else:
                self.output_selections[selection] = True
        
        # Add dictionary selections
        with open('rfm_lookup_dictionaries.json', 'r') as f:
            dictionaries = json.load(f)
            for dictionary in dictionaries:
                name = dictionary['name']
                if master is not None:
                    self.output_selections[name] = tk.BooleanVar(value=True)
                else:
                    self.output_selections[name] = True
        
        # Add scoring method selection
        self.scoring_methods = {
            "Percentile (Original VBA)": RFMScorer.percentile_scoring,
            "Quartile": RFMScorer.quartile_scoring,
            "Equal Width": RFMScorer.equal_width_scoring,
            "Z-Score": RFMScorer.zscore_scoring,
            "Logarithmic": RFMScorer.logarithmic_scoring,
            "Threshold": self.threshold_scoring
        }
        
        # Default scoring method
        self.scoring_method_value = "Percentile (Original VBA)"
        self.threshold_values = "100,500,1000,5000,10000"
        
        # Only create UI elements if master is not None
        if master is not None:
            # Add scoring method dropdown
            scoring_frame = tk.Frame(self)
            scoring_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(scoring_frame, text="RFM Scoring Method:").pack(side=tk.LEFT)
            self.scoring_method = ttk.Combobox(scoring_frame, 
                                             values=list(self.scoring_methods.keys()),
                                             state="readonly")
            self.scoring_method.set(self.scoring_method_value)
            self.scoring_method.pack(side=tk.LEFT, padx=5)
            
            # Add output selection button
            output_button = ttk.Button(scoring_frame, text="Select Output Columns...", command=self.show_output_dialog)
            output_button.pack(side=tk.LEFT, padx=5)
            
            # Add threshold inputs
            self.threshold_frame = tk.Frame(self)
            tk.Label(self.threshold_frame, text="Custom Thresholds (comma-separated):").pack(side=tk.LEFT)
            self.threshold_entry = tk.Entry(self.threshold_frame)
            self.threshold_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            self.threshold_entry.insert(0, self.threshold_values)
            
            # Initially hide threshold frame
            if self.scoring_method.get() != "Threshold":
                self.threshold_frame.pack_forget()
                
            # Bind the selection event
            self.scoring_method.bind('<<ComboboxSelected>>', self.on_scoring_method_change)

    def show_output_dialog(self):
        """Show the output selection dialog"""
        OutputSelectionDialog(self, self.output_selections)
    
    def on_scoring_method_change(self, event):
        if self.scoring_method.get() == "Threshold":
            self.threshold_frame.pack(fill=tk.X, padx=5, pady=5)
        else:
            self.threshold_frame.pack_forget()

    def get_scoring_method(self):
        """Get current scoring method, handling both UI and non-UI cases"""
        if hasattr(self, 'scoring_method'):
            return self.scoring_method.get()
        return self.scoring_method_value
    
    def get_thresholds(self):
        """Get threshold values, handling both UI and non-UI cases"""
        if hasattr(self, 'threshold_entry'):
            return self.threshold_entry.get()
        return self.threshold_values
    
    def threshold_scoring(self, series, ascending=True):
        try:
            thresholds = [float(x.strip()) for x in self.get_thresholds().split(',')]
            return RFMScorer.threshold_scoring(series, thresholds, ascending)
        except ValueError as e:
            if hasattr(self, 'threshold_entry'):
                messagebox.showerror("Error", "Invalid threshold values. Please enter comma-separated numbers.")
            raise e

    def get_dictionary_df(self, path):
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
            (1000000, float('inf'), 'R. $1,000,000+')
        ]
        
        if pd.isna(amount):
            return 'No Amount'
            
        for low, high, label in ranges:
            if low <= amount < high:
                return label
        return 'R. $1,000,000+'

    def calculate_date_range_a(self, date):
        """Calculate date range category A"""
        if pd.isna(date):
            return "No Date"
            
        now = pd.Timestamp.now()
        days_diff = (now - date).days
        
        if days_diff < 0:
            return "Future Date"
        elif days_diff == 0:
            return "Today"
        elif days_diff <= 7:
            return "Last Week"
        elif days_diff <= 30:
            return "Last Month"
        elif days_diff <= 90:
            return "Last Quarter"
        elif days_diff <= 365:
            return "Last Year"
        else:
            return "Over a Year"

    def calculate_date_range_b(self, date):
        """Calculate date range category B"""
        if pd.isna(date):
            return "No Date"
            
        now = pd.Timestamp.now()
        days_diff = (now - date).days
        
        if days_diff < 0:
            return "Future Date"
        elif days_diff <= 30:
            return "Last Month"
        elif days_diff <= 90:
            return "Last Quarter"
        elif days_diff <= 180:
            return "Last 6 Months"
        elif days_diff <= 365:
            return "Last Year"
        elif days_diff <= 730:
            return "Last 2 Years"
        else:
            return "Over 2 Years"

    def calculate_giving_segment_a(self, total_gifts):
        if total_gifts == 0:
            return "A. No Gifts"
        elif total_gifts == 1:
            return "B. First Gift"
        elif total_gifts == 2:
            return "C. Second Gift"
        elif total_gifts == 3:
            return "D. Third Gift"
        elif total_gifts == 4:
            return "E. Fourth Gift"
        elif total_gifts == 5:
            return "F. Fifth Gift"
        elif total_gifts > 5:
            return "G. More than 5 Gifts"
        return "H. None"

    def calculate_giving_segment_b(self, total_gifts):
        if total_gifts < 13:
            return "A. Less than 13"
        elif 13 <= total_gifts <= 17:
            return "B. 13-17 gifts"
        elif 18 <= total_gifts <= 22:
            return "C. 18-22 gifts"
        elif 23 <= total_gifts <= 27:
            return "D. 23-27 gifts"
        elif 28 <= total_gifts <= 32:
            return "E. 28-32 gifts"
        elif 33 <= total_gifts <= 36:
            return "F. 33-36 gifts"
        elif total_gifts >= 37:
            return "G. 37+ gifts"
        return "H. None"

    def calculate_percentile_category(self, value, series):
        try:
            ranks = series.rank(pct=True) * 100
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
            return "F. Bottom 50%"

    def calculate_percentile_category_vectorized(self, series):
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

    def log(self, message):
        """Override log method to handle non-UI mode"""
        if self.is_ui_mode:
            super().log(message)
        else:
            print(message)
    
    def progress_queue_put(self, value):
        """Handle progress updates in non-UI mode"""
        if self.is_ui_mode:
            self.progress_queue.put(value)
        else:
            print(f"Progress: {value}%")
    
    def process_data(self):
        """Process data in a separate thread."""
        try:
            self.log("Starting RFM analysis...")
            self.progress_queue_put(10)

            # Read the input file
            df = pd.read_excel(self.input_file_path)
            self.progress_queue_put(20)

            # Process the data
            result_df = self.rfm_analyzer(df)
            self.progress_queue_put(90)

            # Save the results
            output_filename = f"rfm_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            self.output_path = os.path.join(os.path.dirname(self.input_file_path), output_filename)
            result_df.to_excel(self.output_path, index=False)
            
            self.progress_queue_put(100)
            self.log(f"Analysis complete. Results saved to: {self.output_path}")
            
            # Update UI in the main thread
            self.master.after(0, self.update_ui_on_completion)

        except Exception as e:
            self.log(f"Error during processing: {str(e)}")
            self.log(f"Error details: {traceback.format_exc()}")
            # Show error in the main thread
            self.master.after(0, lambda: messagebox.showerror("Error", f"An error occurred during processing: {str(e)}"))
            raise

    def update_ui_on_completion(self):
        """Update UI elements after processing is complete."""
        self.result_label.config(text="Analysis complete!")
        self.file_link.config(text="Click here to open output file location")
        self.process_button.config(state=tk.NORMAL)

    def start_processing(self):
        """Implement the required start_processing method from BaseToolFrame."""
        if not hasattr(self, 'input_file_path') or not self.input_file_path:
            messagebox.showerror("Error", "Please select an input file first.")
            return

        # Disable the process button while processing
        self.process_button.config(state=tk.DISABLED)
        
        # Start processing in a separate thread
        processing_thread = threading.Thread(target=self.process_data)
        processing_thread.daemon = True  # Thread will be terminated when main program exits
        processing_thread.start()

    def rfm_analyzer(self, df):
        try:
            self.log("Starting RFM analysis...")
            
            df['Date Clean'] = pd.to_datetime(df['Date Clean'])
            recurring_ids = set(df[df['Recurring ID'].notna()]['Relationship ID'].unique())
            grouped = df.groupby('Relationship ID')
            
            rfm_data = []
            total_groups = len(grouped)
            
            for i, (name, group) in enumerate(grouped):
                recency = group['Date Clean'].max()
                frequency = len(group['Transaction ID'].unique())
                monetary = group['Amount'].sum()
                
                first_gift_idx = group['Date Clean'].idxmin()
                last_gift_idx = group['Date Clean'].idxmax()
                largest_gift_idx = group['Amount'].idxmax()
                
                monthly_gifts = group[group['Recurring ID'].notna()]
                last_monthly_gift = None if monthly_gifts.empty else monthly_gifts.loc[monthly_gifts['Date Clean'].idxmax()]
                
                platform_counts = group['Giving Platform'].value_counts()
                primary_platform = platform_counts.index[0] if not platform_counts.empty else None
                primary_platform_pct = (platform_counts.iloc[0] / len(group) * 100) if not platform_counts.empty else 0
                
                customer_info = {}
                
                # Basic Info
                if isinstance(self.output_selections['Basic Info'], tk.BooleanVar):
                    include_basic = self.output_selections['Basic Info'].get()
                else:
                    include_basic = self.output_selections['Basic Info']
                if include_basic:
                    customer_info.update({
                        'Relationship VAN ID': name,
                        'Total Number of Gifts': frequency,
                        'Lifetime Giving': monetary,
                        'Lifetime Giving Sum': monetary  # Additional column from final_rfm_analyzer
                    })
                
                # RFM Scores
                if isinstance(self.output_selections['RFM Scores'], tk.BooleanVar):
                    include_rfm = self.output_selections['RFM Scores'].get()
                else:
                    include_rfm = self.output_selections['RFM Scores']
                if include_rfm:
                    customer_info.update({
                        'Recency Criteria': recency,
                        'Frequency Criteria': frequency,
                        'Monetary Criteria': monetary
                    })
                
                # First Gift Info
                if isinstance(self.output_selections['First Gift Info'], tk.BooleanVar):
                    include_first = self.output_selections['First Gift Info'].get()
                else:
                    include_first = self.output_selections['First Gift Info']
                if include_first:
                    customer_info.update({
                        'First Gift Date': group.loc[first_gift_idx, 'Date Clean'],
                        'First Gift Amount': group.loc[first_gift_idx, 'Amount'],
                        'First Gift Platform': group.loc[first_gift_idx, 'Giving Platform']
                    })
                
                # Campaign Info
                if isinstance(self.output_selections['Campaign Info'], tk.BooleanVar):
                    include_campaign = self.output_selections['Campaign Info'].get()
                else:
                    include_campaign = self.output_selections['Campaign Info']
                if include_campaign:
                    customer_info.update({
                        'First Gift Campaign Name': group.loc[first_gift_idx, 'Campaign Name'] if 'Campaign Name' in group else None,
                        'Last Gift Campaign Name': group.loc[last_gift_idx, 'Campaign Name'] if 'Campaign Name' in group else None,
                        'Largest Gift Campaign Name': group.loc[largest_gift_idx, 'Campaign Name'] if 'Campaign Name' in group else None
                    })
                
                # Appeal Info
                if isinstance(self.output_selections['Appeal Info'], tk.BooleanVar):
                    include_appeal = self.output_selections['Appeal Info'].get()
                else:
                    include_appeal = self.output_selections['Appeal Info']
                if include_appeal:
                    customer_info.update({
                        'First Gift Appeal Name': group.loc[first_gift_idx, 'Appeal Name'] if 'Appeal Name' in group else None,
                        'Last Gift Appeal Name': group.loc[last_gift_idx, 'Appeal Name'] if 'Appeal Name' in group else None,
                        'Largest Gift Appeal Name': group.loc[largest_gift_idx, 'Appeal Name'] if 'Appeal Name' in group else None
                    })
                
                # Last Gift Info
                if isinstance(self.output_selections['Last Gift Info'], tk.BooleanVar):
                    include_last = self.output_selections['Last Gift Info'].get()
                else:
                    include_last = self.output_selections['Last Gift Info']
                if include_last:
                    customer_info.update({
                        'Last Gift Date': recency,
                        'Last Gift Amount': group.loc[last_gift_idx, 'Amount'],
                        'Last Gift Platform': group.loc[last_gift_idx, 'Giving Platform']
                    })
                
                # Largest Gift Info
                if isinstance(self.output_selections['Largest Gift Info'], tk.BooleanVar):
                    include_largest = self.output_selections['Largest Gift Info'].get()
                else:
                    include_largest = self.output_selections['Largest Gift Info']
                if include_largest:
                    customer_info.update({
                        'Largest Gift Date': group.loc[largest_gift_idx, 'Date Clean'],
                        'Largest Gift Amount': group.loc[largest_gift_idx, 'Amount'],
                        'Largest Gift Platform': group.loc[largest_gift_idx, 'Giving Platform']
                    })
                
                # Monthly Gift Info
                if isinstance(self.output_selections['Monthly Gift Info'], tk.BooleanVar):
                    include_monthly = self.output_selections['Monthly Gift Info'].get()
                else:
                    include_monthly = self.output_selections['Monthly Gift Info']
                if include_monthly:
                    customer_info.update({
                        'Last Monthly Gift Date': last_monthly_gift['Date Clean'] if last_monthly_gift is not None else None,
                        'Last Monthly Gift Amount': last_monthly_gift['Amount'] if last_monthly_gift is not None else None,
                        'Digital Monthly Indicator': 'Digital Monthly' if name in recurring_ids else 'Not Digital Monthly'
                    })
                
                # Giving Segments
                if isinstance(self.output_selections['Giving Segments'], tk.BooleanVar):
                    include_segments = self.output_selections['Giving Segments'].get()
                else:
                    include_segments = self.output_selections['Giving Segments']
                if include_segments:
                    customer_info.update({
                        'Giving Segment A': self.calculate_giving_segment_a(frequency),
                        'Giving Segment B': self.calculate_giving_segment_b(frequency)
                    })
                
                # Platform Info
                if isinstance(self.output_selections['Platform Info'], tk.BooleanVar):
                    include_platform = self.output_selections['Platform Info'].get()
                else:
                    include_platform = self.output_selections['Platform Info']
                if include_platform:
                    customer_info.update({
                        'Primary Giving Platform': primary_platform,
                        'Primary Giving Platform %': primary_platform_pct
                    })
                
                # Contact Info
                if isinstance(self.output_selections['Contact Info'], tk.BooleanVar):
                    include_contact = self.output_selections['Contact Info'].get()
                else:
                    include_contact = self.output_selections['Contact Info']
                if include_contact:
                    customer_info.update({
                        'Email': group['Donor Email'].iloc[0] if 'Donor Email' in group else None,
                        'Phone': group['Donor Phone'].iloc[0] if 'Donor Phone' in group else None,
                        'Address': group['Donor Address Line 1'].iloc[0] if 'Donor Address Line 1' in group else None
                    })
                
                # Geographic Info
                if isinstance(self.output_selections['Geographic Info'], tk.BooleanVar):
                    include_geo = self.output_selections['Geographic Info'].get()
                else:
                    include_geo = self.output_selections['Geographic Info']
                if include_geo:
                    customer_info.update({
                        'City': group['Donor City'].iloc[0] if 'Donor City' in group else None,
                        'State': group['Donor State'].iloc[0] if 'Donor State' in group else None,
                        'Zip': group['Donor ZIP'].iloc[0] if 'Donor ZIP' in group else None,
                        'Longitude': group['Longitude'].iloc[0] if 'Longitude' in group else None,
                        'Latitude': group['Latitude'].iloc[0] if 'Latitude' in group else None
                    })
                
                # Employer Info
                if isinstance(self.output_selections['Employer Info'], tk.BooleanVar):
                    include_employer = self.output_selections['Employer Info'].get()
                else:
                    include_employer = self.output_selections['Employer Info']
                if include_employer:
                    customer_info.update({
                        'Current Employer Name': group['Current Employer Name'].iloc[0] if 'Current Employer Name' in group else None
                    })
                
                # Portfolio Info
                if isinstance(self.output_selections['Portfolio Info'], tk.BooleanVar):
                    include_portfolio = self.output_selections['Portfolio Info'].get()
                else:
                    include_portfolio = self.output_selections['Portfolio Info']
                if include_portfolio:
                    customer_info.update({
                        'Current Portfolio Assignment in Database': group['Current Portfolio Assignment in Database'].iloc[0] if 'Current Portfolio Assignment in Database' in group else None
                    })
                
                # DS Scores
                if isinstance(self.output_selections['DS Scores'], tk.BooleanVar):
                    include_ds = self.output_selections['DS Scores'].get()
                else:
                    include_ds = self.output_selections['DS Scores']
                if include_ds:
                    customer_info.update({
                        'Most Recent DS Score in Database': group['Most Recent DS Score in Database'].iloc[0] if 'Most Recent DS Score in Database' in group else None,
                        'Most Recent DS Wealth Based Capacity in Database': group['Most Recent DS Wealth Based Capacity in Database'].iloc[0] if 'Most Recent DS Wealth Based Capacity in Database' in group else None
                    })
                
                # Contact Channel
                if isinstance(self.output_selections['Contact Channel'], tk.BooleanVar):
                    include_channel = self.output_selections['Contact Channel'].get()
                else:
                    include_channel = self.output_selections['Contact Channel']
                if include_channel:
                    customer_info.update({
                        'Contact Channel Status': group['Contact Channel Status'].iloc[0] if 'Contact Channel Status' in group else None
                    })
                
                # Date Ranges
                if isinstance(self.output_selections['Date Ranges'], tk.BooleanVar):
                    include_dates = self.output_selections['Date Ranges'].get()
                else:
                    include_dates = self.output_selections['Date Ranges']
                if include_dates:
                    if 'First Gift Date' in customer_info:
                        customer_info['First Gift Date Range A'] = self.calculate_date_range_a(customer_info['First Gift Date'])
                        customer_info['First Gift Date Range B'] = self.calculate_date_range_b(customer_info['First Gift Date'])
                    if 'Last Gift Date' in customer_info:
                        customer_info['Last Gift Date Range'] = self.calculate_date_range_a(customer_info['Last Gift Date'])
                    if 'Largest Gift Date' in customer_info:
                        customer_info['Largest Gift Date Range'] = self.calculate_date_range_a(customer_info['Largest Gift Date'])
                    if 'Last Monthly Gift Date' in customer_info:
                        customer_info['Last Monthly Gift Date Range'] = self.calculate_date_range_a(customer_info['Last Monthly Gift Date'])
                
                # Amount Ranges
                if isinstance(self.output_selections['Amount Ranges'], tk.BooleanVar):
                    include_amounts = self.output_selections['Amount Ranges'].get()
                else:
                    include_amounts = self.output_selections['Amount Ranges']
                if include_amounts:
                    if 'First Gift Amount' in customer_info:
                        customer_info['First Gift Amount Range'] = self.calculate_gift_amount_range(customer_info['First Gift Amount'])
                    if 'Last Gift Amount' in customer_info:
                        customer_info['Last Gift Amount Range'] = self.calculate_gift_amount_range(customer_info['Last Gift Amount'])
                    if 'Largest Gift Amount' in customer_info:
                        customer_info['Largest Gift Amount Range'] = self.calculate_gift_amount_range(customer_info['Largest Gift Amount'])
                    if 'Last Monthly Gift Amount' in customer_info:
                        customer_info['Last Monthly Giving Amount Range'] = self.calculate_gift_amount_range(customer_info['Last Monthly Gift Amount'])
                    if 'Lifetime Giving' in customer_info:
                        customer_info['Lifetime Giving Range A'] = self.calculate_gift_amount_range(customer_info['Lifetime Giving'])
                        customer_info['Lifetime Giving Range B'] = self.calculate_gift_amount_range(customer_info['Lifetime Giving'])
                
                # Individual Dictionary Lookups
                for lookup in self.dict_manager.lookups:
                    dict_name = lookup['name']
                    if isinstance(self.output_selections[dict_name], tk.BooleanVar):
                        include_dict = self.output_selections[dict_name].get()
                    else:
                        include_dict = self.output_selections[dict_name]
                    
                    if include_dict:
                        try:
                            if dict_name == "MSA Dictionary":
                                customer_info["MSA"] = None  # Default value
                                zip_code = group['Donor ZIP'].iloc[0] if 'Donor ZIP' in group.columns else None
                                if zip_code:
                                    try:
                                        dict_df = self.get_dictionary_df(lookup['path'])
                                        if lookup['lookup_column'] in dict_df.columns:
                                            matches = dict_df[dict_df[lookup['lookup_column']] == zip_code]
                                            if not matches.empty and lookup['output_column'] in matches.columns:
                                                customer_info["MSA"] = matches[lookup['output_column']].iloc[0]
                                    except Exception as e:
                                        self.log(f"Error in MSA lookup: {str(e)}")
                            elif dict_name == "Split Dictionary":
                                customer_info["Split?"] = lookup.get('default_value', None)  # Default value
                                url = group['Contribution Form URL'].iloc[0] if 'Contribution Form URL' in group.columns else None
                                if url:
                                    try:
                                        dict_df = self.get_dictionary_df(lookup['path'])
                                        if lookup['lookup_column'] in dict_df.columns:
                                            matches = dict_df[dict_df[lookup['lookup_column']] == url]
                                            if not matches.empty and lookup['output_column'] in matches.columns:
                                                customer_info["Split?"] = matches[lookup['output_column']].iloc[0]
                                    except Exception as e:
                                        self.log(f"Error in Split lookup: {str(e)}")
                            elif dict_name == "Membership Dictionary":
                                # Set default value based on configuration
                                if lookup.get('use_default_value', False):
                                    customer_info["Member?"] = lookup.get('default_value', None)
                                elif lookup.get('use_empty_value', False):
                                    customer_info["Member?"] = lookup.get('empty_value', None)
                                else:
                                    customer_info["Member?"] = None
                                
                                rel_id = name  # Using the group name which is the Relationship ID
                                try:
                                    dict_df = self.get_dictionary_df(lookup['path'])
                                    if lookup['lookup_column'] in dict_df.columns:
                                        matches = dict_df[dict_df[lookup['lookup_column']] == rel_id]
                                        if not matches.empty and lookup['output_column'] in matches.columns:
                                            customer_info["Member?"] = matches[lookup['output_column']].iloc[0]
                                except Exception as e:
                                    self.log(f"Error in Member lookup: {str(e)}")
                            elif lookup.get('use_multiple_values', False):
                                try:
                                    dict_df = self.get_dictionary_df(lookup['path'])
                                    value_columns = dict_df.columns[1:]
                                    
                                    if lookup.get('include_in_last_gift', False):
                                        for col in value_columns:
                                            if col in group.columns:
                                                customer_info[f"Last Gift {col}"] = group.loc[last_gift_idx, col]
                                    else:
                                        for col in value_columns:
                                            if col in group.columns:
                                                customer_info[col] = group.iloc[0][col]
                                except Exception as e:
                                    self.log(f"Error in multiple values lookup: {str(e)}")
                            else:
                                try:
                                    output_column = lookup['output_column']
                                    if output_column and output_column in group.columns:
                                        if lookup.get('include_in_last_gift', False):
                                            customer_info[f"Last Gift {output_column}"] = group.loc[last_gift_idx, output_column]
                                        else:
                                            customer_info[output_column] = group[output_column].iloc[0]
                                except Exception as e:
                                    self.log(f"Error in single value lookup: {str(e)}")
                        except Exception as e:
                            self.log(f"Error in dictionary lookup for {dict_name}: {str(e)}")
                
                rfm_data.append(customer_info)
                
                if i % 100 == 0 or i == total_groups - 1:
                    progress = 30 + int((i / total_groups) * 60)
                    self.progress_queue_put(progress)
            
            result = pd.DataFrame(rfm_data)
            
            if isinstance(self.output_selections['RFM Scores'], tk.BooleanVar):
                include_rfm_scores = self.output_selections['RFM Scores'].get()
            else:
                include_rfm_scores = self.output_selections['RFM Scores']
            if include_rfm_scores:
                scoring_method = self.scoring_methods[self.get_scoring_method()]
                
                result['Recency Score'] = scoring_method(result['Recency Criteria'], ascending=False)
                result['Frequency Score'] = scoring_method(result['Frequency Criteria'], ascending=True)
                result['Monetary Score'] = scoring_method(result['Monetary Criteria'], ascending=True)
                result['RFM Score'] = result['Recency Score'] + result['Frequency Score'] + result['Monetary Score']
                
                result['RFM Percentile'] = self.calculate_percentile_category_vectorized(result['RFM Score'])
                result['Recency Percentile'] = self.calculate_percentile_category_vectorized(result['Recency Score'])
                result['Frequency Percentile'] = self.calculate_percentile_category_vectorized(result['Frequency Score'])
                result['Monetary Percentile'] = self.calculate_percentile_category_vectorized(result['Monetary Score'])
            
            return result
            
        except Exception as e:
            self.log(f"Error in rfm_analyzer: {str(e)}")
            self.log(f"Error details: {traceback.format_exc()}")
            raise

if __name__ == "__main__":
    root = tk.Tk()
    app = FinalRFMAnalyzer(root)
    root.mainloop()
