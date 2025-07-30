import pandas as pd
import json
import os

class DictionaryLookupManager:
    def __init__(self, dictionary_file='lookup_dictionaries.json'):
        self.dictionary_file = dictionary_file
        self.lookups = self.load_dictionaries()

    def load_dictionaries(self):
        """Load dictionaries from JSON file."""
        try:
            with open(self.dictionary_file, 'r') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return []

    def save_dictionaries(self):
        """Save dictionaries to JSON file."""
        with open(self.dictionary_file, 'w') as f:
            json.dump(self.lookups, f)

    def apply_lookup_dictionaries(self, df):
        """Apply all configured lookup dictionaries to the dataframe."""
        for lookup in self.lookups:
            if lookup.get('use_multiple_values', False):
                df = self.apply_multiple_values_lookup(df, lookup)
            elif lookup.get('use_post_merger', False):
                df = self.apply_post_merger_logic(df, lookup)
            elif lookup.get('use_zip_validation', False):
                df = self.apply_zip_validation_logic(df, lookup)
            else:
                df = self.apply_standard_lookup(df, lookup)
        return df

    def _process_empty_dictionary_values(self, value_dict, lookup):
        """Process dictionary values before mapping.
        
        Handles empty values in the dictionary itself:
        - If use_empty_value is True, replaces empty values in the dictionary with empty_value
        
        Args:
            value_dict: Dictionary of lookup values
            lookup: Dictionary containing lookup configuration
        
        Returns:
            Processed dictionary with empty values handled
        """
        if lookup.get('use_empty_value', False):
            empty_value = lookup.get('empty_value', 'EMPTY')
            # Replace empty values in the dictionary itself
            value_dict = {
                k: empty_value if pd.isna(v) or v == '' or v is None 
                else v 
                for k, v in value_dict.items()
            }
        return value_dict

    def _handle_default_values(self, series, lookup):
        """Handle default values for missing mappings.
        
        Args:
            series: pandas.Series containing the mapped values
            lookup: Dictionary containing lookup configuration
        
        Returns:
            pandas.Series with default values applied to missing mappings
        """
        if lookup.get('use_default_value', False):
            default_value = lookup.get('default_value', '')
            series = series.fillna(default_value)
        return series

    def apply_multiple_values_lookup(self, df, lookup):
        """Apply multiple values dictionary lookup."""
        # Read the dictionary file with headers
        dict_df = pd.read_excel(lookup['path'])
        
        # Get the key column (first column) and value columns (all other columns)
        key_column = dict_df.columns[0]
        value_columns = dict_df.columns[1:]
        
        # Create a dictionary for each value column
        for col in value_columns:
            # Create mapping dictionary for this column
            value_dict = dict(zip(dict_df[key_column], dict_df[col]))
            
            # Process dictionary values to handle empty values
            value_dict = self._process_empty_dictionary_values(value_dict, lookup)
            
            # Apply mapping and handle default values
            mapped_values = df[lookup['lookup_column']].map(value_dict)
            df[col] = self._handle_default_values(mapped_values, lookup)
        
        return df

    def apply_standard_lookup(self, df, lookup):
        """Apply standard dictionary lookup."""
        # Read dictionary without headers for simple key-value mapping
        dict_df = pd.read_excel(lookup['path'], header=None, names=['key', 'value'])
        lookup_dict = dict(zip(dict_df['key'], dict_df['value']))
        
        # Process dictionary values to handle empty values
        lookup_dict = self._process_empty_dictionary_values(lookup_dict, lookup)
        
        # Apply mapping and handle default values
        mapped_values = df[lookup['lookup_column']].map(lookup_dict)
        df[lookup['output_column']] = self._handle_default_values(mapped_values, lookup)
        
        return df

    def apply_post_merger_logic(self, df, lookup):
        """Apply post-merger logic to the lookup."""
        # Create dictionaries for the lookup values
        value_dict = {value['key']: value['value'] for value in lookup['values']}
        value_dict.update({value['merger_key']: value['value'] 
                          for value in lookup['values'] if value['merger_key']})
        
        clean_name_dict = {value['key']: value.get('clean_name', value['key']) 
                          for value in lookup['values']}
        clean_merger_name_dict = {value['merger_key']: value.get('clean_merger_name', value['merger_key']) 
                                 for value in lookup['values'] if value['merger_key']}
        
        # Process dictionary values to handle empty values
        value_dict = self._process_empty_dictionary_values(value_dict, lookup)
        
        # Apply mapping and handle default values
        mapped_values = df[lookup['lookup_column']].map(value_dict)
        df[lookup['output_column']] = self._handle_default_values(mapped_values, lookup)
        
        # Apply the clean name logic to the lookup column
        df[lookup['lookup_column']] = df[lookup['lookup_column']].apply(
            lambda x: clean_name_dict.get(x, clean_merger_name_dict.get(x, x))
        )
        
        return df

    def apply_zip_validation_logic(self, df, lookup):
        """Apply zip code validation logic to the lookup."""
        dict_df = pd.read_excel(lookup['path'], header=None, names=['key', 'value'])
        lookup_dict = dict(zip(dict_df['key'], dict_df['value']))
        
        # Process dictionary values to handle empty values
        lookup_dict = self._process_empty_dictionary_values(lookup_dict, lookup)
        
        # Process zip codes and apply lookup with default value handling
        processed_zips = df[lookup['lookup_column']].apply(self._process_zip_code)
        mapped_values = processed_zips.map(lookup_dict)
        df[lookup['output_column']] = self._handle_default_values(mapped_values, lookup)
        
        return df

    def _process_zip_code(self, zip_code):
        """Process zip code for validation."""
        if pd.isna(zip_code):
            return None
        
        if isinstance(zip_code, float):
            zip_code = str(int(zip_code))
        else:
            zip_code = str(zip_code).strip()

        # Handle zip+4 format (e.g., 33613-7716)
        if '-' in zip_code:
            zip_code = zip_code.split('-')[0]

        # Remove leading zeros
        zip_code = zip_code.lstrip('0')

        # Ensure it's a number
        if zip_code.isdigit():
            return int(zip_code)
        else:
            return None

    def get_last_gift_columns(self, df, lookups=None):
        """Get last gift values for specified columns."""
        if lookups is None:
            lookups = self.lookups

        last_gift_columns = [
            'Transaction ID', 'Date Clean', 'Amount', 'Giving Platform', 
            'Gift Range Chart', 'Gift Segment', 'Is Recurring'
        ]
        
        # Add lookup dictionary columns that should be included in last gift values
        for lookup in lookups:
            if lookup.get('include_in_last_gift', False):
                if lookup.get('use_multiple_values', False):
                    # For multiple values dictionaries, include all value columns
                    dict_df = pd.read_excel(lookup['path'])
                    value_columns = dict_df.columns[1:]  # All columns except the first
                    last_gift_columns.extend(value_columns)
                else:
                    # For standard dictionaries, include the output column
                    last_gift_columns.append(lookup['output_column'])

        result_df = df.copy()
        for col in last_gift_columns:
            if col in df.columns:
                result_df[f'Last Gift {col}'] = df.groupby('Relationship ID')[col].shift()

        return result_df
