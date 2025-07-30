import pandas as pd
import logging
from abc import ABC, abstractmethod

class Platform(ABC):
    def __init__(self, name, file_pattern, date_field, amount_field, id_field, secondary_id_field, is_base_platform, relationship_id_key, has_display_name=False, date_fallback_field=None):
        self.sample_columns = []  # Store available columns from sample file
        self.sample_file_path = None  # Store path to sample file
        self.recurring_true_value = None  # Store value that indicates recurring donation
        self._recurring_values = []  # Store available recurring values
        self.name = name
        self.file_pattern = file_pattern
        self.date_field = date_field
        self.date_fallback_field = date_fallback_field
        self.amount_field = amount_field
        self.id_field = id_field
        self.secondary_id_field = secondary_id_field
        self._is_base_platform = is_base_platform
        self.column_mapping = {}
        self.relationship_id_key = relationship_id_key

    def process_data(self, df):
        logging.info(f"Processing data for platform: {self.name}")
        logging.info(f"Is base platform: {self._is_base_platform}")
        logging.info(f"Relationship ID key: {self.relationship_id_key}")

        # Handle date fallback if configured
        if self.date_fallback_field and self.date_fallback_field in df.columns:
            logging.info(f"Checking date fallback field: {self.date_fallback_field}")
            mask = pd.isna(df[self.date_field]) | (df[self.date_field] == '')
            df.loc[mask, self.date_field] = df.loc[mask, self.date_fallback_field]
            logging.info(f"Applied date fallback for {mask.sum()} rows")

        # Set Date Clean directly from processed date field
        try:
            df['Date Clean'] = pd.to_datetime(df[self.date_field]).dt.date
            logging.info("Successfully set Date Clean field")
        except Exception as e:
            logging.error(f"Error setting Date Clean field: {str(e)}")
            df['Date Clean'] = ''

        # Copy amount field to Amount column
        try:
            df['Amount'] = df[self.amount_field]
            logging.info("Successfully set Amount field")
        except Exception as e:
            logging.error(f"Error setting Amount field: {str(e)}")
            df['Amount'] = ''

        df['Giving Platform'] = self.get_platform_name()

        # Initialize has_display_name if not set
        if not hasattr(self, 'has_display_name'):
            self.has_display_name = False
            
        # Initialize Display Name
        df['Display Name'] = ''
        
        # Only initialize Recurring ID if it doesn't exist or is empty
        if 'Recurring ID' not in df.columns or df['Recurring ID'].isna().all():
            df['Recurring ID'] = ''
            
        # Initialize Is Recurring if it doesn't exist with proper boolean value
        if 'Is Recurring' not in df.columns:
            df['Is Recurring'] = False

        for target_col, mapping in self.column_mapping.items():
            source_col = mapping['target']
            default_value = mapping['default']
            logging.info(f"Mapping column: Target = {target_col}, Source = {source_col}")
            
            try:
                if target_col == 'Recurring ID':
                    if source_col != 'N/A':
                        if source_col in df.columns:
                            df[target_col] = df[source_col]
                        else:
                            logging.warning(f"Source column {source_col} not found for Recurring ID. Using default value \"{default_value}\"")
                            df[target_col] = default_value
                else:
                    if source_col != 'N/A':
                        if source_col in df.columns:
                            df[target_col] = df[source_col]
                            logging.info(f"Successfully mapped {source_col} to {target_col}")
                        else:
                            logging.warning(f"Source column {source_col} not found in DataFrame. Using default value \"{default_value}\" for {target_col}")
                            df[target_col] = default_value
                    elif target_col != 'Recurring ID':  # Don't set default for Recurring ID when source is N/A
                        if target_col == 'Is Recurring':
                            if self.recurring_true_value and source_col in df.columns:
                                # Use configured recurring true value if available
                                df[target_col] = df[source_col].astype(str) == self.recurring_true_value
                                logging.info(f"Set Is Recurring based on value '{self.recurring_true_value}'")
                            else:
                                # Fall back to default TRUE/FALSE behavior
                                bool_value = default_value.upper() == 'TRUE'
                                df[target_col] = bool_value
                                logging.info(f"Set Is Recurring to boolean value {bool_value}")
                        else:
                            df[target_col] = default_value
                            logging.info(f"Using default value \"{default_value}\" for {target_col}")
            except Exception as e:
                logging.error(f"Error processing column {target_col}: {str(e)}")
                logging.info(f"Continuing to next column")
                continue

        if self._is_base_platform:
            duplicate_col = self.get_duplicate_column_name()
            if duplicate_col in df.columns:
                df = df[df[duplicate_col] != 'Duplicate']
                logging.info(f"Filtered out duplicate entries using column: {duplicate_col}")
            else:
                logging.warning(f"Duplicate column {duplicate_col} not found. Skipping duplicate filtering.")
        
        # Store important columns that need to be preserved
        recurring_ids = df['Recurring ID'].copy()

        # Handle Match? column based on Reason column if it exists
        df['Match?'] = 'Not a Match'  # Default value
        if 'Reason' in df.columns:
            df.loc[df['Reason'] == 'Match', 'Match?'] = 'Match'

        # Handle Display Name logic after all column mappings are done
        if not self.has_display_name:
            logging.info("Processing display name logic")
            
            # Get mapped donor name fields (these will be in the final output)
            first_name = df.get('Donor First Name', pd.Series(''))
            last_name = df.get('Donor Last Name', pd.Series(''))
            
            # Create initial display name from first/last name
            df['Display Name'] = first_name.fillna('').str.strip() + ' ' + last_name.fillna('').str.strip()
            df['Display Name'] = df['Display Name'].str.strip()
            
            # Identify rows where display name is empty or just whitespace
            empty_display_mask = (df['Display Name'].isna()) | (df['Display Name'] == '')
            logging.debug(f"Found {empty_display_mask.sum()} rows with empty display names")
            
            # Try to get contact name from source data (before column mapping)
            # Check various possible column names for contact name
            contact_name_columns = ['Contact Name', 'ContactName', 'contact_name', 'CONTACT NAME', 'Primary Contact']
            source_contact_name = None
            
            for col in contact_name_columns:
                if col in df.columns:
                    source_contact_name = df[col]
                    logging.info(f"Found contact name in source column: {col}")
                    break
            
            if source_contact_name is not None:
                # Use contact name as fallback where available
                has_contact_mask = source_contact_name.notna() & (source_contact_name != '')
                fallback_mask = empty_display_mask & has_contact_mask
                
                if fallback_mask.any():
                    logging.info(f"Using contact name as fallback for {fallback_mask.sum()} rows")
                    df.loc[fallback_mask, 'Display Name'] = source_contact_name[fallback_mask].str.strip()
            
            # Log final results
            final_empty_mask = (df['Display Name'].isna()) | (df['Display Name'] == '')
            logging.info(f"Final count of empty display names: {final_empty_mask.sum()}")

        # Restore Recurring ID after display name processing
        df['Recurring ID'] = recurring_ids

        logging.info(f"Finished processing data for platform: {self.name}")
        return df

    def get_unique_transaction_key(self, row):
        return f"{row[self.id_field]}{int(pd.to_datetime(row[self.date_field]).date().strftime('%Y%m%d'))}{row[self.amount_field]}"

    def get_relationship_id_key(self, row):
        return str(row[self.relationship_id_key]).lower() if pd.notnull(row[self.relationship_id_key]) else ''

    def get_platform_name(self):
        return self.name

    def get_date_field(self):
        return self.date_field

    def get_amount_field(self):
        return self.amount_field

    def get_id_field(self):
        return self.id_field

    def get_secondary_id_field(self):
        return self.secondary_id_field

    def is_base_platform(self):
        return self._is_base_platform

    def get_duplicate_column_name(self):
        return f'Duplicate Platform {self.name}'

    def update_recurring_values(self, df=None):
        """Update recurring values from dataframe or sample file"""
        self._recurring_values = []
        
        if df is None and self.sample_file_path:
            try:
                df = pd.read_excel(self.sample_file_path)
            except Exception as e:
                logging.error(f"Error reading sample file: {str(e)}")
                return

        if df is not None and isinstance(df, pd.DataFrame):
            # Get the mapped 'Is Recurring' column name from column_mapping
            is_recurring_mapping = self.column_mapping.get('Is Recurring', {})
            source_col = is_recurring_mapping.get('target') if is_recurring_mapping else None
            
            if source_col and source_col != 'N/A' and source_col in df.columns:
                # Get unique non-null values
                mask = df[source_col].notna()
                if mask.any():
                    values = df.loc[mask, source_col].astype(str).unique()
                    # Sort and take top 10 non-empty values that aren't 'nan'
                    values = sorted([v for v in values if v and v.lower() != 'nan'])[:10]
                    self._recurring_values = values

    def get_recurring_values(self):
        """Get cached recurring values"""
        return self._recurring_values

    def to_dict(self):
        return {
            'name': self.name,
            'file_pattern': self.file_pattern,
            'date_field': self.date_field,
            'date_fallback_field': self.date_fallback_field,
            'amount_field': self.amount_field,
            'id_field': self.id_field,
            'secondary_id_field': self.secondary_id_field,
            'is_base_platform': self._is_base_platform,
            'column_mapping': self.column_mapping,
            'relationship_id_key': self.relationship_id_key,
            'has_display_name': getattr(self, 'has_display_name', False),
            'sample_columns': self.sample_columns,
            'sample_file_path': self.sample_file_path,
            'recurring_true_value': self.recurring_true_value,
            'recurring_values': self._recurring_values
        }

    @classmethod
    def from_dict(cls, data):
        platform = cls(
            data['name'],
            data['file_pattern'],
            data['date_field'],
            data['amount_field'],
            data['id_field'],
            data['secondary_id_field'],
            data['is_base_platform'],
            data['relationship_id_key'],
            data.get('has_display_name', False),  # Optional field with default False
            data.get('date_fallback_field')  # Optional field
        )
        platform.column_mapping = data['column_mapping']
        platform.sample_columns = data.get('sample_columns', [])
        platform.sample_file_path = data.get('sample_file_path')
        platform.recurring_true_value = data.get('recurring_true_value')
        platform._recurring_values = data.get('recurring_values', [])
        return platform
