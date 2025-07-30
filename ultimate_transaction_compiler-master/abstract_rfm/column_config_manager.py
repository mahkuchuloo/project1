import json
import os

class ColumnConfigManager:
    def __init__(self, config_file='column_config.json'):
        self.config_file = config_file
        self.config_path = os.path.join(os.path.dirname(__file__), config_file)
        self.selected_columns = []
        self.load_config()

    def load_config(self):
        """Load column configuration from file"""
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r') as f:
                    config = json.load(f)
                    self.selected_columns = config.get('selected_columns', [])
            except Exception as e:
                print(f"Error loading column configuration: {str(e)}")
                self.selected_columns = []
        else:
            # Initialize with default columns
            self.selected_columns = [
                "Recency Criteria",
                "Frequency Criteria",
                "Monetary Criteria",
                "RFM Score",
                "RFM Percentile",
                "Recency Score",
                "Frequency Score",
                "Monetary Score",
                "Recency Percentile",
                "Frequency Percentile",
                "Monetary Percentile"
            ]
            self.save_config()

    def save_config(self):
        """Save column configuration to file"""
        config = {
            'selected_columns': self.selected_columns
        }
        try:
            with open(self.config_path, 'w') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            print(f"Error saving column configuration: {str(e)}")

    def set_columns(self, columns):
        """Set selected columns and save configuration"""
        self.selected_columns = columns
        self.save_config()

    def get_columns(self):
        """Get selected columns"""
        return self.selected_columns
