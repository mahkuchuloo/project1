import sys
import os
import pandas as pd
import json
import time
import tkinter as tk

# Add the parent directory to the sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from rfm_analyzer_helper import RFMAnalyzerHelper

def ensure_output_dir():
    """Create output directory if it doesn't exist"""
    output_dir = os.path.join(os.path.dirname(__file__), 'test_output')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    return output_dir

def save_to_excel(df, output_file, max_retries=3, retry_delay=1):
    """Save DataFrame to Excel with retry logic"""
    for attempt in range(max_retries):
        try:
            df.to_excel(output_file, index=False)
            return True
        except PermissionError:
            if attempt < max_retries - 1:
                print(f"File {output_file} is locked. Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
            else:
                print(f"Failed to save {output_file} after {max_retries} attempts.")
                return False

def run_test(name, df, export_type):
    """Run a single test with specified export type"""
    print(f"\n=== Test: {name} ===")
    
    # Create a root window (hidden)
    root = tk.Tk()
    root.withdraw()
    
    # Initialize the analyzer
    analyzer = RFMAnalyzerHelper(root)
    
    # Set the test data
    analyzer.final_data = df
    
    # Set output paths
    output_dir = ensure_output_dir()
    output_file = os.path.join(output_dir, f"rfm_output_{name.lower().replace(' ', '_')}.xlsx")
    analyzer.output_paths = {export_type: output_file}
    
    # Process the data based on export type
    if export_type == "Export B":
        analyzer.create_export_b()
    elif export_type == "Export F":
        analyzer.create_export_f()
    elif export_type == "Output 1":
        analyzer.create_output_1()
    
    # Read and verify the result
    if os.path.exists(output_file):
        result = pd.read_excel(output_file)
        print(f"Number of records: {len(result)}")
        print("\nColumns included:")
        print("\n".join(result.columns))
        print(f"\nResults saved to: {output_file}")
    else:
        print(f"Failed to create output file: {output_file}")

def test_rfm_analysis():
    """Test RFM analysis with different export types"""
    input_file = r"D:\Projects\CTools\TransactionCompiler\Example Files\generated files\DynamicFinalFile 22122024 - 223516.xlsx"
    
    # Read input data
    df = pd.read_excel(input_file)
    
    # Test Export B
    run_test("Export B", df, "Export B")
    
    # Test Export F
    run_test("Export F", df, "Export F")
    
    # Test Output 1
    run_test("Output 1", df, "Output 1")

if __name__ == "__main__":
    test_rfm_analysis()
