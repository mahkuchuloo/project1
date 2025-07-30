import sys
import os
import pandas as pd
import json
import time

# Add the parent directory to the sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from abstract_rfm.final_rfm_analyzer import FinalRFMAnalyzer
from dictionary_lookup_manager import DictionaryLookupManager

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

def run_test(name, df, output_selections):
    """Run a single test with specified output selections"""
    print(f"\n=== Test: {name} ===")
    
    # Create analyzer instance
    analyzer = FinalRFMAnalyzer(None)
    
    # Configure outputs
    for key in analyzer.output_selections:
        analyzer.output_selections[key] = key in output_selections
    
    # Process the data
    result = analyzer.rfm_analyzer(df)
    
    # Verify results
    print(f"Number of records: {len(result)}")
    print("\nColumns included:")
    print("\n".join(result.columns))
    
    # Save results
    output_dir = ensure_output_dir()
    output_file = os.path.join(output_dir, f"rfm_output_{name.lower().replace(' ', '_')}.xlsx")
    
    if save_to_excel(result, output_file):
        print(f"\nResults saved to: {output_file}")
    else:
        print(f"\nFailed to save results to: {output_file}")
        print("Please close any open Excel files and try again.")

def test_rfm_analysis():
    """Test RFM analysis with different output combinations"""
    input_file = r"D:\Projects\CTools\TransactionCompiler\Example Files\generated files\DynamicFinalFile 22122024 - 223516.xlsx"
    
    # Read input data
    df = pd.read_excel(input_file)
    
    # Load dictionary configurations
    with open('rfm_lookup_dictionaries.json', 'r') as f:
        dictionaries = json.load(f)
        dict_names = [d['name'] for d in dictionaries]
    
    # Test 1: Basic info and RFM scores only
    run_test("Basic RFM", df, [
        'Basic Info',
        'RFM Scores',
        'Date Ranges',
        'Amount Ranges'
    ])
    
    # Test 2: Gift information focus
    run_test("Gift Info", df, [
        'Basic Info',
        'First Gift Info',
        'Last Gift Info',
        'Largest Gift Info',
        'Campaign Info',
        'Appeal Info',
        'Date Ranges',
        'Amount Ranges'
    ])
    
    # Test 3: Donor profile focus
    run_test("Donor Profile", df, [
        'Basic Info',
        'Contact Info',
        'Geographic Info',
        'Employer Info',
        'Portfolio Info',
        'DS Scores'
    ])
    
    # Test 4: Giving patterns focus
    run_test("Giving Patterns", df, [
        'Basic Info',
        'Monthly Gift Info',
        'Giving Segments',
        'Platform Info',
        'Contact Channel',
        'Date Ranges',
        'Amount Ranges'
    ])
    
    # Test 5: Minimal output
    run_test("Minimal", df, [
        'Basic Info',
        'Amount Ranges'
    ])
    
    # Test 6: MSA Dictionary only
    run_test("MSA Only", df, [
        'Basic Info',
        'MSA Dictionary'
    ])
    
    # Test 7: All dictionaries
    run_test("All Dictionaries", df, [
        'Basic Info'
    ] + dict_names)
    
    # Test 8: Full analysis with all dictionaries
    run_test("Full Analysis", df, [
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
    ] + dict_names)

if __name__ == "__main__":
    test_rfm_analysis()
