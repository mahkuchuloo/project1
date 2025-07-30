import pandas as pd
from collections import defaultdict

def find_problematic_rel_ids(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Get the first column name (assuming it's the relationship ID column)
    rel_id_col = df.columns[0]
    
    # Create dictionaries to store our findings
    standalone_ids = set()
    concatenated_ids = defaultdict(list)
    
    # First pass: Categorize IDs as standalone or part of concatenated
    for idx, row in df.iterrows():
        rel_id = str(row[rel_id_col])
        if '+' in rel_id:
            # Split concatenated IDs and store them with their original concatenated value
            parts = rel_id.split('+')
            for part in parts:
                concatenated_ids[part.strip()].append(rel_id)
        else:
            standalone_ids.add(rel_id)
    
    # Find problematic cases
    problematic_cases = []
    for standalone_id in standalone_ids:
        if standalone_id in concatenated_ids:
            problematic_cases.append({
                'standalone_id': standalone_id,
                'appears_in_concatenated': concatenated_ids[standalone_id]
            })
    
    # Print findings
    print("\nProblematic Relationship IDs Found:")
    print("===================================")
    for case in problematic_cases:
        print(f"\nStandalone ID: {case['standalone_id']}")
        print("Appears as part of concatenated IDs:")
        for concat_id in case['appears_in_concatenated']:
            # Find rows where these IDs appear
            standalone_rows = df[df[rel_id_col] == case['standalone_id']].index.tolist()
            concatenated_rows = df[df[rel_id_col] == concat_id].index.tolist()
            
            print(f"  - {concat_id}")
            print(f"    * Standalone appears in row(s): {standalone_rows}")
            print(f"    * Concatenated appears in row(s): {concatenated_rows}")

# Usage:
file_path = "D:\\Projects\\CTools\\TransactionCompiler\\Example Files\\generated files\\DynamicFinalFile 09022025 - 233737.xlsx"  # Replace with your file path
find_problematic_rel_ids(file_path)