import pandas as pd
import sys
import os

def analyze_excel_file(file_path):
    print(f"\nAnalyzing {file_path}\n")
    
    # Read the entire file without headers first
    df = pd.read_excel(file_path, header=None)
    
    # Extract and print metadata (first 9 rows)
    print("=== METADATA SECTION ===")
    for idx in range(9):
        row = df.iloc[idx]
        if pd.notna(row[0]) and pd.notna(row[1]):  # Only print rows with content
            print(f"{row[0]}: {row[1]}")
    
    # Find header row
    header_row = None
    for idx, row in df.iterrows():
        if pd.notna(row[1]) and str(row[1]).strip() == 'Project':
            header_row = idx
            break

    if header_row is not None:
        print("\n=== COLUMN HEADERS (Row {}) ===".format(header_row))
        headers = df.iloc[header_row]
        for idx, header in enumerate(headers):
            if pd.notna(header) and str(header).strip() != '':
                print(f"{idx}: {header}")
    
        # Read data with proper headers
        data_df = pd.read_excel(file_path, skiprows=header_row+1)
        
        print("\n=== DATA ROWS ===")
        print("First 3 rows:")
        print(data_df.head(3).to_string())
        
        # Find totals section
        print("\n=== TOTALS SECTION ===")
        totals_found = False
        for idx, row in data_df.iterrows():
            if pd.notna(row.iloc[0]):  # Check first column
                first_col = str(row.iloc[0]).strip().lower()
                if 'total' in first_col:
                    totals_found = True
                    print(f"Found totals at row {idx}:")
                    print(row.to_string())
                elif totals_found and pd.notna(row.iloc[0]):
                    print(row.to_string())
    
    print("\n=== FILE STRUCTURE ===")
    print(f"Total rows in file: {len(df)}")
    if header_row is not None:
        print(f"Metadata rows: 0-{header_row-1}")
        print(f"Header row: {header_row}")
        print(f"Data starts at row: {header_row+1}")
    
    # Save to Excel for reference
    output_path = "benevity_analysis.xlsx"
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Raw Data', index=True)
    
    print(f"\nDetailed analysis saved to: {os.path.abspath(output_path)}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python analyze_benevity.py <excel_file_path>")
        sys.exit(1)
        
    try:
        analyze_excel_file(sys.argv[1])
    except Exception as e:
        print(f"Error analyzing file: {str(e)}")
        sys.exit(1)
