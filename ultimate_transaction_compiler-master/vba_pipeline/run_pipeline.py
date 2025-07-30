import os
import sys
import win32com.client
import shutil
from pathlib import Path

# Add parent directory to path for importing test_rfm_helper
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from abstract_rfm.test_rfm_helper import test_rfm_analysis


def setup_vba_workbook():
    """Open the template workbook with VBA macro"""
    template_path = os.path.abspath("rfm_template.xlsm")
    if not os.path.exists(template_path):
        print(f"Template file not found at: {template_path}")
        print("Please run create_template.py first")
        return None, None, None
    
    # Create Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.EnableEvents = True
    excel.DisplayAlerts = False
    
    # Open template workbook
    wb = excel.Workbooks.Open(template_path)
    
    return excel, wb, template_path

def run_pipeline():
    """Run the complete pipeline"""
    print("Step 1: Running test_rfm_helper to generate outputs...")
    test_rfm_analysis()
    
    print("\nStep 2: Setting up VBA workbook...")
    excel, wb, macro_path = setup_vba_workbook()
    if excel is None:
        print("Failed to set up VBA workbook. Exiting...")
        return
    
    print("\nStep 3: Setting up file paths...")
    # Get absolute paths to the generated files
    output_dir = os.path.abspath(os.path.join('..', 'abstract_rfm', 'test_output'))
    output_b = os.path.abspath(os.path.join(output_dir, 'rfm_output_export_b.xlsx'))
    output_f = os.path.abspath(os.path.join(output_dir, 'rfm_output_export_f.xlsx'))
    output_1 = os.path.abspath(os.path.join(output_dir, 'rfm_output_output_1.xlsx'))
    output_save = os.path.abspath("output")
    
    print(f"Setting file paths:")
    print(f"Output 1: {output_1}")
    print(f"Export B: {output_b}")
    print(f"Export F: {output_f}")
    print(f"Save path: {output_save}")
    
    # Set paths in Excel
    sheet = wb.Sheets("RFM Analyzer")
    sheet.Range("B6").Value = output_1  # Output File 1
    sheet.Range("B9").Value = output_b  # Export B
    sheet.Range("B12").Value = output_f  # Export F
    sheet.Range("B15").Value = output_save  # Save path
    sheet.Range("B18").Value = "RFM_Analysis_Result"  # Output filename
    
    print("\nStep 4: Running VBA macro...")
    try:
        # Make sure the RFM Analyzer sheet is active
        sheet = wb.Sheets("RFM Analyzer")
        sheet.Activate()
        
        # Run the macro
        excel.Run("CreateOutput1")
        print("VBA macro completed successfully!")
    except Exception as e:
        print(f"Error running VBA macro: {e}")
    finally:
        try:
            wb.Close(SaveChanges=True)
            excel.Quit()
        except:
            pass

if __name__ == "__main__":
    # Create output directory if it doesn't exist
    os.makedirs("output", exist_ok=True)
    
    # Run the pipeline
    run_pipeline()
