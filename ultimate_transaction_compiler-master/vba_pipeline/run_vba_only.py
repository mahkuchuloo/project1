import os
import win32com.client
import time

def run_vba_analysis():
    """Run VBA analysis on existing test output files"""
    print("Starting VBA analysis...")
    
    # Get absolute paths to the input files
    print("\nSetting up file paths...")
    output_dir = os.path.abspath(os.path.join('..', 'abstract_rfm', 'test_output'))
    output_b = os.path.abspath(os.path.join(output_dir, 'rfm_output_export_b.xlsx'))
    output_f = os.path.abspath(os.path.join(output_dir, 'rfm_output_export_f.xlsx'))
    output_1 = os.path.abspath(os.path.join(output_dir, 'rfm_output_output_1.xlsx'))
    output_save = os.path.abspath("output")
    
    print("\nInput files:")
    print(f"Output 1: {output_1}")
    print(f"Export B: {output_b}")
    print(f"Export F: {output_f}")
    print(f"Save path: {output_save}")
    
    # Verify input files exist
    print("\nChecking input files...")
    missing_files = []
    for file_path in [output_b, output_f, output_1]:
        if os.path.exists(file_path):
            print(f"Found: {file_path}")
        else:
            print(f"Missing: {file_path}")
            missing_files.append(file_path)
    
    if missing_files:
        print("\nError: Missing required input files:")
        for file in missing_files:
            print(f"- {file}")
        print("\nPlease run test_rfm_helper.py first to generate these files.")
        return
    
    # Create output directory
    print("\nCreating output directory...")
    os.makedirs(output_save, exist_ok=True)
    
    print("\nSetting up Excel...")
    excel = win32com.client.DispatchEx('Excel.Application')
    print("Setting Excel properties...")
    time.sleep(1)  # Short delay to ensure Excel is ready
    excel.Visible = 1  # Use integer instead of boolean
    excel.EnableEvents = 1
    excel.DisplayAlerts = 0
    
    try:
        # Open template with VBA macro
        template_path = os.path.abspath("rfm_template.xlsm")
        if not os.path.exists(template_path):
            print("\nError: Template not found at:", template_path)
            print("Please run create_template.py first")
            return
            
        print("\nOpening template...")
        wb = excel.Workbooks.Open(template_path)
        
        # Set file paths in Excel
        print("\nConfiguring file paths in template:")
        sheet = wb.Sheets("RFM Analyzer")
        print("Setting Output 1 path...")
        sheet.Range("B6").Value = output_1
        print("Setting Export B path...")
        sheet.Range("B9").Value = output_b
        print("Setting Export F path...")
        sheet.Range("B12").Value = output_f
        print("Setting save path...")
        sheet.Range("B15").Value = output_save
        print("Setting output filename...")
        sheet.Range("B18").Value = "RFM_Analysis_Result"
        
        # Run the macro
        print("\nRunning VBA macro...")
        sheet.Activate()
        try:
            excel.Run("CreateOutput1")
            print("VBA macro execution completed")
        except Exception as e:
            print(f"Error running VBA macro: {str(e)}")
            print("This might be because:")
            print("1. The macro security settings in Excel are too restrictive")
            print("2. The VBA code wasn't properly imported into the template")
            print("3. There's an error in the VBA code itself")
            raise
        
        # Check if output was created
        expected_output = os.path.join(output_save, "RFM_Analysis_Result.xlsx")
        if os.path.exists(expected_output):
            print(f"\nSuccess! Output created at: {expected_output}")
            print(f"File size: {os.path.getsize(expected_output):,} bytes")
        else:
            print("\nError: Output file was not created at:", expected_output)
            
    except Exception as e:
        print(f"\nError during execution: {str(e)}")
        raise
    finally:
        try:
            print("\nCleaning up...")
            wb.Close(SaveChanges=False)
            excel.Quit()
        except:
            pass

if __name__ == "__main__":
    run_vba_analysis()
