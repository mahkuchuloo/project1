import win32com.client
import os

def create_template():
    """Create Excel template with VBA macro"""
    print("Starting template creation...")
    
    # Create Excel application
    print("Launching Excel...")
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        print("Setting Excel properties...")
        # Set properties after a short delay to ensure Excel is ready
        import time
        time.sleep(1)
        excel.Visible = 1  # Use integer instead of boolean
        excel.DisplayAlerts = 0
        # Create new workbook
        print("Creating new workbook...")
        wb = excel.Workbooks.Add()
        
        # Rename first sheet
        print("Setting up RFM Analyzer sheet...")
        wb.Sheets(1).Name = "RFM Analyzer"
        
        # Read VBA code
        print("Reading VBA code...")
        vba_path = os.path.abspath('simplified_macro.vba')
        print(f"VBA file path: {vba_path}")
        if not os.path.exists(vba_path):
            raise FileNotFoundError(f"VBA file not found at: {vba_path}")
            
        with open(vba_path, 'r') as f:
            vba_code = f.read()
        
        # Add VBA module
        print("Adding VBA module...")
        try:
            vba_module = wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            vba_module.CodeModule.AddFromString(vba_code)
        except Exception as e:
            print("Error adding VBA module. Please ensure Excel macro security settings allow programmatic access to VBA project")
            print("To fix this:")
            print("1. Open Excel")
            print("2. Go to File > Options > Trust Center > Trust Center Settings > Macro Settings")
            print("3. Enable 'Trust access to the VBA project object model'")
            raise e
        
        # Save as macro-enabled workbook
        print("Saving template...")
        template_path = os.path.abspath("rfm_template.xlsm")
        wb.SaveAs(template_path, FileFormat=52)  # 52 = xlOpenXMLWorkbookMacroEnabled
        print(f"Template created successfully at: {template_path}")
        
    except Exception as e:
        print(f"Error creating template: {str(e)}")
        raise
    finally:
        try:
            print("Cleaning up...")
            wb.Close(SaveChanges=False)
            excel.Quit()
        except:
            pass

if __name__ == "__main__":
    create_template()
