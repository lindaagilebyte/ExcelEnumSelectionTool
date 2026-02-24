import os
import sys
import win32com.client
import shutil

# --- Configuration ---
SOURCE_DIR = os.path.join(os.path.dirname(__file__), "Source", "DataWorkbook")
VBA_FILES = {
    "Module": "Module_EnumSelector.bas",
    "Form": "Form_EnumSelect.frm",
    "ThisWorkbook": "ThisWorkbook.cls"
}

xlOpenXMLWorkbookMacroEnabled = 52 # Constant for .xlsm

def inject_vba(target_path):
    # 1. Path Validation
    target_path = os.path.abspath(target_path)
    if not os.path.exists(target_path):
        print(f"[Error] Target file not found: {target_path}")
        return False
        
    for key, filename in VBA_FILES.items():
        vba_path = os.path.abspath(os.path.join(SOURCE_DIR, filename))
        if not os.path.exists(vba_path):
            print(f"[Error] Source VBA file not found: {vba_path}")
            return False

    print(f"Injecting into: {target_path}")
    
    # 2. Determine Output Path and Save Strategy
    ext = os.path.splitext(target_path)[1].lower()
    is_xlsm = (ext == ".xlsm")
    
    if is_xlsm:
        output_xlsm_path = target_path
        print(f"Target is already .xlsm. Will modify in-place.")
    else:
        dir_name = os.path.dirname(target_path)
        base_name = os.path.splitext(os.path.basename(target_path))[0]
        output_xlsm_path = os.path.join(dir_name, f"{base_name}_MacroEnabled.xlsm")
        
        # Clean up old output if it exists
        if os.path.exists(output_xlsm_path):
            print(f"Removing old output file: {output_xlsm_path}")
            try:
                os.remove(output_xlsm_path)
            except Exception as e:
                print(f"[Error] Cannot delete old file. Please close Excel. {e}")
                return False
        
    # 3. COM Automation
    print("Starting Excel COM application...")
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.EnableEvents = False # Crucial for opening .xlsm without triggering auto-macros
        
        # Close any auto-recovered or already-open workbooks forcefully
        try:
            for w in excel.Workbooks:
                w.Close(False)
        except:
            pass

        # Open Workbook
        print("Opening target workbook...")
        wb = excel.Workbooks.Open(target_path)
        
        # Ensure VBE Access is trusted
        try:
            vbp = wb.VBProject
        except Exception as e:
            print("[Error] Cannot access VBA Project Object Model.")
            print("Please open Excel, go to File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings")
            print("And check 'Trust access to the VBA project object model'.")
            return False
            
        # 4. Inject Files
        print("Injecting VBA files...")
        
        # Inject standard Module dynamically to avoid encoding bugs in COM Import()
        try:
            module_comp = vbp.VBComponents("Module_EnumSelector")
            print("  + Found existing Module_EnumSelector. Clearing...")
            if module_comp.CodeModule.CountOfLines > 0:
                module_comp.CodeModule.DeleteLines(1, module_comp.CodeModule.CountOfLines)
        except:
            print("  + Creating Module dynamically...")
            module_comp = vbp.VBComponents.Add(1) # 1 = vbext_ct_StdModule
            module_comp.Name = "Module_EnumSelector"
            
        module_path = os.path.abspath(os.path.join(SOURCE_DIR, VBA_FILES["Module"]))
        with open(module_path, 'r', encoding='utf-8') as f:
            module_code_raw = f.read()
            
        # Strip export headers
        module_code_lines = module_code_raw.split('\n')
        module_start_idx = 0
        for i, line in enumerate(module_code_lines):
            if line.startswith('Option Explicit') or line.startswith('Public') or line.startswith('Private'):
                module_start_idx = i
                break
                
        module_clean_code = '\n'.join(module_code_lines[module_start_idx:])
        module_comp.CodeModule.AddFromString(module_clean_code)
        print(f"  + Injected code into {VBA_FILES['Module']}")
        
        # Inject UserForm programmatically to avoid .frx dependency
        try:
            form_comp = vbp.VBComponents("Form_EnumSelect")
            print("  + Found existing Form_EnumSelect. Clearing...")
            if form_comp.CodeModule.CountOfLines > 0:
                form_comp.CodeModule.DeleteLines(1, form_comp.CodeModule.CountOfLines)
            
            # Remove all existing controls
            controls = form_comp.Designer.Controls
            ctrl_names = [c.Name for c in controls]
            for c_name in ctrl_names:
                controls.Remove(c_name)
        except:
            print("  + Creating UserForm dynamically...")
            form_comp = vbp.VBComponents.Add(3) # 3 = vbext_ct_MSForm
            form_comp.Name = "Form_EnumSelect"
            
        form_comp.Properties("Caption").Value = "Select Value"
        form_comp.Properties("Width").Value = 240
        form_comp.Properties("Height").Value = 220
        
        # Create Header Label
        lbl = form_comp.Designer.Controls.Add("Forms.Label.1")
        lbl.Name = "lblHeader"
        lbl.Caption = "Please select a value:"
        lbl.Top = 6
        lbl.Left = 6
        lbl.Width = 222
        lbl.Height = 12
        
        # Create ListBox
        lst = form_comp.Designer.Controls.Add("Forms.ListBox.1")
        lst.Name = "lstEnums"
        lst.Top = 24
        lst.Left = 6
        lst.Width = 222
        lst.Height = 132
        
        # Create Confirm Button
        btnCnf = form_comp.Designer.Controls.Add("Forms.CommandButton.1")
        btnCnf.Name = "btnConfirm"
        btnCnf.Caption = "[確認]"
        btnCnf.Top = 162
        btnCnf.Left = 6
        btnCnf.Width = 100
        btnCnf.Height = 24
        
        # Create Cancel Button
        btnCan = form_comp.Designer.Controls.Add("Forms.CommandButton.1")
        btnCan.Name = "btnCancel"
        btnCan.Caption = "[取消]"
        btnCan.Top = 162
        btnCan.Left = 128
        btnCan.Width = 100
        btnCan.Height = 24
        
        # Create Refresh Button
        btnRef = form_comp.Designer.Controls.Add("Forms.CommandButton.1")
        btnRef.Name = "btnRefresh"
        btnRef.Caption = "Refresh Cache"
        btnRef.Top = 192
        btnRef.Left = 6
        btnRef.Width = 222
        btnRef.Height = 24
        
        # Set Form Properties
        form_comp.Properties("Caption").Value = "Select Value"
        form_comp.Properties("Width").Value = 240
        form_comp.Properties("Height").Value = 250
        
        # Inject UserForm code
        form_path = os.path.abspath(os.path.join(SOURCE_DIR, VBA_FILES["Form"]))
        with open(form_path, 'r', encoding='utf-8') as f:
            form_code_raw = f.read()
            
        form_code_lines = form_code_raw.split('\n')
        form_start_idx = 0
        for i, line in enumerate(form_code_lines):
            if line.startswith('Option Explicit') or line.startswith('Private Sub') or line.startswith('Public Sub'):
                form_start_idx = i
                break
                
        form_clean_code = '\n'.join(form_code_lines[form_start_idx:])
        
        # FIX ENCODING: The system locale is Shift-JIS (mbcs), but our source is UTF-8. 
        # When COM receives the string, it mangles it. 
        # But VBE CodeModule accepts standard strings if we just pass the raw decoded result.
        # Actually, python's win32com handles unicode strings automatically if they are valid.
        # However, the previous Module import failed to decode the file itself.
        
        form_comp.CodeModule.AddFromString(form_clean_code)
        print(f"  + Injected code into Form_EnumSelect")
        
        # Inject ThisWorkbook code
        this_wb_path = os.path.abspath(os.path.join(SOURCE_DIR, VBA_FILES["ThisWorkbook"]))
        with open(this_wb_path, 'r', encoding='utf-8') as f:
            cls_code = f.read()
            
        # We need to strip out the VBA export headers (the first several lines starting with VERSION and Attribute)
        code_lines = cls_code.split('\n')
        start_idx = 0
        for i, line in enumerate(code_lines):
            if line.startswith('Option Explicit') or line.startswith('Private Sub'):
                start_idx = i
                break
                
        clean_code = '\n'.join(code_lines[start_idx:])
        
        # Overwrite the built-in ThisWorkbook component safely
        this_wb_comp = vbp.VBComponents("ThisWorkbook")
        if this_wb_comp.CodeModule.CountOfLines > 0:
            this_wb_comp.CodeModule.DeleteLines(1, this_wb_comp.CodeModule.CountOfLines)
        this_wb_comp.CodeModule.AddFromString(clean_code)
        
        # Save and close immediately before Excel tries to compile it
        print(f"  + Injected code into ThisWorkbook")
        
        # Close VBE window if it opened
        try:
            excel.VBE.MainWindow.Visible = False
        except:
            pass

        # 6. Save Workbook
        if is_xlsm:
            print(f"Saving changes to existing workbook...")
            excel.DisplayAlerts = False
            wb.Save()
            wb.Saved = True
        else:
            print(f"Saving as Macro-Enabled Workbook (.xlsm)...")
            wb.SaveAs(output_xlsm_path, FileFormat=xlOpenXMLWorkbookMacroEnabled)
        print(f"Success! Saved to: {output_xlsm_path}")
        return True
        
    except Exception as e:
        print(f"[Error] Injection failed: {str(e)}")
        return False
    finally:
        # Cleanup COM objects properly
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.EnableEvents = True
                excel.Quit()
            except:
                pass
                
        # Force win32com to garbage collect the COM handles
        wb = None
        excel = None
        
        # Force kill Excel to prevent COM from hanging the python exit
        os.system("taskkill /F /IM excel.exe > nul 2>&1")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python inject_vba.py <path_to_target.xlsx>")
        sys.exit(1)
        
    target = sys.argv[1]
    inject_vba(target)
