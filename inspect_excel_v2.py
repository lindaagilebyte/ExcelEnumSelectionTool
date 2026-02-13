import pandas as pd
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_contacts = r"c:\Project\ExcelEnumSelectionTool\reference\Form\Contacts.xlsx"
file_enum = r"c:\Project\ExcelEnumSelectionTool\reference\列舉定義(企劃用).xlsx"

def inspect_sheet(path, sheet_name, rows=15):
    print(f"\n--- Inspecting {os.path.basename(path)} : Sheet '{sheet_name}' ---")
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=rows)
        # Fill NaN with empty string for cleaner output
        df = df.fillna("")
        for i, row in df.iterrows():
            # Convert to list and filter out completely empty trailing cells for display
            row_list = [str(x) for x in row.tolist()]
            # Only print if row isn't effectively empty
            if any(x.strip() for x in row_list):
                 print(f"Row {i+1}: {row_list}")
            else:
                 print(f"Row {i+1}: [EMPTY]")
    except Exception as e:
        print(f"Error: {e}")

# Inspect Contacts sheets
inspect_sheet(file_contacts, "ContactNode")
inspect_sheet(file_contacts, "ContactInfo")

# Inspect EnumDef sheet
inspect_sheet(file_enum, "人脈系統")
