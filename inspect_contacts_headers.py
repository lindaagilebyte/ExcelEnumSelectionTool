import pandas as pd
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_contacts = r"c:\Project\ExcelEnumSelectionTool\reference\Form\Contacts.xlsx"

def inspect_headers(path, sheet_name):
    print(f"\n--- Inspecting {os.path.basename(path)} : Sheet '{sheet_name}' Headers ---")
    try:
        # Read first 5 rows to identify header row
        df = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=5)
        df = df.fillna("")
        for i, row in df.iterrows():
            row_list = [str(x) for x in row.tolist()]
            if any(x.strip() for x in row_list):
                 print(f"Row {i+1}: {row_list}")
    except Exception as e:
        print(f"Error: {e}")

inspect_headers(file_contacts, "ContactInfo")
