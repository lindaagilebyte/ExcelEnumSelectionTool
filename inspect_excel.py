import pandas as pd
import os
import sys

# Force UTF-8 output for console
sys.stdout.reconfigure(encoding='utf-8')

files = {
    "Contacts": r"c:\Project\ExcelEnumSelectionTool\reference\Form\Contacts.xlsx",
    "EnumDef": r"c:\Project\ExcelEnumSelectionTool\reference\列舉定義(企劃用).xlsx"
}

def inspect_file(name, path):
    print(f"\n--- Inspecting {name} ---")
    if not os.path.exists(path):
        print(f"Error: File not found at {path}")
        return

    try:
        # Load the workbook to see sheet names
        xl = pd.ExcelFile(path)
        print(f"Sheets: {xl.sheet_names}")

        # Read the first sheet
        df = pd.read_excel(path, sheet_name=0, header=None, nrows=10)
        print(f"First 10 rows of '{xl.sheet_names[0]}':")
        # Print representation to avoid encoding issues hiding structure
        for i, row in df.iterrows():
            clean_row = [str(x) for x in row.dropna().tolist()]
            if clean_row:
                print(f"Row {i+1}: {clean_row}")

    except Exception as e:
        print(f"Error reading {name}: {e}")

for name, path in files.items():
    inspect_file(name, path)
