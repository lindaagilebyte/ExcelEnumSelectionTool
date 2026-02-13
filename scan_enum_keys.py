import pandas as pd
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_enum = r"c:\Project\ExcelEnumSelectionTool\reference\列舉定義(企劃用).xlsx"

def list_enum_keys(path, sheet_name):
    print(f"\n--- Scanning Keys in {sheet_name} ---")
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=None)
        # Scan valid string cells that might be keys
        # We assume keys are standalone English words or distinct codes
        # and likely sitting above a "定義(巨集顯示)" block.
        
        # Strategy: Find "定義(巨集顯示)" and look around it.
        
        for r in range(len(df)):
            for c in range(len(df.columns)):
                cell_val = str(df.iloc[r, c]).strip()
                if "定義(巨集顯示)" in cell_val:
                    # Look at (r-1, c-1) for Key
                    if r > 0 and c > 0:
                        key_candidate = str(df.iloc[r-1, c-1]).strip()
                        print(f"Found Block at ({r},{c}). Potential Key at ({r-1},{c-1}): '{key_candidate}'")
                    # Also check (r-1, c) just in case
                    if r > 0:
                         key_candidate_2 = str(df.iloc[r-1, c]).strip()
                         print(f"  Alternative Key at ({r-1},{c}): '{key_candidate_2}'")

    except Exception as e:
        print(f"Error: {e}")

list_enum_keys(file_enum, "人脈系統")
