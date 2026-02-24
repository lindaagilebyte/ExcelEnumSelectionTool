# Proposal: Excel Enum Selection Tool (VBA)

## Goal
Provide a surgical, high-speed selection interface for specific "Type" columns in distributed game data files. This tool ensures that categorical strings in data files (e.g., `Contacts.xlsx`) match the definitions in the central reference file (`列舉定義(企劃用).xlsx`).

## Technical Strategy
1. **Surgical Activation:**
   - **Target File:** `Contacts.xlsx` (Sheet: `ContactInfo`).
   - **Trigger:** Hook `Worksheet_SelectionChange`. Check **Row 2** of the active column.
   - **Condition:** If the cell value in Row 2 (e.g., `MissionType`) matches a defined Enum, show the form.
   - **Safety:** Ignore clicks above Row 4 (Data area starts at Row 4).

2. **Multi-Stage Path Resolution:**
   - Check SVN Relative: `..\列舉定義(企劃用).xlsx`
   - Check Local Sandbox: `.\列舉定義(企劃用).xlsx`
   - Manual Fallback: Open Windows File Dialog.

3. **Reference Parsing & Extraction (EnumDef):**
   - **Search Strategy:** scan `EnumDef` sheets for the **Row 2 Key** (e.g., `MissionType`).
   - **Block Identification:** Once the Key is found (e.g., at `A10`), look for the sub-header `定義(巨集顯示)` in the vicinity (e.g., `B11`).
   - **Extraction:** Read the vertical list below `定義(巨集顯示)` until an empty cell is reached.
   - **Structure:** Supports multiple definitions on the same sheet, arranged both side-by-side and vertically.
4. **Performance Optimization:**
   - Use a `Scripting.Dictionary` to store the mapping (Header ID -> Sheet/Cell Location) in memory.
   - Perform a "Deep Scan" only on the first call or manual refresh to eliminate lag on large sheets.
5. **User Interface:**
   - Simple VBA UserForm containing a ListBox for single-click selection.