# Proposal: Excel Enum Selection Tool (VBA)

## Goal
Provide a surgical, high-speed selection interface for specific "Type" columns in distributed game data files. This tool ensures that categorical strings in data files (e.g., `NPC.xlsx`, `Items.xlsx`) match the definitions in the central reference file (`列舉定義(企劃用).xlsx`).

## Technical Strategy
1. **Surgical Activation:**
   - The VBA code and UserForm reside within each individual Data Workbook (e.g., NPC.xlsx) rather than the reference file.
   - Hook `Worksheet_SelectionChange`.
   - Immediate termination if the active cell is above Row 4 or if the Row 2 header of the current column does not exist in the Enum mapping.
2. **Multi-Stage Path Resolution:**
   - Check SVN Relative: `..\列舉定義(企劃用).xlsx`
   - Check Local Sandbox: `.\列舉定義(企劃用).xlsx`
   - Manual Fallback: Open Windows File Dialog if both paths fail.
3. **Reference Parsing & Extraction:**
   - Open reference file as **Read-Only**.
   - Exhaustively scan every sheet in the reference workbook until the target Row 2 Header ID is located.
   - Within the identified block, find the sub-header `定義(巨集顯示)`.
   - Extract the vertical list of strings downward until an empty cell is encountered.
4. **Performance Optimization:**
   - Use a `Scripting.Dictionary` to store the mapping (Header ID -> Sheet/Cell Location) in memory.
   - Perform a "Deep Scan" only on the first call or manual refresh to eliminate lag on large sheets.
5. **User Interface:**
   - Simple VBA UserForm containing a ListBox for single-click selection. 
   - UI closes immediately upon selection.