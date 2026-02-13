# Design: Excel Enum Selection Tool

## Context
A VBA-based tool embedded in Excel files to assist data entry by providing valid enum options from a reference file.

## Technical Decisions

### 1. Data Structure (Caching)
To avoid reading the reference file on every click, we will use a global `Scripting.Dictionary`:
- **Key:** EnumHeader (e.g., "MissionType")
- **Value:** `Collection` or `Array` of valid strings.
- **Cache Invalidation:** The cache is built on the first click. A "Refresh" button on the UserForm can force a rebuild.

### 2. Path Resolution
The tool needs to find `列舉定義(企劃用).xlsx`.
Strategy:
1. `Target = ThisWorkbook.Path & "\..\列舉定義(企劃用).xlsx"` (SVN relative)
2. `Target = ThisWorkbook.Path & "\列舉定義(企劃用).xlsx"` (Same dir)
3. If both fail, prompt user to select file -> Store path in a hidden Name or Registry (optional, for now just prompt per session).

### 3. UserForm Design
- **Controls:**
  - `lstEnums` (ListBox): Shows options.
  - `lblHeader` (Label): Shows current Enum Key.
  - `btnRefresh` (Button): Reloads reference file.
- **Behavior:**
  - `lstEnums_Click`: Writes value to ActiveCell and Unloads form.
  - `UserForm_Initialize`: Populates list from Cache.

### 4. Modules
- `ThisWorkbook`: Handles `SheetSelectionChange`.
- `Module_EnumSelector`:
  - `Public Sub OnSelectionChange(Target As Range)`: Main entry point.
  - `Private Function LoadEnumDef(Key As String) As Variant`: Loads data.
