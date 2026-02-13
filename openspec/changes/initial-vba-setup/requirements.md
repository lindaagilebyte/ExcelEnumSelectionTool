# Requirements: Excel Enum Selection Tool

## Functional Requirements
- Distributed Deployment: The tool's VBA modules must be present in each Data Workbook to monitor and process that specific file’s data surgical triggers.
- **Targeted Triggering:** The tool must only activate for columns where Row 2 contains a valid Enum ID defined in the reference file. All numerical and general text columns must be ignored.
- **Traditional Chinese UI:** All UserForm labels and system messages must be in Traditional Chinese for coworker accessibility.
- **Vertical List Extraction:** Correctly parse lists under the `定義(巨集顯示)` sub-header, supporting side-by-side data blocks on the same sheet.
- **Data Integrity:** Prevent typos by forcing selection from the defined list for relevant columns.

## Defensive Programming & Quality
- **Non-Disruptive Feedback:** Use the Excel **Status Bar** for loading/scanning updates to avoid interruptive pop-up windows.
- **Explicit Read-Only:** Ensure the reference workbook is never locked for other SVN users.
- **No Silent Failures:** Display a clear message if a Row 2 ID exists but the corresponding `定義(巨集顯示)` sub-header or data block is missing.
- **Sanitization:** Trim leading/trailing whitespace from all IDs and enum strings.

## Performance
- **Scalability:** Must handle a reference file where enum definitions are distributed across any number of sheets and data files with potentially 100+ sheets without causing UI lag.
- **Memory Management:** Clear the `Scripting.Dictionary` on workbook close to release resources.
- **Debug Logging:** Include a `CONST_DEBUG_MODE` to print file paths and search results to the VBA Immediate Window.